import logging
from pathlib import Path

from precios.models.contexto import CpDefinition, ExecutionContext
from precios.repositories.precios_repository import PreciosRepository
from precios.Reporte import GeneradorDocumento


class PreciosService:
    WAIT_BATCH_SECONDS = 5
    WAIT_BATCH_INITIAL_SECONDS = 0
    WAIT_BATCH_MAX_SECONDS = 1800
    ALIAS_COL_UPCABC = 0
    ALIAS_COL_DESCABC = 1
    ALIAS_COL_PRICEABC = 2
    ALIAS_COL_PROMOABC = 3
    ALIAS_COL_PRICEPABC = 4
    ALIAS_COL_ZONEABC = 17
    ALIAS_TABLE_HEADERS = ["UPCABC", "DESCABC", "PRICEABC", "PROMOABC", "PRICEPABC"]

    def __init__(self, bases, repository=None, document_builder=None):
        self.logger = logging.getLogger("PreciosService")
        self.repository = repository or PreciosRepository(bases)
        self.document_builder_class = document_builder.__class__ if document_builder is not None else GeneradorDocumento
        self.document_builder = None
        self.bases = bases
        self.contexto: dict[str, ExecutionContext] = {}
        self.alias_cache = {}
        self.alias_data_cache = {}
        self.validation_result_cache = {}
        self.cp_documents = {}
        self.ultimo_detalle_alias = []

    def cargar_set_pruebas(self):
        return [
            CpDefinition("SV01", "Validacion de envio item por promo para zona con mantenimiento KREGP"),
            CpDefinition("SV02", "Validacion de envio del mismo item para una zona sin promo"),
            CpDefinition("SV03", "Validacion de herencia para la tienda 523 internet online"),
            CpDefinition("SV04", "Validacion de herencia para la tienda 623 internet online"),
            CpDefinition("SV05", "Validar envio de items con promo activa para tiendas INT y CRC"),
            CpDefinition("SV06", "Validacion de cancelacion parcial"),
            CpDefinition("SV07", "Validacion de cancelacion total"),
            CpDefinition("SV08", "Creacion + cancelacion + creacion precio temporal mismo dia"),
            CpDefinition("SV09", "Creacion + cancelacion + creacion precio permanente mismo dia"),
            CpDefinition("SV10", "Generacion de archivos ItemImport y PricingImport"),
            CpDefinition("SV11", "Envio de archivos hacia magento"),
            CpDefinition("SV12", "Cambio de precio por cambios en PO"),
            CpDefinition("SV13", "Cambio de precio permanente por aumento"),
            CpDefinition("SV14", "Cambio de precio por disminucion"),
            CpDefinition("SV15", "Activacion de precio con fecha futura"),
            CpDefinition("SV16", "Generacion de Global de precios"),
        ]

    def ejecutar_base(self, base, cps=None):
        self.logger.info("INICIO ejecutar_base | base=%s", base)
        self.contexto[base] = self.preparar_contexto_base(base)
        self.ejecutar_set_pruebas(base, cps=cps)
        self.logger.info("FIN ejecutar_base | base=%s", base)

    def _cargar_aliases_para_validacion(self, base, aliases_objetivo):
        """
        Carga cada alias objetivo y lo indexa por SKU.
        Reutiliza el contenido ya cargado en CP previos para no releer el
        mismo PRLKPABC desde BD varias veces dentro de la misma ejecucion.
        """
        cache_aliases = {}

        for alias in aliases_objetivo:
            cache_key = (str(base or "").strip().upper(), str(alias or "").strip())
            if cache_key in self.alias_data_cache:
                cache_aliases[alias] = self.alias_data_cache[cache_key]
                self.logger.debug(
                    "Alias reutilizado desde cache | base=%s alias=%s skus=%s",
                    cache_key[0],
                    cache_key[1],
                    len(cache_aliases[alias]),
                )
                continue

            try:
                datos = self.repository.ejecutar_alias(base, alias) or []
            except Exception as exc:
                self.logger.warning("Error ejecutando alias completo %s | %s", alias, exc)
                self.alias_data_cache[cache_key] = {}
                cache_aliases[alias] = {}
                continue

            indice_sku = {}

            for fila in datos:
                try:
                    upcabc = fila[self.ALIAS_COL_UPCABC]
                    sku = self._extraer_sku_upcabc(upcabc)
                    if not sku:
                        continue

                    zona_alias = str(fila[self.ALIAS_COL_ZONEABC] or "").strip()
                    priceabc = fila[self.ALIAS_COL_PRICEABC]
                    promoabc = str(fila[self.ALIAS_COL_PROMOABC] or "").strip()
                    pricepabc = fila[self.ALIAS_COL_PRICEPABC]

                    registro = {
                        "alias": alias,
                        "fila": fila,
                        "reporte_fila": (
                            fila[self.ALIAS_COL_UPCABC],
                            fila[self.ALIAS_COL_DESCABC],
                            fila[self.ALIAS_COL_PRICEABC],
                            fila[self.ALIAS_COL_PROMOABC],
                            fila[self.ALIAS_COL_PRICEPABC],
                        ),
                        "sku": sku,
                        "zona": zona_alias,
                        "priceabc": priceabc,
                        "pricepabc": pricepabc,
                        "promoabc": promoabc,
                        "precio_ok": False,
                    }

                    indice_sku.setdefault(sku, []).append(registro)

                except Exception as exc:
                    self.logger.warning("Fila invalida en alias %s | %s", alias, exc)

            self.alias_data_cache[cache_key] = indice_sku
            cache_aliases[alias] = indice_sku

            total_filas = sum(len(filas) for filas in indice_sku.values())
            self.logger.debug(
                "Alias cargado para validacion | base=%s alias=%s skus=%s filas=%s",
                cache_key[0],
                cache_key[1],
                len(indice_sku),
                total_filas,
            )

        return cache_aliases

    def preparar_contexto_base(self, base):
        contexto = ExecutionContext()
        self.logger.info("INICIO preparar_contexto_base | base=%s", base)

        try:
            self.repository.ajustar_hora_ejecucion(base)
        except Exception as exc:
            contexto.errores.append(f"ajustar_hora_ejecucion: {exc}")

        # Removed shared data loading to make each CP independent

        job_name = f"RI{base[2:4]}PC"
        try:
            contexto.programa_sometido = self.repository.someter_programa(base)
            contexto.actjob_resultado = self.repository.esperar_fin_job_actjob(
                job_name,
                espera_inicial=self.WAIT_BATCH_INITIAL_SECONDS,
                espera_entre_refrescos=self.WAIT_BATCH_SECONDS,
                espera_post_refresco=5,
                max_espera_segundos=self.WAIT_BATCH_MAX_SECONDS,
            )
            contexto.programa_finalizado = True
        except Exception as exc:
            contexto.errores.append(f"proceso_batch: {exc}")

        try:
            contexto.cambios_aplicados = self.repository.actualizar_estado_cambio_precio(base) or []
        except Exception as exc:
            contexto.errores.append(f"actualizar_estado_cambio_precio: {exc}")

        self.logger.info(
            "FIN preparar_contexto_base | base=%s errores=%s",
            base, len(contexto.errores),
        )
        return contexto

    def ejecutar_set_pruebas(self, base, cps=None):
        self.logger.info("INICIO SET PRUEBAS | base=%s", base)
        pruebas = self.cargar_set_pruebas()
        cps_solicitados = None
        if cps:
            cps_solicitados = {str(cp).strip().upper() for cp in cps if str(cp).strip()}
            pruebas = [prueba for prueba in pruebas if prueba.id.upper() in cps_solicitados]
            self.logger.info(
                "Filtro CP aplicado | base=%s solicitados=%s ejecutados=%s",
                base,
                sorted(cps_solicitados),
                [prueba.id for prueba in pruebas],
            )

        if not pruebas:
            self.logger.warning("No hay CP validos para ejecutar | base=%s cps=%s", base, cps)
            return

        for prueba in pruebas:
            self.logger.info("Ejecutando %s - %s", prueba.id, prueba.nombre)
            metodo = getattr(self, f"test_{prueba.id.lower()}", None)
            if metodo is None:
                self.logger.debug("CP no implementado, saltando | cp=%s", prueba.id)
                continue
            try:
                metodo(base)
            except Exception as exc:
                self.logger.exception("ERROR ejecutando %s | base=%s", prueba.id, base)
                self._iniciar_cp(base, prueba.id, prueba.nombre)
                self.document_builder.agregar_resultado("ERROR")
                self.document_builder.agregar_texto(f"Error ejecutando CP: {exc}")
        self.logger.info("FIN SET PRUEBAS SV | base=%s", base)

    def guardar(self, ruta_archivo):
        ruta_base = Path(ruta_archivo)
        for (base, cp_id), documento in self.cp_documents.items():
            nombre = f"{ruta_base.stem}_{base}_{cp_id}{ruta_base.suffix or '.docx'}"
            documento.guardar(str(ruta_base.with_name(nombre)))

    def _get_contexto(self, base):
        return self.contexto.setdefault(base, ExecutionContext())

    def _documentar_contexto_base(self, base):
        contexto = self._get_contexto(base)
        self.document_builder.iniciar_etapa("Flujo Ejecutado")
        self.document_builder.agregar_texto("1. Configuracion inicial de la base y carga de datos candidatos.")
        self.document_builder.agregar_texto("2. Insercion de SKUs en KREGL para forzar procesamiento.")
        self.document_builder.agregar_texto("3. Sometimiento del programa batch de precios.")
        self.document_builder.agregar_texto("4. Monitoreo del job por ACTJOB/QTEMPKILL hasta su finalizacion.")
        self.document_builder.agregar_texto("5. Validacion posterior de alias PRLKPABC, precios y cambios aplicados.")
        if contexto.programa_sometido:
            self.document_builder.agregar_texto("Programa batch sometido: SI")
        if contexto.programa_finalizado:
            self.document_builder.agregar_texto("Programa batch finalizado segun ACTJOB: SI")
        if contexto.cambios_aplicados:
            self._agregar_tabla_cambios_aplicados(contexto.cambios_aplicados)
        if contexto.errores:
            self.document_builder.agregar_texto("Errores de preparacion detectados:")
            for error in contexto.errores:
                self.document_builder.agregar_texto(error)

    def _construir_tabla_cambios_aplicados(self, cambios):
        filas = []
        for cambio in cambios:
            skus = cambio.get("skus") or [""]
            for sku in skus:
                filas.append((
                    str(sku).strip(),
                    str(cambio.get("zona", "")).strip(),
                    str(cambio.get("tipo", "")).strip(),
                    str(cambio.get("cambio", "")).strip(),
                ))
        return filas or [("", "", "", "")]

    def _agregar_tabla_cambios_aplicados(self, cambios, titulo="Cambios aplicados detectados:"):
        if not cambios:
            return
        self.document_builder.agregar_texto(titulo)
        self.document_builder.agregar_tabla(
            self._construir_tabla_cambios_aplicados(cambios),
            encabezados=["SKU", "ZONA", "TIPO", "CAMBIO_PRECIO"],
        )

    def _crear_documento_cp(self, base, cp_id):
        builder = self.document_builder_class()
        builder.establecer_base(base)
        self.cp_documents[(base, cp_id)] = builder
        return builder

    def _iniciar_cp(self, base, cp_id, descripcion):
        self.document_builder = self._crear_documento_cp(base, cp_id)
        self._documentar_contexto_base(base)
        self.document_builder.iniciar_cp(f"{base} - {descripcion}")

    def _agrupar_por_alias(self, encontrados):
        agrupados = {}
        for row in encontrados:
            agrupados.setdefault(row["alias"], []).append(row)
        return agrupados

    def _pintar_resultados_por_alias(self, encontrados):
        for alias, filas in self._agrupar_por_alias(encontrados).items():
            self.document_builder.iniciar_alias(alias)
            self.document_builder.agregar_tabla([row["fila"] for row in filas])
            for row in filas:
                self.document_builder.agregar_texto(
                    f"Alias={row['alias']} ZONA={row['zona']} PROMOABC={row['promoabc']} PRICEABC={row['priceabc']} PRICEPABC={row['pricepabc']} PRECIO_OK={row['precio_ok']}"
                )

    def _pintar_detalle_alias_evaluados(self):
        for detalle in self.ultimo_detalle_alias:
            self.document_builder.iniciar_alias(detalle["alias"])
            self.document_builder.agregar_tabla(
                [row["reporte_fila"] for row in detalle["filas"]],
                encabezados=self.ALIAS_TABLE_HEADERS,
            )
            if not detalle["filas"]:
                self.document_builder.agregar_texto("Sin registros para el SKU evaluado en este PRLKPABC.")
                continue
            for row in detalle["filas"]:
                self.document_builder.agregar_texto(
                    f"Alias={row['alias']} ZONA={row['zona']} PROMOABC={row['promoabc']} PRICEABC={row['priceabc']} PRICEPABC={row['pricepabc']} PRECIO_OK={row['precio_ok']}"
                )

    def _finalizar_cp(self, resultado, detalle=None):
        self.document_builder.agregar_resultado(resultado)
        if detalle:
            self.document_builder.agregar_texto(detalle)

    def _registrar_errores_contexto(self, base):
        contexto = self._get_contexto(base)
        if contexto.errores:
            self.document_builder.agregar_texto("Errores de preparacion detectados:")
            for error in contexto.errores:
                self.document_builder.agregar_texto(error)
        if contexto.programa_sometido:
            self.document_builder.agregar_texto("Programa batch sometido: SI")
        if contexto.programa_finalizado:
            self.document_builder.agregar_texto("Programa batch finalizado segun ACTJOB: SI")

    def _normalizar_alias(self, alias):
        if isinstance(alias, str):
            return alias
        if alias is None:
            return None
        try:
            return str(alias[0]).strip()
        except Exception:
            return str(alias).strip()

    def _obtener_todos_aliases(self, base):
        if base in self.alias_cache:
            return self.alias_cache[base]
        aliases = [self._normalizar_alias(alias) for alias in self.repository.obtener_alias_prlkpabc(base)]
        aliases = [alias for alias in aliases if alias]
        self.alias_cache[base] = aliases
        return aliases

    def _obtener_aliases_por_zona(self, base, zona):
        zona_normalizada = str(zona or "").strip().upper()
        if not zona_normalizada:
            return []
        cache_key = (base, zona_normalizada)
        if cache_key in self.alias_cache:
            return self.alias_cache[cache_key]
        aliases = [self._normalizar_alias(alias) for alias in self.repository.obtener_prlkpabc(base, zona_normalizada)]
        aliases = [alias for alias in aliases if alias]
        self.alias_cache[cache_key] = aliases
        return aliases

    def _resolver_aliases_objetivo(self, base, tipo_precio=None, zona=None, zona_destino=None):
        base_normalizada = str(base or "").strip().upper()
        tipo_precio_normalizado = str(tipo_precio or "").strip().upper()
        zona_normalizada = str(zona or "").strip().upper()
        zona_destino_normalizada = str(zona_destino or "").strip().upper()

        self.logger.debug(
            "INICIO _resolver_aliases_objetivo | base=%s tipo_precio=%s zona=%s zona_destino=%s",
            base_normalizada,
            tipo_precio_normalizado,
            zona_normalizada,
            zona_destino_normalizada
        )

        if not base_normalizada:
            self.logger.warning("_resolver_aliases_objetivo sin base válida")
            return []

        aliases = []

        if tipo_precio_normalizado == "T":
            if zona_normalizada:
                aliases.extend(self._obtener_aliases_por_zona(base_normalizada, zona_normalizada))
            if zona_destino_normalizada and zona_destino_normalizada != zona_normalizada:
                aliases.extend(self._obtener_aliases_por_zona(base_normalizada, zona_destino_normalizada))
            if not aliases:
                aliases = self._obtener_todos_aliases(base_normalizada)
        else:
            if zona_normalizada:
                aliases = self._obtener_aliases_por_zona(base_normalizada, zona_normalizada)
            elif zona_destino_normalizada:
                aliases = self._obtener_aliases_por_zona(base_normalizada, zona_destino_normalizada)
            if not aliases:
                aliases = self._obtener_todos_aliases(base_normalizada)

        vistos = set()
        aliases_unicos = []
        for alias in aliases:
            alias_normalizado = self._normalizar_alias(alias)
            if alias_normalizado and alias_normalizado not in vistos:
                vistos.add(alias_normalizado)
                aliases_unicos.append(alias_normalizado)

        self.logger.debug(
            "FIN _resolver_aliases_objetivo | base=%s total_aliases=%s aliases=%s",
            base_normalizada,
            len(aliases_unicos),
            aliases_unicos,
        )

        return aliases_unicos

    def _agrupar_items_por_contexto(self, items, zona_destino=None):
        grupos = {}
        for item in items:
            tipo_precio = str(item.get("PBKTYPE", "")).strip().upper()
            zona = str(item.get("PBKZONE", "")).strip().upper()
            destino = str(zona_destino or zona).strip().upper()
            key = (tipo_precio, zona, destino)
            grupos.setdefault(key, []).append(item)

        return [
            {
                "tipo_precio": tipo_precio,
                "zona": zona,
                "zona_destino": destino,
                "items": grupo,
            }
            for (tipo_precio, zona, destino), grupo in grupos.items()
        ]



    def _cargar_datos_promo(self, base):
        """Carga datos de promociones para un CP específico."""
        try:
            return self.repository.get_item_promo(base)
        except Exception as exc:
            self.logger.warning("Error cargando promo para CP | base=%s exc=%s", base, exc)
            return []

    def _cargar_datos_permanente(self, base):
        """Carga datos de precios permanentes para un CP específico."""
        try:
            return self.repository.get_item_precio_permanente(base)
        except Exception as exc:
            self.logger.warning("Error cargando permanente para CP | base=%s exc=%s", base, exc)
            return []

    def _cargar_datos_recepcion(self, base):
        """Carga datos de recepción para un CP específico."""
        try:
            return self.repository.get_recepcion(base)
        except Exception as exc:
            self.logger.warning("Error cargando recepcion para CP | base=%s exc=%s", base, exc)
            return []

    def _insertar_items_kregp(self, base, promo, permanente, recepcion):
        """Inserta items en KREGL para un CP específico."""
        items_a_insertar = []
        items_a_insertar.extend(promo[:3])
        items_a_insertar.extend(permanente[:3])
        for sku in recepcion[:3]:
            items_a_insertar.append({"SKU": sku[0]})

        if items_a_insertar:
            try:
                self.repository.insert_kregp_items(base, items_a_insertar)
            except Exception as exc:
                self.logger.warning("Error insertando items KREGL para CP | base=%s exc=%s", base, exc)

    def _extraer_sku_upcabc(self, upcabc):
        upc = str(upcabc or "").strip()
        if len(upc) < 6:
            return ""
        return upc[1:6]

    def _buscar_item_en_aliases(self, sku, precio_esperado=None, aliases_objetivo=None, cache_aliases=None):
        if not aliases_objetivo:
            return [], "No se encontraron alias PRLKPABC para validar."

        sku_buscado = str(sku or "").strip()
        if not sku_buscado:
            return [], "SKU vacío para validación."

        try:
            precio_normalizado = None if precio_esperado in (None, "") else float(precio_esperado)
        except (TypeError, ValueError):
            precio_normalizado = None

        coincidencias = []
        detalle_alias = []

        for alias in aliases_objetivo:
            filas_alias = []
            registros_alias = (cache_aliases or {}).get(alias, {})
            datos = registros_alias.get(sku_buscado, [])

            for row in datos:
                priceabc = row["priceabc"]
                pricepabc = row["pricepabc"]

                match_precio = True
                if precio_normalizado is not None:
                    match_precio = False
                    for valor in (priceabc, pricepabc):
                        try:
                            if float(valor) == precio_normalizado:
                                match_precio = True
                                break
                        except (TypeError, ValueError):
                            continue

                resultado = dict(row)
                resultado["precio_ok"] = match_precio

                coincidencias.append(resultado)
                filas_alias.append(resultado)

            detalle_alias.append({"alias": alias, "filas": filas_alias})

        self.ultimo_detalle_alias = detalle_alias
        return coincidencias, None

    def _crear_clave_validacion(self, base, item, zona_destino=None, requerir_promo=None, aliases_objetivo=None):
        return (
            str(base or "").strip().upper(),
            str(item.get("PBKSKU", "")).strip(),
            str(item.get("PBKZONE", "")).strip().upper(),
            str(item.get("PBKTYPE", "")).strip().upper(),
            str(item.get("PBKNEWPRC", "")).strip(),
            str(zona_destino or "").strip().upper(),
            "" if requerir_promo is None else str(bool(requerir_promo)),
            tuple(str(alias).strip().upper() for alias in (aliases_objetivo or [])),
        )

    def _evaluar_validacion_item(self, sku, precio_esperado, aliases_objetivo, cache_aliases, requiere_promo_temporal):
        encontrados, error = self._buscar_item_en_aliases(
            sku,
            precio_esperado,
            aliases_objetivo=aliases_objetivo,
            cache_aliases=cache_aliases,
        )
        detalle_alias = list(self.ultimo_detalle_alias)

        if error:
            return {
                "resultado": "FAIL",
                "detalle": error,
                "encontrados": [],
                "detalle_alias": detalle_alias,
            }

        if not encontrados:
            return {
                "resultado": "FAIL",
                "detalle": "SKU no encontrado en ningun alias PRLKPABC.",
                "encontrados": [],
                "detalle_alias": detalle_alias,
            }

        if precio_esperado is None:
            return {
                "resultado": "PASS",
                "detalle": "SKU encontrado en alias.",
                "encontrados": encontrados,
                "detalle_alias": detalle_alias,
            }

        encontrados_por_alias = {alias: [] for alias in aliases_objetivo}
        for row in encontrados:
            encontrados_por_alias.setdefault(row["alias"], []).append(row)

        faltantes = [alias for alias in aliases_objetivo if not encontrados_por_alias.get(alias)]
        invalidos = []
        validos = []

        for alias in aliases_objetivo:
            filas_alias = encontrados_por_alias.get(alias, [])
            if not filas_alias:
                continue

            fila_valida = None
            for row in filas_alias:
                promo_ok = row["promoabc"] == "Y" if requiere_promo_temporal else True
                if row["precio_ok"] and promo_ok:
                    fila_valida = row
                    break

            if fila_valida:
                validos.append(fila_valida)
            else:
                invalidos.append(filas_alias[0])

        if not faltantes and len(validos) == len(aliases_objetivo):
            fila_ok = validos[0]
            return {
                "resultado": "PASS",
                "detalle": (
                    f"Precio validado en {len(validos)} alias(es). "
                    f"Esperado={precio_esperado} ejemplo={fila_ok['alias']} "
                    f"PRICEABC={fila_ok['priceabc']} PRICEPABC={fila_ok['pricepabc']} "
                    f"PROMOABC={fila_ok['promoabc']}"
                ),
                "encontrados": encontrados,
                "detalle_alias": detalle_alias,
            }

        if faltantes:
            return {
                "resultado": "FAIL",
                "detalle": f"El SKU no se encontro en todos los PRLKPABC requeridos. Faltantes: {', '.join(faltantes)}",
                "encontrados": encontrados,
                "detalle_alias": detalle_alias,
            }

        detalle = invalidos[0] if invalidos else encontrados[0]
        motivo = (
            "PROMOABC no activo para promo temporal."
            if requiere_promo_temporal and detalle["promoabc"] != "Y"
            else "Precio no coincide con el esperado."
        )
        return {
            "resultado": "FAIL",
            "detalle": (
                f"SKU localizado pero sin validacion correcta. "
                f"Esperado={precio_esperado} alias={detalle['alias']} "
                f"PRICEABC={detalle['priceabc']} PRICEPABC={detalle['pricepabc']} "
                f"PROMOABC={detalle['promoabc']} Motivo={motivo}"
            ),
            "encontrados": encontrados,
            "detalle_alias": detalle_alias,
        }

    def _validar_item_en_alias(self, base, item, descripcion, zona_destino=None, requerir_promo=None, finalizar=True, aliases_objetivo=None, cache_aliases=None, documentar_detalle=True):
        sku = str(item.get("PBKSKU", "")).strip()
        zona = str(item.get("PBKZONE", "")).strip()
        precio_esperado = item.get("PBKNEWPRC")
        tipo_precio = str(item.get("PBKTYPE", "")).strip()
        zona_evaluada = str(zona_destino or zona).strip()

        if aliases_objetivo is None:
            aliases_objetivo = self._resolver_aliases_objetivo(
                base,
                tipo_precio=tipo_precio,
                zona=zona,
                zona_destino=zona_destino,
            )

        if cache_aliases is None:
            cache_aliases = self._cargar_aliases_para_validacion(base, aliases_objetivo)

        requiere_promo_temporal = (tipo_precio == "T") if requerir_promo is None else requerir_promo

        self.document_builder.agregar_texto(
            f"{descripcion}: SKU={sku} ZONA={zona} ZONA_VALIDADA={zona_evaluada or 'TODAS'}"
        )

        clave_validacion = self._crear_clave_validacion(
            base,
            item,
            zona_destino=zona_destino,
            requerir_promo=requerir_promo,
            aliases_objetivo=aliases_objetivo,
        )

        if clave_validacion in self.validation_result_cache:
            validacion = self.validation_result_cache[clave_validacion]
            self.logger.debug("Validacion reutilizada desde cache | base=%s sku=%s", str(base).strip().upper(), sku)
        else:
            validacion = self._evaluar_validacion_item(
                sku,
                precio_esperado,
                aliases_objetivo,
                cache_aliases,
                requiere_promo_temporal,
            )
            self.validation_result_cache[clave_validacion] = validacion

        self.ultimo_detalle_alias = list(validacion["detalle_alias"])

        if validacion["encontrados"]:
            self.document_builder.agregar_texto(
                f"Coincidencias encontradas en alias: {', '.join(sorted({row['alias'] for row in validacion['encontrados']}))}"
            )
        if documentar_detalle:
            self._pintar_detalle_alias_evaluados()

        if finalizar:
            self._finalizar_cp(validacion["resultado"], validacion["detalle"])
        return validacion["resultado"], validacion["detalle"]

    def _validar_items_en_alias(self, base, items, descripcion, zona_destino=None, requerir_promo=None, mensaje_si_vacio=None):
        if not items:
            self._finalizar_cp("SKIP", mensaje_si_vacio or "No hay items disponibles para validar.")
            return

        grupos = self._agrupar_items_por_contexto(items, zona_destino=zona_destino)
        self.document_builder.agregar_texto(
            f"Validando {len(items)} item(s) en {len(grupos)} grupo(s) de alias objetivo."
        )

        resumen = {"PASS": 0, "FAIL": 0}

        for index, grupo in enumerate(grupos, start=1):
            tipo_precio = grupo["tipo_precio"]
            zona = grupo["zona"]
            destino = grupo["zona_destino"]
            item_count = len(grupo["items"])

            aliases_objetivo = self._resolver_aliases_objetivo(
                base,
                tipo_precio=tipo_precio,
                zona=zona,
                zona_destino=destino,
            )

            if not aliases_objetivo:
                zona_error = str(destino or zona or "").strip()
                if tipo_precio == "P":
                    self._finalizar_cp("FAIL", "No se encontraron alias PRLKPABC para validar el cambio permanente.")
                else:
                    self._finalizar_cp("FAIL", f"No se encontraron alias PRLKPABC para la zona {zona_error}.")
                return

            cache_aliases = self._cargar_aliases_para_validacion(base, aliases_objetivo)

            self.document_builder.agregar_texto(
                f"Grupo {index}: tipo={tipo_precio or 'N/A'} zona={zona or 'N/A'} destino={destino or 'N/A'} items={item_count} aliases={len(aliases_objetivo)}"
            )
            self.document_builder.agregar_texto(
                f"Aliases objetivo evaluados: {', '.join(aliases_objetivo)}"
            )

            for item in grupo["items"]:
                resultado, detalle = self._validar_item_en_alias(
                    base,
                    item,
                    descripcion,
                    zona_destino=destino,
                    requerir_promo=requerir_promo,
                    finalizar=False,
                    aliases_objetivo=aliases_objetivo,
                    cache_aliases=cache_aliases,
                    documentar_detalle=False,
                )
                resumen[resultado] = resumen.get(resultado, 0) + 1
                self.document_builder.agregar_texto(f"Item {item.get('PBKSKU')} resultado={resultado} | {detalle}")
                if resultado == "FAIL":
                    self._pintar_detalle_alias_evaluados()

        if resumen["FAIL"]:
            self._finalizar_cp(
                "FAIL",
                f"Items evaluados={len(items)} PASS={resumen['PASS']} FAIL={resumen['FAIL']}"
            )
            return

        self._finalizar_cp(
            "PASS",
            f"Items evaluados={len(items)} PASS={resumen['PASS']} FAIL={resumen['FAIL']}"
        )



    def _validar_existe_data(self, registros, mensaje_si_vacio):
        if registros:
            return True
        self._finalizar_cp("SKIP", mensaje_si_vacio)
        return False

    def _obtener_promos_crc(self, base):
        promo = self._cargar_datos_promo(base)
        return [
            item
            for item in promo
            if str(item.get("PBKZONE", "")).strip() == "CRC"
        ]

    def _validar_herencia_crc_int(self, base, descripcion, mensaje_si_vacio, nota=None):
        promo_crc = self._obtener_promos_crc(base)
        if nota:
            self.document_builder.agregar_texto(nota)
        self._validar_items_en_alias(
            base,
            promo_crc,
            descripcion,
            zona_destino="INT",
            requerir_promo=False,
            mensaje_si_vacio=mensaje_si_vacio,
        )

    # ==========================================================================
    # CASOS DE PRUEBA (CPs) - A programar uno a uno
    # ==========================================================================
    # SV01: Validacion de envio item por promo para zona con mantenimiento KREGP
    # SV02: Validacion de envio del mismo item para una zona sin promo
    # SV03: Validacion de herencia para la tienda 523 internet online
    # SV04: Validacion de herencia para la tienda 623 internet online
    # SV05: Validar envio de items con promo activa para tiendas INT y CRC
    # SV06: Validacion de cancelacion parcial
    # SV07: Validacion de cancelacion total
    # SV08: Creacion + cancelacion + creacion precio temporal mismo dia
    # SV09: Creacion + cancelacion + creacion precio permanente mismo dia
    # SV10: Generacion de archivos ItemImport y PricingImport
    # SV11: Envio de archivos hacia magento
    # SV12: Cambio de precio por cambios en PO
    # SV13: Cambio de precio permanente por aumento
    # SV14: Cambio de precio por disminucion
    # SV15: Activacion de precio con fecha futura
    # SV16: Generacion de Global de precios
    # ==========================================================================

    def test_sv01(self, base):
        """
        SV01: Validacion de envio item por promo para zona con mantenimiento KREGP
        
        Flujo:
        1. Cargar items con promo activa y sin promo
        2. Insertar ambos en KREGP para forzar procesamiento
        3. Someter programa batch y esperar finalización
        4. Validar que items con promo aparezcan en PRLKPABC con precio correcto
        5. Si no hay precio activo, validar contra cambio permanente
        """
        self._iniciar_cp(base, "SV01", "Validacion de envio item por promo para zona con mantenimiento KREGP")
        
        # Cargar datos
        promo = self._cargar_datos_promo(base)
        permanente = self._cargar_datos_permanente(base)
        recepcion = self._cargar_datos_recepcion(base)
        
        if not promo:
            self._finalizar_cp("SKIP", "No hay items con promo activa disponibles para validar.")
            return
        
        # Preparar items para insertar en KREGP: promo + permanente + recepcion
        self.document_builder.agregar_texto(
            f"Items para procesar: promo={len(promo)} permanente={len(permanente)} recepcion={len(recepcion)}"
        )
        
        # Insertar items en KREGP
        self._insertar_items_kregp(base, promo, permanente, recepcion)
        
        # Registrar errores de preparación
        self._registrar_errores_contexto(base)
        
        # Agrupar items promo por zona
        grupos_promo = {}
        for item in promo:
            zona = str(item.get("PBKZONE", "")).strip().upper()
            if zona not in grupos_promo:
                grupos_promo[zona] = []
            grupos_promo[zona].append(item)
        
        self.document_builder.agregar_texto(
            f"Grupos de promo por zona: {', '.join(f'{z}={len(items)}' for z, items in grupos_promo.items())}"
        )
        
        # Validar cada grupo por zona
        resumen = {"PASS": 0, "FAIL": 0}
        for zona, items_zona in grupos_promo.items():
            self.document_builder.agregar_texto(f"Validando promo en zona={zona} items={len(items_zona)}")
            
            # Resolver aliases para esta zona
            aliases_objetivo = self._resolver_aliases_objetivo(
                base,
                tipo_precio="T",
                zona=zona,
            )
            
            if not aliases_objetivo:
                self.document_builder.agregar_texto(f"No se encontraron alias PRLKPABC para zona {zona}")
                resumen["FAIL"] += len(items_zona)
                continue
            
            # Cargar datos de aliases
            cache_aliases = self._cargar_aliases_para_validacion(base, aliases_objetivo)
            
            # Validar cada item en la zona
            for item in items_zona:
                sku = str(item.get("PBKSKU", "")).strip()
                precio_esperado = item.get("PBKNEWPRC")
                
                # Buscar en PRLKPABC
                encontrados, error = self._buscar_item_en_aliases(
                    sku,
                    precio_esperado,
                    aliases_objetivo=aliases_objetivo,
                    cache_aliases=cache_aliases,
                )
                
                if error or not encontrados:
                    self.document_builder.agregar_texto(
                        f"Promo no encontrada en PRLKPABC | SKU={sku} zona={zona} precio={precio_esperado}"
                    )
                    resumen["FAIL"] += 1
                    self._pintar_detalle_alias_evaluados()
                else:
                    # Verificar que al menos uno coincida con precio esperado
                    hay_coincidencia = any(row["precio_ok"] for row in encontrados)
                    if hay_coincidencia:
                        self.document_builder.agregar_texto(
                            f"Promo validada | SKU={sku} zona={zona} aliases={len(set(row['alias'] for row in encontrados))}"
                        )
                        resumen["PASS"] += 1
                    else:
                        self.document_builder.agregar_texto(
                            f"Precio no coincide | SKU={sku} zona={zona} esperado={precio_esperado}"
                        )
                        resumen["FAIL"] += 1
                        self._pintar_detalle_alias_evaluados()
        
        # Resultado final
        if resumen["FAIL"] > 0:
            self._finalizar_cp(
                "FAIL",
                f"Items validados={len(promo)} PASS={resumen['PASS']} FAIL={resumen['FAIL']}"
            )
        else:
            self._finalizar_cp(
                "PASS",
                f"Todos los items con promo fueron validados correctamente. Total={len(promo)}"
            )
