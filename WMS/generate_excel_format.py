from openpyxl import Workbook

def generate_excel_template(filename="transferencias.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Transferencias"

    # Define headers based on the expected input from leer_transferencias_desde_excel and agrupar_transferencias_por_id
    headers = ["id", "tipo", "bodega", "referencia", "tienda", "sku"]
    ws.append(headers)

    wb.save(filename)
    print(f"Archivo '{filename}' generado con éxito con los encabezados: {', '.join(headers)}")

if __name__ == "__main__":
    generate_excel_template()
