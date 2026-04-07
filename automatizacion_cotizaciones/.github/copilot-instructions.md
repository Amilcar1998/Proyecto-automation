# AI Coding Instructions for Automatización de Cotizaciones

## Project Overview
This is a web automation project that generates purchase order quotations by:
1. Logging into an internal Unicomer web system (`http://ursvtols8b3c:8080/engage-unicomer-web/`)
2. Reading product codes from Excel files
3. Automatically adding products to quotation forms
4. Extracting quotation numbers and generating Word reports

## Architecture & Key Components

### Core Scripts
- **`automatizacion_cotizaciones_v2.py`** - Main automation script with Excel integration, screenshot capture, and Word generation
- **`crear_excel_ejemplo.py`** - Utility to generate sample Excel files
- **`instalar_geckodriver.py`** - Firefox WebDriver setup automation

### Data Flow
1. Excel input (`productos_ejemplo.xlsx`) → pandas DataFrame
2. Selenium WebDriver automation → web form interaction + screenshot capture
3. Screenshots embedded in Word document → automatic cleanup
4. Quotation extraction → comprehensive Word report generation

## Critical Development Patterns

### Selenium Element Discovery Strategy
The project uses progressive fallback selectors due to unstable web UI:
```python
selectores = [
    (By.XPATH, "//input[@value='Comience una Orden de Compra']"),
    (By.CSS_SELECTOR, "a[href='/engage-unicomer-web/saleOrder/index']"),
    (By.PARTIAL_LINK_TEXT, "Comience una Orden de Compra"),
    # Multiple fallback strategies...
]
```
**Always add new selectors to these arrays when encountering element location failures.**

### Excel Column Flexibility
Product codes are read with flexible column naming:
```python
codigo = producto.get('codigo') or producto.get('producto') or producto.get('Codigo') or producto.get('Producto')
```
**When adding new Excel processing, maintain this case-insensitive, multi-name approach.**

### Error Handling & Debugging
The project emphasizes visual debugging through:
- **Systematic screenshot capture** at each critical step using `capturar_screenshot()` 
- **Automatic embedding** of screenshots in Word documents with descriptive titles
- **Automatic cleanup** of screenshot files after processing
- Comprehensive logging with structured messages

**Always use the `capturar_screenshot()` function when adding new automation steps.**

## Development Workflows

### Running the Main Automation
```powershell
python automatizacion_cotizaciones_v2.py productos_ejemplo.xlsx
```

### Generating Test Data
```powershell
python crear_excel_ejemplo.py  # Creates productos_ejemplo.xlsx
```

### Environment Setup
```powershell
python instalar_geckodriver.py  # Downloads and installs Firefox driver
```

## Key Integration Points

### Web System Dependencies
- **Login credentials**: Hard-coded as `EMPLOYEE_ID = "pos"`, `PASSWORD = "pos1234"`
- **Target URL**: Internal system at `http://ursvtols8b3c:8080/engage-unicomer-web/`
- **Critical UI elements**: 
  - Login: `#employeeId`, `#password`
  - Product search: `#saleItemSearchInput`
  - Quotation number: `li.quotationNumberTitle p:last-child`

### Error Recovery Mechanisms
- Automatic popup dismissal for "ERROR TO OBTAIN THE PARAMATER: permitirNoInventariadosSeguros"
- Multiple login verification strategies (text search, URL change detection)
- Graceful continuation despite individual product failures

### Output Generation
- Word documents using `python-docx` with structured reporting and embedded screenshots
- **Screenshot management**: Automatic capture, embedding, and cleanup via `SCREENSHOTS_TOMADAS` list
- Filename pattern: `Reporte_Cotizaciones_YYYYMMDD_HHMMSS.docx`
- **No persistent debug files**: All screenshots are embedded then deleted automatically

## Project-Specific Conventions

### Constants & Configuration
Place hardcoded values at module top:
```python
URL_LOGIN = "http://ursvtols8b3c:8080/engage-unicomer-web/"
TIMEOUT = 15
```

### Function Naming
Use Spanish descriptive names matching domain:
- `realizar_login()` - performs login
- `buscar_boton_orden_compra()` - searches for purchase order button
- `agregar_producto()` - adds product to quotation

### Logging Pattern
Use structured logging with checkmarks:
```python
LOGGER.info("✓ Login exitoso - Mensaje 'Bienvenido' detectado")
LOGGER.error("✗ No se pudo encontrar el botón")
```

## Working with This Codebase

When modifying automation logic:
1. **Use `capturar_screenshot()` function** instead of direct `driver.save_screenshot()`
2. **Always provide descriptive titles** for screenshots to appear properly in Word document
3. **Update selector arrays** when web UI changes break existing automation
4. **Preserve Spanish naming conventions** and logging patterns
5. **Test with `productos_ejemplo.xlsx`** for consistent Excel processing

When adding features:
- Follow the modular function structure (each automation step = one function)
- Add comprehensive error handling with visual debugging via `capturar_screenshot()`
- Screenshots are automatically managed - just call the function, cleanup is automatic
- Update Word document generation to include new data fields
- Maintain backward compatibility with existing Excel column names
- siempre utiliza el estilo calibri para la generacion de todo el contenido de documentos
- cuando tomes capturas asegurate que las imagenes sean claras y legibles 
- no hagas versiones alternativas de funciones ya existentes; extiende su funcionalidad dentro de la misma funcionalidad.  
-no agregaras iconos, simbolos o emojis
-no tienes que generar codigo redundante y debes aplicar arquitecturas para hacer el codigo mas eficiente y facil de mantener 
- trabaja con los archivos de prueba que se te estan pasando no tienes que generar nuevos archivos de prueba
- cada vez que tengamos una version estable o tratemos de agregar una funcionalidad nueva  haz backup
-evita codigo duplicado e inecesario
-evita ignorar las funcionalidades ya existentes 
