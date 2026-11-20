# ==============================================================================
# PROTOTIPO DE MESA DE PARTES DIGITAL CON VALIDACIÓN DE PAGO EN BANCO DE LA NACIÓN
# Curso: Fundamentos de Programación (CIIN1205P)
# ==============================================================================

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------------------------
# ESTRUCTURA DE DATOS: ARREGLOS PARALELOS EN MEMORIA
# ------------------------------------------------------------------------------
codigos_exp = []
dnis_ciudadano = []
nombres_ciudadano = []
asuntos_tramite = []
codigos_pago_bn = []    # Lista paralela para el voucher del Banco de la Nación
estados_tramite = []    # RECIBIDO, EN PROCESO, OBSERVADO, RESUELTO, ARCHIVADO

NOMBRE_ARCHIVO_TXT = "expedientes.txt"
NOMBRE_ARCHIVO_EXCEL = "Mesa_de_Partes_Digital_LosOlivos.xlsx"

# ------------------------------------------------------------------------------
# MÓDULO DE PERSISTENCIA Y MANEJO DE ARCHIVOS (.TXT)
# ------------------------------------------------------------------------------
def cargar_desde_archivo():
    """Carga los registros previos almacenados en el archivo local .txt"""
    try:
        archivo = open(NOMBRE_ARCHIVO_TXT, "r")
        lineas = archivo.readlines()
        archivo.close()
        
        for linea in lineas:
            datos = linea.strip().split(",")
            if len(datos) == 6:
                codigos_exp.append(datos[0])
                dnis_ciudadano.append(datos[1])
                nombres_ciudadano.append(datos[2])
                asuntos_tramite.append(datos[3])
                codigos_pago_bn.append(datos[4])
                estados_tramite.append(datos[5])
        print(" Archivos de expedientes cargados correctamente desde .txt")
    except FileNotFoundError:
        print(" Creando nueva base de datos local (expedientes.txt)...")

def guardar_en_archivo():
    """Guarda permanentemente las listas paralelas en el archivo .txt"""
    archivo = open(NOMBRE_ARCHIVO_TXT, "w")
    for i in range(len(codigos_exp)):
        linea = (codigos_exp[i] + "," + 
                 dnis_ciudadano[i] + "," + 
                 nombres_ciudadano[i] + "," + 
                 asuntos_tramite[i] + "," + 
                 codigos_pago_bn[i] + "," + 
                 estados_tramite[i] + "\n")
        archivo.write(linea)
    archivo.close()
    print(" Datos sincronizados con el archivo de texto (.txt).")

# ------------------------------------------------------------------------------
# MÓDULO DE GENERACIÓN DE REPORTES EXCEL (OPENPYXL)
# ------------------------------------------------------------------------------
def exportar_a_excel():
    """Genera un reporte institucional en formato .xlsx usando openpyxl"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Registro de Expedientes"
    ws.views.sheetView[0].showGridLines = True

    # Encabezado Institucional
    ws.merge_cells('A1:F1')
    ws['A1'] = "MUNICIPALIDAD DISTRITAL DE LOS OLIVOS - MESA DE PARTES DIGITAL"
    ws['A1'].font = Font(name='Calibri', size=13, bold=True, color='FFFFFF')
    ws['A1'].fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

    headers = ["Código Expediente", "DNI / RUC", "Ciudadano / Empresa", "Asunto del Trámite", "Cód. Pago BN", "Estado"]
    ws.append(headers)

    # Estilar encabezados
    header_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    header_font = Font(name='Calibri', size=11, bold=True, color='1F4E79')
    
    for col_num in range(1, 7):
        cell = ws.cell(row=2, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Llenar datos desde arreglos paralelos
    for i in range(len(codigos_exp)):
        ws.append([
            codigos_exp[i],
            dnis_ciudadano[i],
            nombres_ciudadano[i],
            asuntos_tramite[i],
            codigos_pago_bn[i],
            estados_tramite[i]
        ])

    # Ajustar ancho de columnas automáticamente
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(NOMBRE_ARCHIVO_EXCEL)
    print(" Reporte actualizado generado en Excel:", NOMBRE_ARCHIVO_EXCEL)

# ------------------------------------------------------------------------------
# MÓDULO DE VALIDACIONES
# ------------------------------------------------------------------------------
def validar_dni_ruc(documento):
    """Valida que tenga 8 (DNI) u 11 (RUC) dígitos numéricos"""
    return (len(documento) == 8 or len(documento) == 11) and documento.isdigit()

def validar_pago_banco_nacion(codigo_pago):
    """
    Simulación de validación de comprobante del Banco de la Nación.
    Exige que el código no esté vacío y tenga al menos 6 caracteres alfanuméricos.
    """
    codigo_limpio = codigo_pago.strip()
    return len(codigo_limpio) >= 6

# ------------------------------------------------------------------------------
# INTERFAZ Y FUNCIONES DE NEGOCIO
# ------------------------------------------------------------------------------
def mostrar_menu():
    print("\n" + "="*50)
    print("    MESA DE PARTES DIGITAL - LOS OLIVOS")
    print("="*50)
    print("1. Registrar nuevo expediente (requiere Pago BN)")
    print("2. Listar todos los expedientes")
    print("3. Buscar expediente por código")
    print("4. Ordenar expedientes (Algoritmo de Burbuja)")
    print("5. Exportar reporte institucional a Excel")
    print("6. Salir")
    print("="*50)

def registrar_expediente():
    print("\n--- NUEVO REGISTRO DE EXPEDIENTE ---")
    nuevo_codigo = "EXP-" + str(1001 + len(codigos_exp))
    
    # 1. Validar DNI / RUC
    dni = input("Ingrese DNI (8 dígitos) o RUC (11 dígitos): ").strip()
    while not validar_dni_ruc(dni):
        print(" Error: Documento inválido. Debe contener 8 u 11 números.")
        dni = input("Ingrese DNI/RUC válido: ").strip()

    # 2. Capturar Nombre
    nombre = input("Ingrese nombre completo o Razón Social: ").strip().upper()
    while len(nombre) == 0:
        print(" Error: El nombre no puede estar vacío.")
        nombre = input("Ingrese el nombre: ").strip().upper()

    # 3. Capturar Asunto
    asunto = input("Ingrese el asunto o motivo del trámite: ").strip()
    while len(asunto) == 0:
        print(" Error: El asunto no puede estar vacío.")
        asunto = input("Ingrese el asunto: ").strip()

    # 4. SIMULACIÓN DE PAGO BANCO DE LA NACIÓN
    print("\n[REQUISITO TASA ADM.] Ingrese el código de voucher del Banco de la Nación.")
    pago_bn = input("Código de Pago BN (Ej: BN-2026-88491): ").strip().upper()
    
    while not validar_pago_banco_nacion(pago_bn):
        print(" Error: Comprobante no válido. Debe ingresar un código de voucher correcto (mínimo 6 caracteres).")
        pago_bn = input("Ingrese nuevamente el código del voucher BN: ").strip().upper()

    print(" Comprobante de Banco de la Nación verificado correctamente.")

    # Guardar en arreglos paralelos
    codigos_exp.append(nuevo_codigo)
    dnis_ciudadano.append(dni)
    nombres_ciudadano.append(nombre)
    asuntos_tramite.append(asunto)
    codigos_pago_bn.append(pago_bn)
    estados_tramite.append("RECIBIDO")

    # Persistencia
    guardar_en_archivo()
    exportar_a_excel()

    print("\n ¡Expediente registrado con éxito!")
    print("Código Asignado:", nuevo_codigo)
    print("Estado Inicial  : RECIBIDO")

def listar_expedientes():
    print("\n--- LISTADO COMPLETO DE EXPEDIENTES ---")
    if len(codigos_exp) == 0:
        print("No hay expedientes registrados en el sistema.")
    else:
        for i in range(len(codigos_exp)):
            print("--------------------------------------------------")
            print("Código Exp. :", codigos_exp[i])
            print("DNI / RUC   :", dnis_ciudadano[i])
            print("Ciudadano   :", nombres_ciudadano[i])
            print("Asunto      :", asuntos_tramite[i])
            print("Voucher BN  :", codigos_pago_bn[i])
            print("Estado      :", estados_tramite[i])

def buscar_expediente():
    print("\n--- BÚSQUEDA DE EXPEDIENTE POR CÓDIGO ---")
    if len(codigos_exp) == 0:
        print("No hay registros para consultar.")
        return

    buscado = input("Ingrese código de expediente a consultar (Ej: EXP-1001): ").strip().upper()
    encontrado = False

    for i in range(len(codigos_exp)):
        if codigos_exp[i] == buscado:
            print("\n ¡EXPEDIENTE ENCONTRADO!")
            print("Código Exp. :", codigos_exp[i])
            print("DNI / RUC   :", dnis_ciudadano[i])
            print("Ciudadano   :", nombres_ciudadano[i])
            print("Asunto      :", asuntos_tramite[i])
            print("Voucher BN  :", codigos_pago_bn[i])
            print("Estado      :", estados_tramite[i])
            encontrado = True
            break
            
    if not encontrado:
        print(" No se encontró ningún expediente con el código:", buscado)

def ordenar_expedientes():
    """Ordenamiento de Burbuja sincronizando las 6 listas paralelas"""
    print("\n--- ORDENAMIENTO DE EXPEDIENTES (ALGORITMO BURBUJA) ---")
    n = len(codigos_exp)
    if n <= 1:
        print("No hay suficientes registros para ejecutar el ordenamiento.")
        return

    for i in range(n - 1):
        for j in range(0, n - i - 1):
            if codigos_exp[j] > codigos_exp[j + 1]:
                # Intercambiar elementos en las 6 listas paralelas simultáneamente
                codigos_exp[j], codigos_exp[j + 1] = codigos_exp[j + 1], codigos_exp[j]
                dnis_ciudadano[j], dnis_ciudadano[j + 1] = dnis_ciudadano[j + 1], dnis_ciudadano[j]
                nombres_ciudadano[j], nombres_ciudadano[j + 1] = nombres_ciudadano[j + 1], nombres_ciudadano[j]
                asuntos_tramite[j], asuntos_tramite[j + 1] = asuntos_tramite[j + 1], asuntos_tramite[j]
                codigos_pago_bn[j], codigos_pago_bn[j + 1] = codigos_pago_bn[j + 1], codigos_pago_bn[j]
                estados_tramite[j], estados_tramite[j + 1] = estados_tramite[j + 1], estados_tramite[j]

    guardar_en_archivo()
    exportar_a_excel()
    print(" ¡Expedientes ordenados por código exitosamente!")

# ------------------------------------------------------------------------------
# PROGRAMA PRINCIPAL
# ------------------------------------------------------------------------------
def main():
    cargar_desde_archivo()
    opcion = ""
    while opcion != "6":
        mostrar_menu()
        opcion = input("Seleccione una opción (1-6): ").strip()
        
        if opcion == "1":
            registrar_expediente()
        elif opcion == "2":
            listar_expedientes()
        elif opcion == "3":
            buscar_expediente()
        elif opcion == "4":
            ordenar_expedientes()
        elif opcion == "5":
            exportar_a_excel()
        elif opcion == "6":
            print("\n¡Gracias por utilizar el Prototipo de Mesa de Partes Digital!")
        else:
            print(" Opción no válida. Intente nuevamente.")

# Ejecutar aplicación
main()
