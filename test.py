import pandas as pd
import os

# Definir el archivo de entrada en la variable
input_file = './ordenes-07-11-2024082211106723.xlsx'

def obtener_extension_archivo(nombre_archivo):
    return os.path.splitext(nombre_archivo)[1].lower()

def leer_archivo(nombre_archivo):
    extension = obtener_extension_archivo(nombre_archivo)
    try:
        if extension == '.csv':
            df = pd.read_csv(nombre_archivo)
        elif extension in ['.xls', '.xlsx']:
            df = pd.read_excel(nombre_archivo)
        else:
            raise ValueError("El archivo debe ser .csv o .xlsx/.xls")
        return df
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
        exit(1)

def detectar_columna_estado(df):
    palabras_clave = ["estado", "status", "situación"]
    for columna in df.columns:
        if any(palabra in columna.lower() for palabra in palabras_clave):
            print(f"Columna detectada para conteo de estados de pedido: '{columna}'")
            return columna
    print("No se encontró ninguna columna que represente estados de pedido.")
    exit(1)

def contar_estados_y_sumar_ganancias(df, columna_estado):
    conteo = df[columna_estado].value_counts().reset_index()
    conteo.columns = ['Estado', 'Cantidad']
    
    if 'ganancia' in df.columns.str.lower():
        columna_ganancia = df.columns[df.columns.str.lower() == 'ganancia'][0]
        entregado_keywords = ['entregado', 'completado', 'finalizado']
        
        ganancias_entregados = df[df[columna_estado].str.lower().isin(entregado_keywords)][columna_ganancia].sum()
        conteo['Ganancia Total'] = conteo['Estado'].apply(
            lambda estado: ganancias_entregados if estado.lower() in entregado_keywords else None
        )
    else:
        conteo['Ganancia Total'] = None
        print("No se detectó una columna de ganancias. La columna de 'Ganancia Total' estará vacía.")
    
    return conteo

def guardar_csv(df, nombre_salida):
    try:
        # Ruta a la carpeta de Descargas del usuario
        carpeta_descargas = os.path.join(os.path.expanduser("~"), "Downloads")  # Usar "Downloads" en inglés para asegurar la compatibilidad
        if not os.path.exists(carpeta_descargas):
            print(f"La carpeta de descargas '{carpeta_descargas}' no existe.")
            exit(1)

        ruta_salida = os.path.join(carpeta_descargas, nombre_salida)
        
        df.to_csv(ruta_salida, index=False)
        print(f"Archivo guardado exitosamente como '{ruta_salida}'.")
    except Exception as e:
        print(f"Error al guardar el archivo: {e}")
        exit(1)

def main():
    print("=== Contador de Estados de Pedido ===")

    if not os.path.isfile(input_file):
        print(f"El archivo '{input_file}' no existe.")
        exit(1)

    df = leer_archivo(input_file)
    columna_estado = detectar_columna_estado(df)
    conteo_estados = contar_estados_y_sumar_ganancias(df, columna_estado)

    nombre_salida = 'conteoooo_estados.csv'
    guardar_csv(conteo_estados, nombre_salida)

# Llamada explícita a la función principal para iniciar el programa
if __name__ == "__main__":
    main()
