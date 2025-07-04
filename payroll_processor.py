import pandas as pd
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


def extract_records(
    input_path: str,
    skip_header_rows: int = 12,
    drop_col_index: int = 8,
    providers: list = None
) -> pd.DataFrame:
    """
    Etapa 1: Carga y extracción de datos del reporte encounters detail.

    Args:
        input_path (str): Ruta al archivo de Excel de entrada.
        skip_header_rows (int): Número de filas a omitir antes de los datos.
        drop_col_index (int): Índice de la columna a eliminar.
        providers (list): Lista de nombres de proveedores a procesar.

    Returns:
        pd.DataFrame: DataFrame con columnas [Provider, Date, Patient, Procedure, Receipt].
    """
    # Configuración por defecto de proveedores
    if providers is None:
        providers = ["Joseph", "Colin", "Lisa", "Katelyn", "Trinity", "Sara", "Nicole"]

    # Cargar datos sin encabezado
    df = pd.read_excel(input_path, header=None, skiprows=skip_header_rows)

    # Eliminar columna no deseada
    df.drop(df.columns[drop_col_index], axis=1, inplace=True)

    records = []
    current_provider = None

    # Iterar filas
    for _, row in df.iterrows():
        col0 = str(row[0]).strip() if pd.notna(row[0]) else ''
        col5 = row[5]

        # Detectar encabezados
        if pd.notna(row[0]) and pd.isna(col5):
            header = col0
            if header.startswith("Total"):
                current_provider = None
            else:
                current_provider = header
            continue
        # Detectar subtotales
        if pd.notna(col0) and col0.startswith("Total"):
            current_provider = None
            continue

        # Extraer registros de citas
        if current_provider and pd.notna(col5):
            patient = str(row[5]).strip()
            date = pd.to_datetime(row[10]).date() if pd.notna(row[10]) else None
            procedure = str(row[13]).strip() if pd.notna(row[13]) else ''
            raw_receipt = str(row[31]) if pd.notna(row[31]) else ''

            # Limpiar y convertir receipt
            nums = re.findall(r"[\d,]*\.?\d+", raw_receipt.replace(',', ''))
            receipt = float(nums[0]) if nums else 0.0

            records.append({
                'Provider': current_provider,
                'Date': date,
                'Patient': patient,
                'Procedure': procedure,
                'Receipt': receipt
            })

    return pd.DataFrame(records)


def generate_frames(df: pd.DataFrame) -> pd.DataFrame:
    """
    Etapa 2: Filtrado de proveedores PMHNP-BC.
    """
    # Filtrar los no PMHNP-BC para output principal
    df_main = df[~df['Provider'].str.contains('PMHNP-BC', na=False)].copy()
    return df_main


def write_consolidated(
    df_main: pd.DataFrame,
    input_path: str,
    output_path: str
) -> None:
    """
    Genera el archivo Excel consolidado con datos y nota split para PMHNP-BC.
    """
    # Identificar PMHNP-BC distintos
    pmhnp_list = sorted({p for p in df_main['Provider'].unique() if 'PMHNP-BC' in p})

    wb = Workbook()
    ws = wb.active
    ws.title = 'Payroll Consolidado'

    # Escribir datos principales
    for r in dataframe_to_rows(df_main, index=False, header=True):
        ws.append(r)

    # Agregar sección de nota al final
    ws.append([])
    ws.append(['Proveedores con credenciales PMHNP-BC:'])
    for prov in pmhnp_list:
        ws.append([f'- {prov}'])
    ws.append([])
    ws.append(['Estos proveedores se pagarán con un split 70/30 (70% al proveedor, 30% para la práctica) a medida que se reciban los pagos.'])

    wb.save(output_path)


def process_payroll(
    input_path: str,
    output_path: str,
    skip_header_rows: int = 12,
    drop_col_index: int = 8
) -> pd.DataFrame:
    """
    Flujo completo para procesar payroll desde un encounters detail.

    Returns:
        pd.DataFrame: DataFrame filtrado (Etapa 2) para inspección.
    """
    # Etapa 1: extracción
    df_all = extract_records(input_path, skip_header_rows, drop_col_index)
    # Etapa 2: filtrado
    df_main = generate_frames(df_all)
    # Etapa 3: Escritura de archivo
    write_consolidated(df_main, input_path, output_path)
    return df_main


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='Procesador de payroll para encounters detail Tebra')
    parser.add_argument('--input', required=True, help='Ruta al archivo de input Excel')
    parser.add_argument('--output', required=True, help='Ruta al archivo de output Excel')
    args = parser.parse_args()

    df_result = process_payroll(args.input, args.output)
    print('Payroll procesado. Resultado en', args.output)