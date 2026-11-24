import re
import os
import glob
from openpyxl import Workbook

pattern04 = re.compile(r'04|(?<!\d)4(?!\d)')
COLS = [(0,15),(30,57),(60,74),(168,179),(335,340),(341,344)]
HEADERS = ['NUMERO_TARJETA','NOMBRE','CEDULA','TELEFONO','COD_TIPO','TIPO_DOC']

def fix_ef(e, f):
    letters = re.sub(r'[^A-Za-z]', '', e)
    leftover = re.sub(r'[A-Za-zÀ-ÿ\s]', '', e)
    return letters, leftover + f

# Buscar el TXT en la carpeta input/
txt_files = glob.glob('input/*.txt')
if not txt_files:
    print("No se encontró ningún .txt en la carpeta input/")
    exit(1)

input_path = txt_files[0]
base_name = os.path.splitext(os.path.basename(input_path))[0]
print(f"Procesando: {input_path}")

seen = set()
results_txt = []
rows_excel = []
discarded_non_num = 0
discarded_overflow = 0
dup_count = 0
total_lines = 0

with open(input_path, 'r', encoding='latin-1') as f:
    for line in f:
        line = line.rstrip('\n').rstrip('\r')
        if len(line) < 343:
            continue
        total_lines += 1
        col1_raw = line[0:15].strip()
        penultima = line[331:343]

        if not col1_raw.isdigit():
            discarded_non_num += 1
            continue

        col1_padded = col1_raw.zfill(15)
        if col1_padded not in seen:
            seen.add(col1_padded)
            results_txt.append(col1_padded)
            row = [line[s:e].strip() if len(line) >= e else '' for s,e in COLS]
            row[0] = col1_padded
            row[4], row[5] = fix_ef(row[4], row[5])
            rows_excel.append(row)

        if pattern04.search(penultima):
            dup_count += 1
            col1_unpadded = col1_raw.lstrip('0') or '0'
            new_val = '999' + col1_unpadded
            if len(new_val) <= 15:
                new_padded = new_val.zfill(15)
                if new_padded not in seen:
                    seen.add(new_padded)
                    results_txt.append(new_padded)
                    new_row = rows_excel[-1].copy() if rows_excel else ['']*6
                    new_row = new_row.copy()
                    new_row[0] = new_padded
                    rows_excel.append(new_row)
            else:
                discarded_overflow += 1

os.makedirs('output', exist_ok=True)

# TXT depurado
txt_path = f'output/{base_name}_DEPURADO.txt'
with open(txt_path, 'w', encoding='latin-1') as f:
    f.write('\n'.join(results_txt) + '\n')

# Excel en dos partes
mid = len(rows_excel) // 2
parts = [rows_excel[:mid], rows_excel[mid:]]
for i, part in enumerate(parts, 1):
    wb = Workbook(write_only=True)
    ws = wb.create_sheet()
    ws.append(HEADERS)
    for r in part:
        ws.append(r)
    path = f'output/{base_name}_Parte{i}.xlsx'
    wb.save(path)
    print(f"Parte {i}: {len(part):,} filas")

added_999 = sum(1 for r in results_txt if r.lstrip('0').startswith('999'))
base_uniq = len(results_txt) - added_999

print(f"\nResumen:")
print(f"  Registros recibidos:      {total_lines:,}")
print(f"  Descartados no numéricos: {discarded_non_num:,}")
print(f"  Registros base únicos:    {base_uniq:,}")
print(f"  Registros 999 agregados:  {added_999:,}")
print(f"  Descartados overflow:     {discarded_overflow:,}")
print(f"  Total único TXT:          {len(results_txt):,}")
print(f"\nArchivos generados en output/")
