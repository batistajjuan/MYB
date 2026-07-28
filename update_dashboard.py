#!/usr/bin/env python3
"""
AutoRefricentro MYB — Dashboard Auto-Update Script
===================================================
Descarga el .bak del ERP, lo restaura en SQL Server, ejecuta el query
de ventas y actualiza index.html con los datos frescos.

Modo BAK  (automático vía GitHub Actions):
  BAK_URL="https://..." SA_PASS="Tu$Contrasena1" python3 update_dashboard.py

Modo CSV  (fallback manual, si el ERP exporta CSV):
  CSV_URL="https://..." python3 update_dashboard.py

Variables de entorno:
  BAK_URL   URL pública del archivo .bak del ERP
  SA_PASS   Contraseña del SA de SQL Server (sólo modo BAK)
  CSV_URL   URL pública del CSV exportado del ERP (modo alternativo)
"""

import csv
import gzip
import io
import json
import os
import re
import subprocess
import sys
import time
import urllib.request
import zipfile
from collections import defaultdict
from datetime import datetime

# ── Configuración ──────────────────────────────────────────────────────────────
HTML_FILE  = "index.html"
DB_NAME    = "AutoRefricentro"
BAK_PATH   = "/tmp/autorefricentro.bak"
CSV_PATH   = "/tmp/ventas_query.csv"
SQLCMD     = "/opt/mssql-tools18/bin/sqlcmd"  # path en Ubuntu GitHub Actions runner

# Nombres a EXCLUIR de los vendedores (en minúsculas)
EXCLUDED = {
    'admin',
    'almacen principal',
    'almacen santiago',
    'bianca iris joaquin joaquin',
    'kendra donaira gonzalez medina',
    'oficina',
    'ruth altagracia alvarez gonzalez',
}

# ── Valores históricos hardcodeados (Ene 2025 — dato de campo no capturado en nuevo CSV) ──
JAN_2025_TOTAL     = 20784174
JAN_2025_CAMPO     = 16914689
JAN_2025_PRINCIPAL =  2697886
JAN_2025_SANTIAGO  =   913757
JAN_2025_ORIENTAL  =   257842

# ── Nombres de meses en español ───────────────────────────────────────────────
MONTH_LABELS_ES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']

# ── Sucursales ─────────────────────────────────────────────────────────────────
SUC_DEF = {
    'TIENDA PRINCIPAL-SD':       ('Tienda Principal-SD',        'T. Principal',  's1'),
    'SUCURSAL SANTIAGO':         ('Sucursal Santiago',          'Santiago',      's2'),
    'SUCURSAL ZONA ORIENTAL-SD': ('Sucursal Zona Oriental-SD',  'Zona Oriental', 's3'),
}
SUC_ORDER = list(SUC_DEF)

# ── SQL query a ejecutar contra el .bak ───────────────────────────────────────
VENTAS_QUERY = """
SET NOCOUNT ON;
SELECT
    YEAR(v.VTA_FECHA)                                         AS ANO,
    MONTH(v.VTA_FECHA)                                        AS MES,
    mt.TRA_DESCRIPCION                                        AS NOMBREVENDEDOR,
    v.VTA_COD_CLIENTE                                         AS CLIENTE,
    v.VTA_RAZON_SOCIAL                                        AS CL_NOMBRE,
    v.VTA_NUMERO                                              AS FACTURA,
    vd.DTV_PRECIO_ORIGINAL * vd.DTV_CANTIDAD                 AS VENTAS,
    vd.DTV_CANTIDAD                                           AS CANTIDAD,
    bs.UBI_DESCRIPCION                                        AS SUCURSAL,
    (vd.DTV_PRECIO_ORIGINAL*vd.DTV_CANTIDAD)
      - (vd.DTV_CANTIDAD*vd.DTV_COSTO)                       AS MARGEN
FROM  MAE_CC_VENTAS V
INNER JOIN MAE_CC_VENTAS_DETALLE VD
       ON  V.VTA_COD_TIPO_DOCUMENTO = VD.DTV_COD_TIPO_DOCUMENTO
       AND V.VTA_NUMERO             = VD.DTV_NUMERO
INNER JOIN MAE_MATERIALES MA
       ON  MA.PRO_CODIGO = VD.DTV_COD_PRODUCTO
INNER JOIN MAE_TRABAJADORES MT
       ON  MT.TRA_CODIGO = VD.DTV_COD_VENDEDOR
INNER JOIN BAS_SUCURSALES bs
       ON  bs.UBI_CODIGO = mt.TRA_COD_SUCURSAL
INNER JOIN CON_TIPO_DOCUMENTO CD
       ON  CD.TDC_CODIGO = V.VTA_COD_TIPO_DOCUMENTO
       AND CD.TDC_FISCAL = 1
WHERE V.VTA_ANULADO = 0
  AND V.VTA_FECHA  >= '2025-01-01';
"""

# ═══════════════════════════════════════════════════════════════════════════════
#  MODO BAK — Descargar, restaurar, ejecutar query
# ═══════════════════════════════════════════════════════════════════════════════

def sqlcmd(query, db=None, extra_args=None):
    """Ejecuta un query via sqlcmd y devuelve (stdout, returncode)."""
    sa_pass = os.environ['SA_PASS']
    cmd = [SQLCMD, '-S', 'localhost', '-U', 'sa', '-P', sa_pass, '-No']
    if db:
        cmd += ['-d', db]
    if extra_args:
        cmd += extra_args
    cmd += ['-Q', query]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    return r.stdout, r.returncode

def wait_for_sqlserver(max_wait=120):
    """Espera hasta que SQL Server acepte conexiones."""
    print("    Esperando SQL Server", end='', flush=True)
    for _ in range(max_wait // 5):
        out, code = sqlcmd("SELECT 1")
        if code == 0 and '1' in out:
            print(" ✓")
            return
        print('.', end='', flush=True)
        time.sleep(5)
    print("\nERROR: SQL Server no respondió en tiempo.")
    sys.exit(1)

def download_file(url, dest):
    print(f"    Descargando → {dest}")
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=300) as r, open(dest, 'wb') as f:
        chunk = 1024 * 1024
        total = 0
        while True:
            data = r.read(chunk)
            if not data:
                break
            f.write(data)
            total += len(data)
            print(f"\r    {total/1e6:.1f} MB descargados...", end='', flush=True)
    print(f"\r    {total/1e6:.1f} MB descargados ✓")

def detect_and_extract_bak(downloaded_path, final_bak_path):
    """
    Detecta el formato del archivo descargado y extrae el .bak.
    Soporta: RAR, ZIP, GZIP, o .bak directo.
    """
    with open(downloaded_path, 'rb') as f:
        magic = f.read(8)

    extract_dir = "/tmp/bak_extract"
    os.makedirs(extract_dir, exist_ok=True)

    # RAR (Rar!\x1a\x07)
    if magic[:4] == b'Rar!':
        print("    Formato detectado: RAR — extrayendo con 7z...")
        r = subprocess.run(
            ['7z', 'e', downloaded_path, f'-o{extract_dir}', '-y'],
            capture_output=True, text=True
        )
        if r.returncode != 0:
            print(f"ERROR al extraer RAR:\n{r.stdout}\n{r.stderr}")
            sys.exit(1)
        _move_largest_file(extract_dir, final_bak_path)
        print("    Extracción RAR completada ✓")

    # ZIP (PK\x03\x04)
    elif magic[:4] == b'PK\x03\x04':
        print("    Formato detectado: ZIP — extrayendo...")
        with zipfile.ZipFile(downloaded_path) as z:
            bak_files = [n for n in z.namelist() if n.lower().endswith('.bak')]
            target = bak_files[0] if bak_files else max(z.namelist(), key=lambda n: z.getinfo(n).file_size)
            print(f"    Extrayendo: {target}")
            with z.open(target) as src, open(final_bak_path, 'wb') as dst:
                while True:
                    chunk = src.read(1024 * 1024)
                    if not chunk:
                        break
                    dst.write(chunk)
        print("    Extracción ZIP completada ✓")

    # GZIP (\x1f\x8b)
    elif magic[:2] == b'\x1f\x8b':
        print("    Formato detectado: GZIP — descomprimiendo...")
        with gzip.open(downloaded_path, 'rb') as src, open(final_bak_path, 'wb') as dst:
            while True:
                chunk = src.read(1024 * 1024)
                if not chunk:
                    break
                dst.write(chunk)
        print("    Descompresión GZIP completada ✓")

    # BAK directo
    else:
        print(f"    Formato detectado: BAK directo")
        if downloaded_path != final_bak_path:
            import shutil
            shutil.copy2(downloaded_path, final_bak_path)

    size = os.path.getsize(final_bak_path)
    print(f"    Archivo .bak listo: {size/1e6:.1f} MB")

def _move_largest_file(folder, dest):
    """Mueve el archivo más grande de una carpeta al destino."""
    import shutil
    files = [os.path.join(folder, f) for f in os.listdir(folder)
             if os.path.isfile(os.path.join(folder, f))]
    if not files:
        print(f"ERROR: No se encontraron archivos en {folder}")
        sys.exit(1)
    largest = max(files, key=os.path.getsize)
    print(f"    Archivo extraído: {os.path.basename(largest)}")
    shutil.move(largest, dest)

def get_logical_names(bak_path):
    """
    Obtiene los nombres lógicos de datos y log del .bak.
    Columnas sqlcmd: LogicalName | PhysicalName | Type | ...
    Índices:              0       |      1       |  2   |
    """
    out, code = sqlcmd(f"RESTORE FILELISTONLY FROM DISK=N'{bak_path}'")
    if code != 0:
        print(f"ERROR FILELISTONLY:\n{out}")
        sys.exit(1)
    data_name = log_name = None
    for line in out.splitlines():
        cols = line.split()
        # Saltar encabezado, separadores y líneas cortas
        if len(cols) < 3:
            continue
        if cols[0] in ('LogicalName', '---', '(2'):
            continue
        logical = cols[0].strip()
        file_type = cols[2].strip()   # D = datos, L = log
        if file_type == 'D' and not data_name:
            data_name = logical
        elif file_type == 'L' and not log_name:
            log_name = logical
    if not data_name or not log_name:
        print(f"ERROR: No se pudieron identificar los archivos lógicos.\nSalida:\n{out}")
        sys.exit(1)
    print(f"    Archivos lógicos: datos='{data_name}'  log='{log_name}'")
    return data_name, log_name

def restore_database(bak_path):
    """Restaura el .bak en SQL Server."""
    data_name, log_name = get_logical_names(bak_path)
    restore_q = f"""
RESTORE DATABASE [{DB_NAME}]
FROM DISK = N'{bak_path}'
WITH REPLACE,
     MOVE N'{data_name}' TO N'/var/opt/mssql/data/{DB_NAME}.mdf',
     MOVE N'{log_name}'  TO N'/var/opt/mssql/data/{DB_NAME}_log.ldf',
     STATS = 10
"""
    print(f"    Restaurando base de datos '{DB_NAME}'...")
    out, code = sqlcmd(restore_q)
    if code != 0 or 'RESTORE DATABASE successfully' not in out:
        print(f"ERROR al restaurar:\n{out}")
        sys.exit(1)
    print("    Base de datos restaurada ✓")

def export_query_to_csv(csv_path):
    """Ejecuta el query de ventas y exporta a CSV (separado por |)."""
    print(f"    Ejecutando query de ventas...")
    cmd = [
        SQLCMD, '-S', 'localhost', '-U', 'sa', '-P', os.environ['SA_PASS'],
        '-No', '-d', DB_NAME,
        '-s', '|',   # separador pipe (evita conflicto con ; en nombres)
        '-W',        # elimina espacios al final
        '-o', csv_path,
        '-Q', VENTAS_QUERY
    ]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        print(f"ERROR al exportar query:\n{r.stderr}")
        sys.exit(1)
    # Verificar que se generó el archivo
    size = os.path.getsize(csv_path)
    print(f"    CSV generado: {size/1e6:.1f} MB ✓")

def run_bak_mode():
    bak_url = os.environ.get('BAK_URL', '').strip()
    if not bak_url:
        print("ERROR: Define BAK_URL con la URL del archivo .bak")
        sys.exit(1)

    raw_path = "/tmp/autorefricentro_raw"
    print("[1/5] Descargando base de datos...")
    download_file(bak_url, raw_path)
    detect_and_extract_bak(raw_path, BAK_PATH)

    print("[2/5] Conectando a SQL Server...")
    wait_for_sqlserver()

    print("[2/5] Restaurando base de datos...")
    restore_database(BAK_PATH)

    print("[3/5] Exportando datos con query de ventas...")
    export_query_to_csv(CSV_PATH)

    print("[4/5] Procesando datos...")
    with open(CSV_PATH, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read()
    return parse_sql_output(content, delimiter='|')

# ═══════════════════════════════════════════════════════════════════════════════
#  MODO CSV — Descargar CSV exportado directamente del ERP
# ═══════════════════════════════════════════════════════════════════════════════

def run_csv_mode():
    csv_url = os.environ.get('CSV_URL', '').strip()
    if not csv_url:
        print("ERROR: Define CSV_URL o BAK_URL")
        sys.exit(1)
    print("[1/5] Descargando CSV del ERP...")
    req = urllib.request.Request(csv_url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=120) as r:
        content = r.read().decode('utf-8-sig', errors='replace')
    print(f"    {len(content.splitlines()):,} líneas descargadas ✓")
    print("[2-4/5] Procesando datos...")
    return parse_sql_output(content, delimiter=';')

# ═══════════════════════════════════════════════════════════════════════════════
#  PARSEO (aplica a ambos modos)
# ═══════════════════════════════════════════════════════════════════════════════

MONTH_NAMES = {
    'january':1,'february':2,'march':3,'april':4,'may':5,'june':6,
    'july':7,'august':8,'september':9,'october':10,'november':11,'december':12,
    'enero':1,'febrero':2,'marzo':3,'abril':4,'mayo':5,'junio':6,
    'julio':7,'agosto':8,'septiembre':9,'octubre':10,'noviembre':11,'diciembre':12,
}

def to_month_int(val):
    """Convierte nombre o número de mes a entero."""
    v = val.strip()
    if v.isdigit():
        return int(v)
    return MONTH_NAMES.get(v.lower(), 0)

def parse_num(s):
    s = s.strip().replace(' ', '').replace(',', '')
    try:
        return float(s)
    except ValueError:
        return 0.0

def parse_sql_output(content, delimiter=';'):
    """
    Parsea el output del query (o CSV V2).
    Detecta automáticamente el esquema de columnas.
    """
    lines = content.splitlines()

    # Buscar fila de encabezado
    col = {}
    header_idx = 0
    for i, line in enumerate(lines):
        parts = [p.strip() for p in line.split(delimiter)]
        # Encabezado del query SQL tiene 'ANO' o 'NOMBREVENDEDOR'
        if 'NOMBREVENDEDOR' in parts or 'ANO' in parts:
            col = {h.strip(): j for j, h in enumerate(parts)}
            header_idx = i
            break

    if not col:
        print("ERROR: No se encontró fila de encabezado con 'NOMBREVENDEDOR'")
        sys.exit(1)

    # Detectar esquema
    # Esquema SQL query:  ANO, MES (número), VENTAS, FACTURA
    # Esquema CSV V2:     FECHA - Year, FECHA - Month (nombre), VENTAS, CABECERA
    if 'ANO' in col:
        yr_col    = 'ANO'
        mo_col    = 'MES'
        sales_col = 'VENTAS'
        inv_col   = 'FACTURA'
        negate    = False          # query devuelve valores positivos
    else:
        yr_col    = 'FECHA - Year'
        mo_col    = 'FECHA - Month'
        sales_col = 'VENTAS'
        inv_col   = 'CABECERA'
        negate    = True           # CSV V2 tiene valores negativos

    print(f"    Esquema: {'SQL query' if 'ANO' in col else 'CSV V2'} | sep='{delimiter}'")

    monthly      = defaultdict(lambda: defaultdict(float))
    suc_monthly  = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    vend_monthly = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    vend_inv     = defaultdict(lambda: defaultdict(set))
    vend_cl      = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    cl_total     = defaultdict(lambda: defaultdict(float))

    reader = csv.reader(lines[header_idx + 1:], delimiter=delimiter)
    rows_ok = 0

    for row in reader:
        try:
            if not row or len(row) <= max(col.values()):
                continue
            vend    = row[col['NOMBREVENDEDOR']].strip()
            yr_str  = row[col[yr_col]].strip()
            mo_str  = row[col[mo_col]].strip()
            ventas  = row[col[sales_col]].strip()
            cl_name = row[col['CL_NOMBRE']].strip()
            suc     = row[col['SUCURSAL']].strip() if 'SUCURSAL' in col else ''
            inv_id  = row[col[inv_col]].strip()    if inv_col in col   else ''

            if vend.lower() in EXCLUDED:
                continue

            yr = int(yr_str)
            mo = to_month_int(mo_str)
            if mo == 0 or yr not in (2025, 2026):
                continue

            v = parse_num(ventas)
            if negate:
                v = abs(v)
            if v <= 0:
                continue

            monthly[yr][mo]            += v
            suc_monthly[suc][yr][mo]   += v
            vend_monthly[yr][vend][mo] += v
            vend_cl[yr][vend][cl_name] += v
            cl_total[yr][cl_name]      += v
            if inv_id:
                vend_inv[yr][vend].add(inv_id)
            rows_ok += 1
        except Exception:
            continue

    print(f"    {rows_ok:,} registros procesados ✓")
    return monthly, suc_monthly, vend_monthly, vend_inv, vend_cl, cl_total

# ═══════════════════════════════════════════════════════════════════════════════
#  CONSTRUCCIÓN DE DATOS PARA EL DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════

def detect_max_month(monthly):
    """Detecta el mes más alto con datos en cualquier año."""
    all_mo = list(monthly[2025].keys()) + list(monthly[2026].keys())
    return min(max(all_mo) if all_mo else 6, 12)

def build_monthly(monthly):
    max_mo = detect_max_month(monthly)
    result = {'_max_mo': max_mo}
    for yr in [2025, 2026]:
        arr = [int(monthly[yr].get(mo, 0)) for mo in range(1, max_mo + 1)]
        if yr == 2025:
            arr[0] = JAN_2025_TOTAL   # histórico hardcodeado
        result[yr] = arr
    return result

def build_sucs(suc_monthly, max_mo):
    sucs = []
    for key in SUC_ORDER:
        name, short, cls = SUC_DEF[key]
        m25 = [int(suc_monthly[key][2025].get(mo, 0)) for mo in range(1, max_mo + 1)]
        m26 = [int(suc_monthly[key][2026].get(mo, 0)) for mo in range(1, max_mo + 1)]
        if key == 'TIENDA PRINCIPAL-SD':        m25[0] = JAN_2025_PRINCIPAL
        if key == 'SUCURSAL SANTIAGO':           m25[0] = JAN_2025_SANTIAGO
        if key == 'SUCURSAL ZONA ORIENTAL-SD':   m25[0] = JAN_2025_ORIENTAL
        sucs.append({'name': name, 'short': short, 'cls': cls, 'm25': m25, 'm26': m26})
    # Campo — solo Ene 2025 (histórico)
    campo_m25 = [0] * max_mo
    campo_m25[0] = JAN_2025_CAMPO
    sucs.append({'name': 'Sin Sucursal/Campo', 'short': 'Sin Sucursal', 'cls': 's4',
                 'm25': campo_m25, 'm26': [0] * max_mo})
    return sucs

def short_name(full):
    p = full.strip().split()
    return p[0] + (' ' + p[1][0] + '.' if len(p) > 1 else '') if p else full

def build_vendors(vend_monthly, vend_inv, vend_cl, max_mo):
    vendors = {}
    for yr in [2025, 2026]:
        ranked = sorted(vend_monthly[yr].items(), key=lambda x: -sum(x[1].values()))
        lst = []
        for vend, months in ranked:
            total = int(sum(months.values()))
            if total < 100:
                continue
            lst.append({
                'n':  vend,
                'sn': short_name(vend),
                't':  total,
                'f':  len(vend_inv[yr][vend]),
                'm':  [int(months.get(mo, 0)) for mo in range(1, max_mo + 1)],
                'cl': [[n, int(t)] for n, t in
                       sorted(vend_cl[yr][vend].items(), key=lambda x: -x[1])[:10]],
            })
        vendors[yr] = lst
    return vendors

def build_top_cl(cl_total):
    return {
        yr: [[n, int(t)] for n, t in sorted(cl_total[yr].items(), key=lambda x: -x[1])[:20]]
        for yr in [2025, 2026]
    }

# ═══════════════════════════════════════════════════════════════════════════════
#  GENERACIÓN DE JS E INYECCIÓN EN HTML
# ═══════════════════════════════════════════════════════════════════════════════

def js_arr(lst):
    return '[' + ','.join(str(x) for x in lst) + ']'

def js_s(s):
    """
    Escapa un string para uso seguro en JavaScript.
    Usa json.dumps para manejar TODOS los caracteres especiales
    (apóstrofes, tildes, saltos de línea, backslashes, comillas, etc.)
    Devuelve un string entre comillas dobles válido en JS.
    """
    return json.dumps(str(s).strip(), ensure_ascii=False)

def render_js(md, sucs, vendors, top_cl):
    now = datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')
    max_mo = md['_max_mo']
    parts = []

    # MN — etiquetas de meses dinámicas
    mn_labels = MONTH_LABELS_ES[:max_mo]
    mn_js = '[' + ','.join(json.dumps(m, ensure_ascii=False) for m in mn_labels) + ']'
    parts.append(f"const MN = {mn_js};")

    # DATE_RANGE — rango de fechas para el badge
    last_mo_label = MONTH_LABELS_ES[max_mo - 1]
    # Si 2026 tiene datos usa 2026, sino 2025
    last_yr = 2026 if any(md[2026]) else 2025
    date_range = f"Datos: 1 Ene 2025 – 30 {last_mo_label} {last_yr}"
    parts.append(f"const DATE_RANGE = {json.dumps(date_range, ensure_ascii=False)};")

    # MONTHLY
    parts.append(
        f"const MONTHLY = {{\n"
        f"  2025: {js_arr(md[2025])},\n"
        f"  2026: {js_arr(md[2026])}\n"
        f"}};"
    )

    # SUCS
    suc_rows = []
    for s in sucs:
        suc_rows.append(
            f"  {{name:{js_s(s['name'])},short:{js_s(s['short'])},cls:{js_s(s['cls'])},"
            f"m25:{js_arr(s['m25'])},m26:{js_arr(s['m26'])}}}"
        )
    parts.append("const SUCS = [\n" + ',\n'.join(suc_rows) + "\n];")

    # VENDORS
    def fmt_v(v):
        cl = '[' + ','.join(f"[{js_s(c[0])},{c[1]}]" for c in v['cl']) + ']'
        return (f"  {{n:{js_s(v['n'])},sn:{js_s(v['sn'])},t:{v['t']},"
                f"f:{v['f']},m:{js_arr(v['m'])},cl:{cl}}}")

    v25 = ',\n'.join(fmt_v(v) for v in vendors[2025])
    v26 = ',\n'.join(fmt_v(v) for v in vendors[2026])
    parts.append(f"const VENDORS = {{\n  2025:[\n{v25}\n  ],\n  2026:[\n{v26}\n  ]\n}};")

    # TOP_CL
    cl25 = '[' + ','.join(f"[{js_s(c[0])},{c[1]}]" for c in top_cl[2025]) + ']'
    cl26 = '[' + ','.join(f"[{js_s(c[0])},{c[1]}]" for c in top_cl[2026]) + ']'
    parts.append(f"const TOP_CL = {{\n  2025:{cl25},\n  2026:{cl26}\n}};")

    # Timestamp
    parts.append(f"const LAST_UPDATED = '{now}';")

    return '\n'.join(parts)

def inject_html(html, new_js):
    START = '// ==DATA-START=='
    END   = '// ==DATA-END=='
    pattern = re.compile(re.escape(START) + r'.*?' + re.escape(END), re.DOTALL)
    new_html, n = pattern.subn(START + '\n' + new_js + '\n' + END, html)
    if n == 0:
        print(f"ERROR: Marcadores '{START}' / '{END}' no encontrados en {HTML_FILE}")
        sys.exit(1)
    return new_html

# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    # Detectar modo
    if os.environ.get('BAK_URL'):
        print("=== AutoRefricentro MYB — Modo: BAK → SQL Server ===\n")
        result = run_bak_mode()
    elif os.environ.get('CSV_URL'):
        print("=== AutoRefricentro MYB — Modo: CSV ===\n")
        result = run_csv_mode()
    else:
        print("ERROR: Define BAK_URL (recomendado) o CSV_URL")
        sys.exit(1)

    monthly_raw, suc_monthly, vend_monthly, vend_inv, vend_cl, cl_total = result

    # Construir estructuras
    md       = build_monthly(monthly_raw)
    max_mo   = md['_max_mo']
    sucs     = build_sucs(suc_monthly, max_mo)
    vendors  = build_vendors(vend_monthly, vend_inv, vend_cl, max_mo)
    top_cl   = build_top_cl(cl_total)

    print(f"    Meses detectados: {max_mo} ({MONTH_LABELS_ES[max_mo-1]} = último mes con datos)")

    # Resumen
    print("\n    Resumen de datos:")
    for yr in [2025, 2026]:
        t = sum(md[yr])
        print(f"      {yr}: RD${t/1e6:.1f}M | {len(vendors[yr])} vendedores")

    # Generar JS
    new_js = render_js(md, sucs, vendors, top_cl)

    # Actualizar HTML
    print(f"\n[5/5] Actualizando {HTML_FILE}...")
    with open(HTML_FILE, 'r', encoding='utf-8') as f:
        html = f.read()
    with open(HTML_FILE, 'w', encoding='utf-8') as f:
        f.write(inject_html(html, new_js))

    ts = datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')
    print(f"\n✅  {HTML_FILE} actualizado exitosamente — {ts}")

if __name__ == '__main__':
    main()
