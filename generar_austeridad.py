"""
GENERADOR FORMATO AUSTERIDAD — Art. 10 LFAR — SADER
Fuentes: Cuenta Pública cierre 2025 (CSV) + SICOP corte actual 2026 (CSV)

USO EN GOOGLE COLAB:
    !python generar_austeridad.py
"""
import importlib, subprocess, sys as _sys
def _ensure(pkg):
    try: importlib.import_module(pkg)
    except ImportError:
        print(f"  📦 Instalando {pkg}...")
        subprocess.check_call([_sys.executable, "-m", "pip", "install", pkg, "-q"])
_ensure("openpyxl"); _ensure("pandas")

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import re, os
from datetime import datetime

# ══════════════════════════════════════════════════════════════════════
#  VALORES FIJOS CP2025
#  Fuente: Cuenta Pública cierre 2025, Sector Central (URs numéricas 3
#  dígitos + G00), columna EJERCIDO.
#  Estos valores no cambian — la CP es un cierre definitivo.
#  Split 32301: UR 513 = Arrendamientos | resto = Fotocopiado
#  33602 Fotocopiado = 0.020000 (dato manual reportado por DGRMIS/UR512)
# ══════════════════════════════════════════════════════════════════════
CP2025 = {
    22102: 0.000000,
    22104: 3.068149,
    22106: 0.207542,
    38501: 0.000000,
    # 32301 split:
    ('32301','arrend'): 26.712381,   # UR 513 exacto del formato referencia
    ('32301','foto'):   14.482541,   # resto URs exacto del formato referencia
    32302: 0.000000,
    32303: 0.000000,
    32201: 37.288754,
    32503: 20.781844,
    32505: 0.000000,
    32601:  4.376538,
    32701:  8.688093,
    32903:  0.013317,
    21401:  2.131185,
    35301:  0.360002,
    51501:  0.000000,
    26102: 18.407324,
    26103: 25.902420,
    26104:  3.502343,
    26105:  0.000000,
    38301:  0.000000,
    31101: 34.094310,
    ('33602','foto'):   0.020000,    # manual DGRMIS
    31603:  0.118319,
    31701: 80.406607,
    35201:  3.349913,
    21101:  2.797589,
    37101:  0.000000,
    37104:  5.660503,
    37106:  0.073745,
    37201:  0.099702,
    37204:  0.232062,
    37206:  0.000000,
    35101: 42.557543,
    62202:  0.000000,
    31401:  7.090882,
    37501:  0.159059,
    37504:  3.921413,
    37602:  0.038895,
}

# ══════════════════════════════════════════════════════════════════════
#  MAPA EXACTO DE FILAS
#  (fila_excel, tipo, label_A, partida_key)
# ══════════════════════════════════════════════════════════════════════
MAPA = [
    (13,'grupo', 'Alimentación',            None),
    (14,'part',  None,                       22102),
    (15,'part',  None,                       22104),
    (16,'part',  None,                       22106),
    (17,'part',  None,                       38501),
    (18,'grupo', 'Arrendamientos',           None),
    (19,'part',  None,                       32201),
    (20,'part',  None,                       ('32301','arrend')),
    (21,'part',  None,                       32302),
    (22,'part',  None,                       32303),
    (23,'part',  None,                       32503),
    (24,'part',  None,                       32505),
    (25,'part',  None,                       32601),
    (26,'part',  None,                       32701),
    (27,'part',  None,                       32903),
    (28,'grupo', 'Bienes informáticos',      None),
    (29,'part',  None,                       21401),
    (30,'part',  None,                       35301),
    (31,'part',  None,                       51501),
    (32,'grupo', 'Combustibles',             None),
    (33,'part',  None,                       26102),
    (34,'part',  None,                       26103),
    (35,'part',  None,                       26104),
    (36,'part',  None,                       26105),
    (37,'grupo', 'Congresos',                None),
    (38,'part',  None,                       38301),
    (39,'grupo', 'Energía eléctrica',        None),
    (40,'part',  None,                       31101),
    (41,'grupo', 'Fotocopiado **',           None),
    (42,'part',  None,                       ('32301','foto')),
    (43,'part',  None,                       ('33602','foto')),
    (44,'grupo', 'Internet',                 None),
    (45,'part',  None,                       31603),
    (46,'part',  None,                       31701),
    (47,'grupo', 'Mobiliario',               None),
    (48,'part',  None,                       35201),
    (49,'grupo', 'Papelería',                None),
    (50,'part',  None,                       21101),
    (51,'grupo', 'Pasajes',                  None),
    (52,'part',  None,                       37101),
    (53,'part',  None,                       37104),
    (54,'part',  None,                       37106),
    (55,'part',  None,                       37201),
    (56,'part',  None,                       37204),
    (57,'part',  None,                       37206),
    (58,'grupo', 'Remodelación de oficinas', None),
    (59,'part',  None,                       35101),
    (60,'part',  None,                       62202),
    (61,'grupo', 'Telefonía',                None),
    (62,'part',  None,                       31401),
    (63,'grupo', 'Viáticos',                 None),
    (64,'part',  None,                       37501),
    (65,'part',  None,                       37504),
    (66,'part',  None,                       37602),
]

# Partidas hijas de cada grupo (para calcular subtotal)
GRUPOS_PARTES = {
    13: [14,15,16,17],
    18: [19,20,21,22,23,24,25,26,27],
    28: [29,30,31],
    32: [33,34,35,36],
    37: [38],
    39: [40],
    41: [42,43],
    44: [45,46],
    47: [48],
    49: [50],
    51: [52,53,54,55,56,57],
    58: [59,60],
    61: [62],
    63: [64,65,66],
}
FILAS_GRUPO = list(GRUPOS_PARTES.keys())

# Número de partida mostrado en col B (para claves tuple)
PARTIDA_NUM = {
    ('32301','arrend'): 32301,
    ('32301','foto'):   32301,
    ('33602','foto'):   33602,
}

DENOMINACIONES = {
    22102: 'Productos alimenticios para personas derivado de la prestación de servicios públicos de carácter social y operativos',
    22104: 'Productos alimenticios para el personal en las instalaciones de las dependencias y entidades',
    22106: 'Productos alimenticios para el personal derivado de actividades extraordinarias o que presten servicios en turnos para atención continua',
    38501: 'Gastos para alimentación de servidores públicos de mando',
    32201: 'Arrendamiento de edificios y locales',
    32301: 'Arrendamiento de equipo y bienes informáticos',
    32302: 'Arrendamiento de mobiliario',
    32303: 'Arrendamiento de equipo de telecomunicaciones',
    32503: 'Arrendamiento de vehículos terrestres, aéreos, marítimos, lacustres y fluviales para servidores públicos de mando',
    32505: 'Arrendamiento de vehículos terrestres, aéreos, marítimos, lacustres y fluviales para labores en campo y de supervisión',
    32601: 'Arrendamiento de maquinaria y equipo',
    32701: 'Patentes, derechos de autor, regalías y otros',
    32903: 'Otros Arrendamientos',
    21401: 'Materiales y útiles consumibles para el procesamiento en equipos y bienes informáticos',
    35301: 'Mantenimiento y conservación de bienes informáticos',
    51501: 'Bienes informáticos',
    26102: 'Combustibles, lubricantes y aditivos para vehículos terrestres, aéreos, marítimos, lacustres y fluviales destinados a servicios públicos y la operación de programas públicos',
    26103: 'Combustibles, lubricantes y aditivos para vehículos terrestres, aéreos, marítimos, lacustres y fluviales destinados a servicios administrativos',
    26104: 'Combustibles, lubricantes y aditivos para vehículos terrestres, aéreos, marítimos, lacustres y fluviales asignados a servidores públicos',
    26105: 'Combustibles, lubricantes y aditivos para maquinaria, equipo de producción y servicios especializados',
    38301: 'Congresos y convenciones',
    31101: 'Servicio de energía eléctrica',
    33602: 'Otros servicios comerciales',
    31603: 'Servicios de internet',
    31701: 'Servicios de conducción de señales analógicas y digitales',
    35201: 'Mantenimiento y conservación de mobiliario y equipo de administración',
    21101: 'Materiales y útiles de oficina',
    37101: 'Pasajes aéreos nacionales para labores en campo y de supervisión',
    37104: 'Pasajes aéreos nacionales para servidores públicos de mando en el desempeño de comisiones y funciones oficiales',
    37106: 'Pasajes aéreos internacionales para servidores públicos en el desempeño de comisiones y funciones oficiales',
    37201: 'Pasajes terrestres nacionales para labores en campo y de supervisión',
    37204: 'Pasajes terrestres nacionales para servidores públicos de mando en el desempeño de comisiones y funciones oficiales',
    37206: 'Pasajes terrestres internacionales para servidores públicos en el desempeño de comisiones y funciones oficiales',
    35101: 'Mantenimiento y conservación de inmuebles para la prestación de servicios administrativos',
    62202: 'Mantenimiento y rehabilitación de edificaciones no habitacionales',
    31401: 'Servicio telefónico convencional',
    37501: 'Viáticos nacionales para labores en campo y de supervisión',
    37504: 'Viáticos nacionales para servidores públicos en el desempeño de funciones oficiales',
    37602: 'Viáticos en el extranjero para servidores públicos en el desempeño de comisiones y funciones oficiales',
}

# ══════════════════════════════════════════════════════════════════════
#  LECTURA SICOP 2026
# ══════════════════════════════════════════════════════════════════════

def _leer_csv(path):
    try:    return pd.read_csv(path, encoding='utf-8', low_memory=False)
    except: return pd.read_csv(path, encoding='latin-1', low_memory=False)

def leer_sicop2026(path):
    """SICOP 2026 — SC + G00 — EJERCIDO + EJERCIDO_TRAMITE"""
    print("  📊 Leyendo SICOP 2026...")
    df = _leer_csv(path)
    df = df.copy()
    df['P5']     = (df['CAPITULO'].astype(int)*10000 + df['CONCEPTO'].astype(int)*1000 +
                    df['PARTIDA_GENERICA'].astype(int)*100 + df['PARTIDA_ESPECIFICA'].astype(int))
    df['SC_G00'] = df['ID_UNIDAD'].astype(str).str.match(r'^\d{3}$') | (df['ID_UNIDAD']=='G00')
    sc = df[df['SC_G00']].copy()
    sc['EJ'] = sc['EJERCIDO'] + sc['EJERCIDO_TRAMITE']

    base = dict(sc.groupby('P5')['EJ'].sum() / 1_000_000)

    # Split 32301
    p32301 = sc[sc['P5']==32301]
    base[('32301','arrend')] = p32301[p32301['ID_UNIDAD']=='513']['EJ'].sum() / 1_000_000
    base[('32301','foto')]   = p32301[p32301['ID_UNIDAD']!='513']['EJ'].sum() / 1_000_000
    # 33602 foto = solo UR 512 (DGRMIS)
    p33602 = sc[sc['P5']==33602]
    base[('33602','foto')]   = p33602[p33602['ID_UNIDAD']=='512']['EJ'].sum() / 1_000_000

    print(f"  ✅ SICOP 2026: {len(sc['P5'].unique())} partidas (SC + G00)")
    return base

def detectar_corte(path):
    MESES = {'ENE':1,'FEB':2,'MAR':3,'ABR':4,'MAY':5,'JUN':6,
             'JUL':7,'AGO':8,'SEP':9,'OCT':10,'NOV':11,'DIC':12,
             'ENERO':1,'FEBRERO':2,'MARZO':3,'ABRIL':4,'MAYO':5,'JUNIO':6,
             'JULIO':7,'AGOSTO':8,'SEPTIEMBRE':9,'OCTUBRE':10,'NOVIEMBRE':11,'DICIEMBRE':12}
    DIAS  = {1:31,2:28,3:31,4:30,5:31,6:30,7:31,8:31,9:30,10:31,11:30,12:31}
    MESES_ES = {1:'enero',2:'febrero',3:'marzo',4:'abril',5:'mayo',6:'junio',
                7:'julio',8:'agosto',9:'septiembre',10:'octubre',11:'noviembre',12:'diciembre'}
    ABR3  = {1:'ENE',2:'FEB',3:'MAR',4:'ABR',5:'MAY',6:'JUN',
             7:'JUL',8:'AGO',9:'SEP',10:'OCT',11:'NOV',12:'DIC'}
    nom   = os.path.basename(path).upper()
    m     = re.search(r'(\d{1,2})[-_\s]([A-Z]+)[-_\s](\d{4})', nom)
    if m:
        mn = MESES.get(m.group(2))
        if mn:
            return {'dia':int(m.group(1)),'mes':mn,'anio':int(m.group(3)),
                    'mes_es':MESES_ES[mn],'abr3':ABR3[mn]}
    hoy = datetime.today()
    return {'dia':DIAS[hoy.month],'mes':hoy.month,'anio':hoy.year,
            'mes_es':MESES_ES[hoy.month],'abr3':ABR3[hoy.month]}

# ══════════════════════════════════════════════════════════════════════
#  ESTILOS (exactos del formato de referencia)
# ══════════════════════════════════════════════════════════════════════
ENC_FILL = 'A50021'
FMT_NUM  = '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-'
FMT_PCT  = '0.00%'
FONT     = 'Calibri'
SZ       = 11

def _f(bold=False): return Font(name=FONT, size=SZ, bold=bold)
def _fw(bold=False): return Font(name=FONT, size=SZ, bold=bold, color='FFFFFF')
def _fill(h): return PatternFill('solid', fgColor=h)
def _a(h='general', v='top', wrap=False): return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _enc(c, val, ha='center', va='center', nf='General', wrap=False):
    c.value = val; c.fill = _fill(ENC_FILL)
    c.font  = _fw(True)
    c.alignment = Alignment(horizontal=ha, vertical=va, wrap_text=wrap)
    c.number_format = nf

def _num(c, val, bold=False, nf=FMT_NUM):
    c.value = val; c.font = _f(bold)
    c.alignment = _a('center', 'top'); c.number_format = nf

def _pct(c, val, bold=False):
    c.value = val; c.font = _f(bold)
    c.alignment = _a('center', 'top'); c.number_format = FMT_PCT

# ══════════════════════════════════════════════════════════════════════
#  GENERACIÓN
# ══════════════════════════════════════════════════════════════════════

def generar_austeridad(ruta_sicop, ruta_salida=None):
    """
    Genera el Formato de Austeridad.
    ruta_sicop: CSV del SICOP del corte actual (2026)
    Los valores de CP2025 están embebidos y son fijos.
    """
    ej26  = leer_sicop2026(ruta_sicop)
    corte = detectar_corte(ruta_sicop)
    d, m_es, abr3, anio = corte['dia'], corte['mes_es'], corte['abr3'], corte['anio']
    print(f"  📅 Corte: {d} de {m_es} de {anio}")

    # Construir tabla completa con valores D (2025 fijo) y E (2026 dinámico)
    # y calcular subtotales de grupo directamente
    fila_vals = {}   # fila → (d25, e26)
    for fila, tipo, label, key in MAPA:
        if tipo == 'part':
            d25 = CP2025.get(key, 0.0)
            e26 = ej26.get(key, 0.0)
            fila_vals[fila] = (d25, e26)

    # Subtotales de grupo = suma de sus partidas hijas
    grupo_vals = {}
    for g_fila, hijas in GRUPOS_PARTES.items():
        d25 = sum(fila_vals.get(h, (0,0))[0] for h in hijas)
        e26 = sum(fila_vals.get(h, (0,0))[1] for h in hijas)
        grupo_vals[g_fila] = (d25, e26)

    # Total general
    tot25 = sum(v[0] for v in grupo_vals.values())
    tot26 = sum(v[1] for v in grupo_vals.values())

    # ── Crear workbook ────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"AUSTERIDAD DGPPF CIERRE {abr3}"

    # Anchos exactos del formato de referencia
    for col, w in [('A',9.43),('B',7.29),('C',67.57),('D',15.14),
                   ('E',15.14),('F',15.14),('G',20.0),('H',11.43)]:
        ws.column_dimensions[col].width = w

    # Altos de fila
    for r, h in [(3,19.5),(4,19.5),(5,15.0),(8,15.75),(9,25.5),(10,6.75),(12,6.75)]:
        ws.row_dimensions[r].height = h

    # ── Títulos ──────────────────────────────────────────────────────
    ws.merge_cells('A3:G3')
    ws['A3'].value = 'SECRETARÍA DE AGRICULTURA Y DESARROLLO RURAL'
    ws['A3'].font  = Font(name=FONT, size=15, bold=True)
    ws['A3'].alignment = _a('center','center')

    ws.merge_cells('A4:G4')
    ws['A4'].value = 'CONCEPTOS SUJETOS A AUSTERIDAD Y RACIONALIDAD'
    ws['A4'].font  = Font(name=FONT, size=15, bold=True)
    ws['A4'].alignment = _a('center','center')

    ws.merge_cells('A5:G6')
    ws['A5'].value = ('(Artículo 10 Ley Federal de Austeridad Republicana y numeral '
                      '10 fracción II de \nlos Lineamientos en Materia de Austeridad '
                      'Republicana de la Administración Pública Federal)')
    ws['A5'].font  = Font(name=FONT, size=SZ, bold=True)
    ws['A5'].alignment = _a('center','center',wrap=True)

    # ── Encabezados (filas 8-9) ───────────────────────────────────────
    ws.merge_cells('A8:A9'); ws.merge_cells('B8:B9'); ws.merge_cells('C8:C9')
    ws.merge_cells('D8:E8'); ws.merge_cells('F8:F9'); ws.merge_cells('G8:G9')

    _enc(ws['A8'], 'Concepto')
    _enc(ws['B8'], 'Partida')
    _enc(ws['C8'], 'Denominación partida')
    _enc(ws['D8'], 'Ejercido*')
    _enc(ws['F8'], 'Diferencia', nf=FMT_NUM)
    _enc(ws['G8'], f'  {anio} vs {anio-1} (variación porcentual)', nf=FMT_NUM, wrap=True)
    _enc(ws['D9'], f'{anio-1}/1')
    _enc(ws['E9'], f'{anio}/2')

    # ── Fila 11: Total general ────────────────────────────────────────
    ws.row_dimensions[11].height = 15.75
    ws['A11'].value = 'Total general'
    ws['A11'].font  = _f(True)
    _num(ws['D11'], tot25, bold=True)
    _num(ws['E11'], tot26, bold=True)
    _num(ws['F11'], tot25 - tot26, bold=True)
    _pct(ws['G11'], (tot26 - tot25) / tot25 if tot25 else 0, bold=True)

    # ── Grupos y partidas ─────────────────────────────────────────────
    fila_map = {f: (t, la, k) for f, t, la, k in MAPA}

    for fila, tipo, label, key in MAPA:
        ws.row_dimensions[fila].height = 15.75 if tipo == 'grupo' else 30.0

        if tipo == 'grupo':
            d25, e26 = grupo_vals[fila]
            ws[f'A{fila}'].value = label
            ws[f'A{fila}'].font  = _f(True)
            ws[f'A{fila}'].alignment = _a('left','top')
            _num(ws[f'D{fila}'], d25, bold=True)
            _num(ws[f'E{fila}'], e26, bold=True)
            _num(ws[f'F{fila}'], d25 - e26, bold=True)
            _pct(ws[f'G{fila}'], (e26 - d25) / d25 if d25 else 0, bold=True)

        elif tipo == 'part':
            d25, e26 = fila_vals[fila]
            # Col B: número de partida
            p_num = key if isinstance(key, int) else PARTIDA_NUM.get(key, 0)
            ws[f'B{fila}'].value = p_num
            ws[f'B{fila}'].font  = _f()
            ws[f'B{fila}'].alignment = _a('left' if fila in (42,43) else 'center', 'top')
            # Col C: denominación
            ws[f'C{fila}'].value = DENOMINACIONES.get(p_num, '')
            ws[f'C{fila}'].font  = _f()
            ws[f'C{fila}'].alignment = _a('left','top',wrap=True)
            # Valores
            _num(ws[f'D{fila}'], d25)
            _num(ws[f'E{fila}'], e26)
            _num(ws[f'F{fila}'], d25 - e26)
            if d25 != 0:
                _pct(ws[f'G{fila}'], (e26 - d25) / d25)
            else:
                ws[f'G{fila}'].value = 0
                ws[f'G{fila}'].font  = _f()
                ws[f'G{fila}'].alignment = _a('center','top')
                ws[f'G{fila}'].number_format = FMT_PCT

    # ── Notas (posición exacta del formato de referencia) ─────────────
    ws['A69'].value = 'Nota: '; ws['A69'].font = _f(True); ws['A69'].alignment = _a('left','top')

    ws['A70'].value = '*';  ws['A70'].font = _f(); ws['A70'].alignment = _a('right','top')
    ws['B70'].value = ('La información presentada considera Sector Central, '
                       'incluyendo las Oficinas de Representación en las Entidades Federativas.')
    ws['B70'].font = _f(); ws['B70'].alignment = _a('left','top')

    ws['B71'].value = ('La información referente al servicio de Fotocopiado deberá ser '
                       'solicitada a la Dirección General de Recursos Materiales, '
                       'Inmuebles y Servicios (Ur 512)')
    ws['B71'].font = _f(); ws['B71'].alignment = _a('left','top')

    ws['A72'].value = '**'; ws['A72'].font = _f(); ws['A72'].alignment = _a('right','top')
    ws.merge_cells('B72:G72')
    ws['B72'].value = ('Para el caso de las partidas 32301 "Arrendamiento de equipo y '
                       'bienes informáticos" y 33602 "Otros servicios comerciales" no es '
                       'posible desagregar de los sistemas hacendarios el gasto del servicio '
                       'de Fotocopiado, toda vez que en estas partidas se cubren diversas '
                       'erogaciones, por lo que la información debiera corresponder a lo '
                       'reportado por las Unidades Responsables ejecutoras del gasto '
                       '(Dirección General de Recursos Materiales, Inmuebles y Servicios '
                       'y las OREF).')
    ws['B72'].font = _f(); ws['B72'].alignment = _a('left','top',wrap=True)
    ws.row_dimensions[72].height = 45.0

    ws['A75'].value = 'Fuente:'; ws['A75'].font = _f(True); ws['A75'].alignment = _a('left','top')
    ws.merge_cells('B76:G76')
    ws['B76'].value = (f'/1. Sistema de Contabilidad y Presupuesto del ejercicio '
                       f'{anio-1} (base para Cuenta Pública).')
    ws['B76'].font = _f()
    ws.merge_cells('B77:G77')
    ws['B77'].value = (f'/2. Sistema de Contabilidad y Presupuesto del ejercicio '
                       f'{anio}, corte al {d} de {m_es} de {anio}.')
    ws['B77'].font = _f()

    # ── Guardar ───────────────────────────────────────────────────────
    if ruta_salida is None:
        ruta_salida = os.path.join(os.getcwd(),
            f"FORMATO_AUSTERIDAD_{abr3}-{str(anio)[2:]}.xlsx")
    wb.save(ruta_salida)
    print(f"\n✅ Archivo generado: {ruta_salida}")
    return ruta_salida

# ══════════════════════════════════════════════════════════════════════
#  COLAB / CLI
# ══════════════════════════════════════════════════════════════════════
def _colab_main():
    try:
        from google.colab import files as _f_
    except ImportError:
        args = _sys.argv[1:]
        if args:
            generar_austeridad(args[0])
        else:
            sc = input('Ruta CSV SICOP corte 2026: ').strip()
            if sc: generar_austeridad(sc)
        return

    print('═'*55)
    print('  FORMATO AUSTERIDAD — SADER')
    print('═'*55)
    print('\n📂 Sube el CSV del SICOP corte actual 2026:')
    sub = _f_.upload()
    if not sub: return
    n, c = next(iter(sub.items()))
    r = f'/content/{n}'
    with open(r,'wb') as f: f.write(c)
    print(f'   ✅ {n}')
    rout = generar_austeridad(r)
    print('\n⬇️  Descargando...')
    _f_.download(rout)
    print('✅ ¡Listo!')

if __name__ == '__main__':
    _colab_main()
