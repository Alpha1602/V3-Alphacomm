import openpyxl, json, os

# ── Archivos fuente ───────────────────────────────────────────────────────────
EXCEL    = "/sessions/tender-amazing-curie/mnt/uploads/Sales YoY .3.xlsx"
PRIMEMX  = "/sessions/tender-amazing-curie/mnt/uploads/Alphacomm_PrimeMX_2026-10-01.xlsx"

wb  = openpyxl.load_workbook(EXCEL,   read_only=True, data_only=True)
wb2 = openpyxl.load_workbook(PRIMEMX, read_only=True, data_only=True)

def get_sheet(*names):
    for n in names:
        if n in wb.sheetnames: return wb[n]
    raise KeyError(f"None found: {names}")

def s(v):   return '' if v is None else str(v).strip()
def flt(v):
    try: return float(str(v).replace('%','').strip())
    except: return 0.0
def fi(v):  return int(round(flt(v)))

MO = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,
      'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
mo_list = list(MO.keys())

# ── 1. CUOTA MES (Sales YoY) ──────────────────────────────────────────────────
sh = get_sheet('1.Cuota Mes', 'Cuota Jul')
rows = list(sh.iter_rows(values_only=True))
cuota_jul = []
for r in rows[3:]:
    reg = s(r[0]); clv = s(r[2])
    if not reg or not clv: continue
    alc = flt(r[6]); alc = round(alc*100,1) if alc < 2 else round(alc,1)
    cuota_jul.append({'region':reg,'sub_region':s(r[1]),'clave':clv,'tienda':s(r[3]),
        'cuota':fi(r[4]),'ventas':fi(r[5]),'proy':0,'alcance':alc,
        'oh_total':0,'oh_cases':0,'oh_cables':0,'oh_chargers':0,'oh_liquid':0,'oh_micas':0,
        'resurtido':fi(r[8]),'faltante':fi(r[9]),'riesgo':s(r[10]) or 'Sin riesgo'})
print(f"Cuota Mes: {len(cuota_jul)} tiendas")

# ── 2. PROYECCIÓN TIENDA ──────────────────────────────────────────────────────
pm = {}
try:
    sh_pt = get_sheet('2.Proyección Tienda', 'Proyección Tienda')
    for r in sh_pt.iter_rows(values_only=True, min_row=4):
        clv = s(r[3])
        if clv: pm[clv] = round(flt(r[8]),1)
except: pass
if not pm:
    for r in get_sheet('3.AR','1.AR').iter_rows(values_only=True, min_row=4):
        clv = s(r[3])
        if clv: pm[clv] = round(flt(r[10]),1)
for c in cuota_jul:
    c['proy'] = pm.get(c['clave'], c['ventas'])
print(f"Proy: {len(pm)} tiendas mapeadas")

# ── 3. AR CLEAN — PrimeMX "Detalle por tienda" ───────────────────────────────
# Cols: 0=Región, 1=Regional, 2=Clave PDV, 3=Nombre PDV,
#       4=Ventas Alphacomm, 5=Total Accesorios, 6=% Alphacomm,
#       7=Venta de equipo, 8=Alphacomm/Equipo
ar_clean = []
for r in wb2['Detalle por tienda'].iter_rows(values_only=True, min_row=2):
    reg = s(r[0]); clv = s(r[2])
    if not reg or not clv: continue
    ar_clean.append({
        'region':     reg,
        'gerente':    s(r[1]),
        'clave':      clv,
        'tienda':     s(r[3]),
        'alphacomm':  fi(r[4]),
        'total_acc':  fi(r[5]),
        'pct_alpha':  round(flt(r[6]),1),
        'equipo':     fi(r[7]),
        'ar_ratio':   round(flt(r[8]),2),
        'proy':       0,
    })
print(f"AR (PrimeMX): {len(ar_clean)} tiendas")

# ── 3b. RESUMEN NACIONAL PrimeMX ─────────────────────────────────────────────
res = {}
for r in wb2['Resumen Nacional'].iter_rows(values_only=True, min_row=2):
    if r[0] and r[1] is not None: res[s(r[0])] = r[1]

primemx_summary = {
    'periodo':    s(res.get('Periodo','')),
    'alpha_units':fi(res.get('Ventas Alphacomm (unidades)',0)),
    'total_acc':  fi(res.get('Total accesorios canal',0)),
    'pct_alpha':  round(flt(res.get('% Alphacomm del total',0)),1),
    'equipo':     fi(res.get('Venta de equipo (unidades)',0)),
    'ar_ratio':   round(flt(res.get('Alphacomm / equipo',0)),2),
    'pdvs_con':   fi(res.get('PDVs con venta Alphacomm',0)),
    'pdvs_sin':   fi(res.get('PDVs sin venta Alphacomm',0)),
}

# ── 3c. POR REGIÓN PrimeMX ───────────────────────────────────────────────────
ar_region = []
for r in wb2['Por región'].iter_rows(values_only=True, min_row=2):
    reg = s(r[0])
    if not reg or reg == 'SIN ASIGNAR': continue
    ar_region.append({
        'region':    reg,
        'alphacomm': fi(r[1]),
        'total_acc': fi(r[2]),
        'pct_alpha': round(flt(r[3]),1),
        'equipo':    fi(r[4]),
        'ar_ratio':  round(flt(r[5]),2),
    })
print(f"AR por región: {len(ar_region)} regiones  |  Periodo: {primemx_summary['periodo']}")

# ── 4. OH DETALLE ─────────────────────────────────────────────────────────────
sh4 = get_sheet('5.OH detalle', 'OH detalle')
t_map = {}; p_map = {}; r_map = {}; reg_set4 = {}

for r in sh4.iter_rows(values_only=True, min_row=4):
    reg = s(r[3]); clv = s(r[1]); tda = s(r[0])
    cat_alpha = s(r[5]) or 'Accesories'
    prod = s(r[6]) or cat_alpha
    sub_cat = s(r[7]) or cat_alpha
    qty = fi(r[10]); clase = s(r[8]) or 'General'
    if not clv or qty <= 0: continue
    reg_set4[reg] = 1
    if clv not in t_map:
        t_map[clv] = {'clave':clv,'tienda':tda,'region':reg,'gerente':'','cantidad':0,'cats':{}}
    t_map[clv]['cantidad'] += qty
    t_map[clv]['cats'][cat_alpha] = t_map[clv]['cats'].get(cat_alpha,0) + qty
    pk = prod+'|'+sub_cat+'|'+clase+'|'+cat_alpha
    if pk not in p_map:
        p_map[pk] = {'producto':prod,'sub_categoria':sub_cat,'clase':clase,'categoria':cat_alpha,'cantidad':0,'by_region':{}}
    p_map[pk]['cantidad'] += qty
    p_map[pk]['by_region'][reg] = p_map[pk]['by_region'].get(reg,0) + qty
    sk = sub_cat+'|'+clase+'|'+cat_alpha
    if sk not in r_map:
        r_map[sk] = {'sub_categoria':sub_cat,'clase':clase,'categoria':cat_alpha,'cantidad':0,'by_region':{}}
    r_map[sk]['cantidad'] += qty
    r_map[sk]['by_region'][reg] = r_map[sk]['by_region'].get(reg,0) + qty

oh_tienda   = sorted(t_map.values(), key=lambda x: x['clave'])
oh_productos = sorted(p_map.values(), key=lambda x: -x['cantidad'])
cat_totals  = {}
for p in oh_productos:
    cat_totals[p['categoria']] = cat_totals.get(p['categoria'],0) + p['cantidad']
oh_regiones = sorted(reg_set4.keys())

def mapcat(cat, qty, target):
    cl = cat.lower()
    if 'case' in cl:   target['oh_cases']    += qty
    elif 'cable' in cl: target['oh_cables']  += qty
    elif 'charg' in cl or 'cargad' in cl: target['oh_chargers'] += qty
    elif 'liquid' in cl: target['oh_liquid'] += qty
    elif 'mica' in cl or 'protec' in cl: target['oh_micas'] += qty

oh_map = {t['clave']: t for t in oh_tienda}
for c in cuota_jul:
    oh = oh_map.get(c['clave'])
    if not oh: continue
    c['oh_total'] = oh['cantidad']
    for cat, qty in oh['cats'].items(): mapcat(cat, qty, c)
print(f"OH: {len(oh_tienda)} tiendas, {len(oh_productos)} productos")

# ── 5. CUOTA LG-CASES ────────────────────────────────────────────────────────
sh3 = get_sheet('4.Cuota LG-Cases', 'Cuota LG-Cases Jul')
rows3 = list(sh3.iter_rows(values_only=True))

def alc_pct(v):
    f = flt(v); return round(f*100,1) if f < 2 else round(f,1)

lg_region = []
for r in rows3[8:15]:
    reg = s(r[0])
    if not reg or reg.upper() == 'TOTAL': continue
    lg_region.append({'region':reg,
        'lg_v':fi(r[3]),'lg_cuota':round(flt(r[5]),1),'lg_alc':alc_pct(r[6]),
        'cases_v':fi(r[7]),'cases_cuota':round(flt(r[9]),1),'cases_alc':alc_pct(r[10])})
tot = rows3[14] if len(rows3) > 14 else [0]*11
lg_totals = {'lg_v':fi(tot[3]),'lg_cuota':round(flt(tot[5]),1),
             'cases_v':fi(tot[7]),'cases_cuota':round(flt(tot[9]),1)}
lg_tienda = []
for r in rows3[18:]:
    reg = s(r[0]); clv = s(r[1])
    if not clv: continue
    lg_tienda.append({'region':reg,'clave':clv,'tienda':s(r[2]),
        'lg_v':fi(r[3]),'lg_cuota':round(flt(r[5]),1),'lg_alc':alc_pct(r[6]),
        'cases_v':fi(r[7]),'cases_cuota':round(flt(r[9]),1),'cases_alc':alc_pct(r[10]),
        'oh_lg':0,'oh_cases':0})
for r in lg_tienda:
    oh = oh_map.get(r['clave'])
    if not oh: continue
    for cat, qty in oh['cats'].items():
        cl = cat.lower()
        if 'liquid' in cl: r['oh_lg']    += qty
        if 'case'   in cl: r['oh_cases'] += qty
print(f"LG-Cases: {len(lg_region)} regiones, {len(lg_tienda)} tiendas")

# ── 6. BASE HISTÓRICO (Sales YoY) ────────────────────────────────────────────
bm = {}; bRegSet = {}; bCatSet = {}; count = 0
sh5 = get_sheet('6.Base', '2.Base')
for i, r in enumerate(sh5.iter_rows(values_only=True, min_row=4)):
    mes = s(r[8]); units = fi(r[7])
    if not mes or units <= 0: continue
    año = fi(r[9]) or 2025
    key = mes+'|'+str(año)
    reg = s(r[0]); cat = s(r[10])
    if key not in bm:
        bm[key] = {'mes':mes,'año':año,'total':0,'by_region':{},'by_cat':{},'by_rc':{}}
    bm[key]['total'] += units
    if reg:
        bm[key]['by_region'][reg] = bm[key]['by_region'].get(reg,0) + units
        bRegSet[reg] = 1
    if cat:
        bm[key]['by_cat'][cat] = bm[key]['by_cat'].get(cat,0) + units
        bCatSet[cat] = 1
    if reg and cat:
        rc = reg+'|'+cat
        bm[key]['by_rc'][rc] = bm[key]['by_rc'].get(rc,0) + units
    count += 1
    if count % 20000 == 0: print(f"  2.Base: {count} filas…")
print(f"2.Base: {count} filas, {len(bm)} meses")

# ── 6b. DETECTAR cuota_mes ANTES de agregar Sep ───────────────────────────────
sorted_keys_tmp = sorted(bm.keys(), key=lambda k:(int(k.split('|')[1]), MO.get(k.split('|')[0],0)))
complete_keys = [k for k in sorted_keys_tmp if bm[k]['total'] >= 5000]
if complete_keys:
    lmo, lyr = complete_keys[-1].split('|')
    idx = mo_list.index(lmo)
    cuota_mes = mo_list[idx+1] if idx < 11 else 'Jan'
    cuota_año = int(lyr) + (1 if idx == 11 else 0)
else:
    cuota_mes, cuota_año = 'Sep', 2026
print(f"Cuota mes detectado: {cuota_mes}|{cuota_año}")

# ── 6c. AGREGAR Sep|2026 desde PrimeMX "Detalle completo" ────────────────────
# Mapeo de regiones PrimeMX → nombre histórico (BAJIO+NORTE2 → BAJIO-NORTE 2)
REG_MAP = {
    'BAJIO':        'BAJIO-NORTE 2',
    'NORTE 2':      'BAJIO-NORTE 2',
    'CENTRO':       'CENTRO',
    'NOROESTE':     'NOROESTE',
    'NORTE 1':      'NORTE 1',
    'PACIFICO':     'PACIFICO',
    'SUR PENINSULA':'SUR PENINSULA',
}
sep_key = 'Sep|2026'
bm[sep_key] = {'mes':'Sep','año':2026,'total':0,'by_region':{},'by_cat':{},'by_rc':{}}

for r in wb2['Detalle completo'].iter_rows(values_only=True, min_row=2):
    reg_raw = s(r[3]); units = fi(r[8])
    if not reg_raw or units <= 0 or reg_raw == 'SIN ASIGNAR': continue
    reg_hist = REG_MAP.get(reg_raw, reg_raw)
    sku = s(r[1]); desc = s(r[2]) or sku
    bm[sep_key]['total'] += units
    bm[sep_key]['by_region'][reg_hist] = bm[sep_key]['by_region'].get(reg_hist,0) + units
    bm[sep_key]['by_cat'][desc] = bm[sep_key]['by_cat'].get(desc,0) + units
    rc = reg_hist+'|'+desc
    bm[sep_key]['by_rc'][rc] = bm[sep_key]['by_rc'].get(rc,0) + units
    bRegSet[reg_hist] = 1

print(f"Sep|2026 agregado: {bm[sep_key]['total']:,} uds  |  regiones: {dict(bm[sep_key]['by_region'])}")

# ── 6d. FINALIZAR base_data ───────────────────────────────────────────────────
sorted_keys = sorted(bm.keys(), key=lambda k:(int(k.split('|')[1]), MO.get(k.split('|')[0],0)))
base_data   = {'byMonth': bm, 'sorted': sorted_keys, 'cats': sorted(bCatSet.keys())}
base_regiones = sorted(bRegSet.keys())

# ── 7. TOTALES ────────────────────────────────────────────────────────────────
regiones   = sorted(set(r['region'] for r in cuota_jul if r['region']))
proy_total = sum(r['proy'] for r in cuota_jul)

# ── 8. EXPORTAR ───────────────────────────────────────────────────────────────
data = {
    'cuota_mes': cuota_mes, 'cuota_año': cuota_año,
    'primemx_summary': primemx_summary,
    'ar_region': ar_region,
    'cuota_jul': cuota_jul, 'ar_clean': ar_clean,
    'lg_region': lg_region, 'lg_tienda': lg_tienda,
    'lg_totals': lg_totals, 'cat_totals': cat_totals,
    'regiones':  regiones,  'proy_total': proy_total,
    'oh_tienda': oh_tienda, 'oh_productos': oh_productos,
    'oh_regiones': oh_regiones, 'rMap': r_map,
    'base_data': base_data, 'base_regiones': base_regiones
}

out = '/sessions/tender-amazing-curie/mnt/outputs/data.json'
with open(out,'w',encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, separators=(',',':'))
sz = os.path.getsize(out)
print(f"\n✅ data.json: {sz/1024:.0f} KB  ({sz/1024/1024:.2f} MB)")
