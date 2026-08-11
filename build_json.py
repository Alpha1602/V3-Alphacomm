import openpyxl, json, os, sys

EXCEL = sys.argv[1] if len(sys.argv) > 1 else "/sessions/tender-amazing-curie/mnt/uploads/Sales YoY .3.xlsx"
wb = openpyxl.load_workbook(EXCEL, read_only=True, data_only=True)

# Helper: pick first existing sheet name
def get_sheet(*names):
    for n in names:
        if n in wb.sheetnames: return wb[n]
    raise KeyError(f"None found: {names}")

def s(v): return '' if v is None else str(v).strip()
def flt(v):
    try: return float(v)
    except: return 0.0
def fi(v): return int(round(flt(v)))

# ── 1. CUOTA JUL ─────────────────────────────────────────────────────────
# header row 2 (idx 2), data from row 3 (idx 3)
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
print(f"Cuota Jul: {len(cuota_jul)} tiendas")

# ── 2. PROYECCION TIENDA → proy por clave ────────────────────────────────
# header row 2, data from row 3: 0=region,3=clave,4=tienda,5=alpha_actual,8=alpha_proy
pm = {}
try:
    sh_pt = get_sheet('2.Proyección Tienda', 'Proyección Tienda')
    for r in sh_pt.iter_rows(values_only=True, min_row=4):
        clv = s(r[3])
        if clv: pm[clv] = round(flt(r[8]),1)   # Alphacomm proy. cierre
except: pass
if not pm:  # fallback: AR sheet col 10
    for r in get_sheet('3.AR','1.AR').iter_rows(values_only=True, min_row=4):
        clv = s(r[3])
        if clv: pm[clv] = round(flt(r[10]),1)
for c in cuota_jul:
    c['proy'] = pm.get(c['clave'], c['ventas'])
print(f"Proy: {len(pm)} tiendas mapeadas")

# ── 3. 3.AR → ar_clean ───────────────────────────────────────────────────
# header row 2: 0=region,2=gerente,3=clave,4=tienda,5=ventas_alpha,8=equipo,10=alpha_proy,11=equipo_proy
ar_clean = []
for r in get_sheet('3.AR','1.AR').iter_rows(values_only=True, min_row=4):
    reg = s(r[0]); clv = s(r[3])
    if not reg or not clv: continue
    ar_clean.append({'region':reg,'gerente':s(r[2]),'clave':clv,'tienda':s(r[4]),
        'alphacomm': fi(r[5]),   # Ventas Alphacomm del período actual
        'equipo':    fi(r[8]),   # Venta de equipo del período actual
        'proy':      round(flt(r[10]),1)})  # Alphacomm proy. cierre
print(f"1.AR: {len(ar_clean)} registros")

# ── 4. OH DETALLE → oh_tienda, oh_productos, _rMap ───────────────────────
# header row 2: 0=tienda,1=clave,2=regional,3=región,4=codigo,5=sku_alpha(catAlpha),
#               6=descripcion(prod),7=modelo(sub_cat),8=subinv(clase),10=cantidad
sh4 = get_sheet('5.OH detalle', 'OH detalle')
t_map = {}   # clave → {clave,tienda,region,gerente,cantidad,cats:{}}
p_map = {}   # pk   → {producto,sub_categoria,clase,categoria,cantidad,by_region:{}}
r_map = {}   # sk   → {sub_categoria,clase,categoria,cantidad,by_region:{}}
reg_set4 = {}

for r in sh4.iter_rows(values_only=True, min_row=4):
    reg = s(r[3]); clv = s(r[1]); tda = s(r[0])
    cat_alpha = s(r[5]) or 'Accesories'
    prod = s(r[6]) or cat_alpha
    sub_cat = s(r[7]) or cat_alpha
    qty = fi(r[10])
    clase = s(r[8]) or 'General'
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

oh_tienda = sorted(t_map.values(), key=lambda x: x['clave'])
oh_productos = sorted(p_map.values(), key=lambda x: -x['cantidad'])
cat_totals = {}
for p in oh_productos:
    cat_totals[p['categoria']] = cat_totals.get(p['categoria'],0) + p['cantidad']
oh_regiones = sorted(reg_set4.keys())

# Patch OH into cuota_jul
def mapcat(cat, qty, target):
    cl = cat.lower()
    if 'case' in cl: target['oh_cases'] += qty
    elif 'cable' in cl: target['oh_cables'] += qty
    elif 'charg' in cl or 'cargad' in cl: target['oh_chargers'] += qty
    elif 'liquid' in cl: target['oh_liquid'] += qty
    elif 'mica' in cl or 'protec' in cl: target['oh_micas'] += qty

oh_map = {t['clave']: t for t in oh_tienda}
for c in cuota_jul:
    oh = oh_map.get(c['clave'])
    if not oh: continue
    c['oh_total'] = oh['cantidad']
    for cat, qty in oh['cats'].items(): mapcat(cat, qty, c)

print(f"OH: {len(oh_tienda)} tiendas, {len(oh_productos)} productos, {len(oh_regiones)} regiones")

# ── 5. CUOTA LG-CASES JUL ────────────────────────────────────────────────
sh3 = get_sheet('4.Cuota LG-Cases', 'Cuota LG-Cases Jul')
rows3 = list(sh3.iter_rows(values_only=True))

def alc_pct(v):
    f = flt(v); return round(f*100,1) if f < 2 else round(f,1)

# Region summary: rows 8-13 (idx), header at idx 7
lg_region = []
for r in rows3[8:15]:
    reg = s(r[0])
    if not reg or reg.upper() == 'TOTAL': continue
    lg_region.append({'region':reg,
        'lg_v':fi(r[3]),'lg_cuota':round(flt(r[5]),1),'lg_alc':alc_pct(r[6]),
        'cases_v':fi(r[7]),'cases_cuota':round(flt(r[9]),1),'cases_alc':alc_pct(r[10])})

# Totals row 14
tot = rows3[14] if len(rows3) > 14 else [0]*11
lg_totals = {'lg_v':fi(tot[3]),'lg_cuota':round(flt(tot[5]),1),
             'cases_v':fi(tot[7]),'cases_cuota':round(flt(tot[9]),1)}

# Tienda detail: header at idx 17, data from 18
lg_tienda = []
for r in rows3[18:]:
    reg = s(r[0]); clv = s(r[1])
    if not clv: continue
    lg_tienda.append({'region':reg,'clave':clv,'tienda':s(r[2]),
        'lg_v':fi(r[3]),'lg_cuota':round(flt(r[5]),1),'lg_alc':alc_pct(r[6]),
        'cases_v':fi(r[7]),'cases_cuota':round(flt(r[9]),1),'cases_alc':alc_pct(r[10]),
        'oh_lg':0,'oh_cases':0})

# Patch OH into lg_tienda
for r in lg_tienda:
    oh = oh_map.get(r['clave'])
    if not oh: continue
    for cat, qty in oh['cats'].items():
        cl = cat.lower()
        if 'liquid' in cl: r['oh_lg'] += qty
        if 'case' in cl: r['oh_cases'] += qty

print(f"LG-Cases: {len(lg_region)} regiones, {len(lg_tienda)} tiendas")

# ── 6. 2.BASE (historical) ────────────────────────────────────────────────
MO = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
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

sorted_keys = sorted(bm.keys(), key=lambda k:(int(k.split('|')[1]), MO.get(k.split('|')[0],0)))
base_data = {'byMonth': bm, 'sorted': sorted_keys, 'cats': sorted(bCatSet.keys())}
base_regiones = sorted(bRegSet.keys())
print(f"2.Base: {count} filas, {len(bm)} meses")

# ── 7. TOTALES + MES CUOTA ACTUAL ─────────────────────────────────────────
regiones = sorted(set(r['region'] for r in cuota_jul if r['region']))
proy_total = sum(r['proy'] for r in cuota_jul)

# Detectar mes de cuota: el mes siguiente al último mes completo (>= 5000 uds)
mo_list = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
complete_keys = [k for k in sorted_keys if bm[k]['total'] >= 5000]
if complete_keys:
    lmo, lyr = complete_keys[-1].split('|')
    idx = mo_list.index(lmo)
    cuota_mes = mo_list[idx+1] if idx < 11 else 'Jan'
    cuota_año = int(lyr) + (1 if idx == 11 else 0)
else:
    cuota_mes, cuota_año = 'Aug', 2026
print(f"Cuota mes detectado: {cuota_mes}|{cuota_año}")

# ── 8. EXPORTAR ───────────────────────────────────────────────────────────
data = {
    'cuota_mes': cuota_mes, 'cuota_año': cuota_año,
    'cuota_jul': cuota_jul, 'ar_clean': ar_clean,
    'lg_region': lg_region, 'lg_tienda': lg_tienda,
    'lg_totals': lg_totals, 'cat_totals': cat_totals,
    'regiones': regiones, 'proy_total': proy_total,
    'oh_tienda': oh_tienda, 'oh_productos': oh_productos,
    'oh_regiones': oh_regiones, 'rMap': r_map,
    'base_data': base_data, 'base_regiones': base_regiones
}

out = sys.argv[2] if len(sys.argv) > 2 else '/sessions/brave-zealous-einstein/mnt/2.Prime/data.json'
with open(out,'w',encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, separators=(',',':'))
sz = os.path.getsize(out)
print(f"\n✅ data.json: {sz/1024:.0f} KB  ({sz/1024/1024:.2f} MB)")
