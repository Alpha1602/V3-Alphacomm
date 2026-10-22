"""Refresca en data.json (estructura actual) las secciones de ventas de Octubre con Ventas_Alphacomm_al_DD.MM.YYYY.xlsx.
Actualiza: cuota_oct, primemx_summary, ar_region, ar_clean. No toca inventario, ventas (semanal PrimeMX), lg/oh, base_data.
Uso: python3 update_oct_ventas.py VENTAS.xlsx data.json"""
import sys, json, calendar, collections, openpyxl
V, OUT = sys.argv[1:3]
d = json.load(open(OUT)); co = d['cuota_oct']
wb = openpyxl.load_workbook(V, read_only=True, data_only=True)
rows = [r for r in list(wb['BD OCTUBRE'].iter_rows(values_only=True))[1:] if r[0]]
cut = max(r[3] for r in rows); dim = calendar.monthrange(cut.year, cut.month)[1]; avt = round(cut.day/dim, 4)
vt = collections.Counter(); vr = collections.Counter(); dia = collections.Counter()
for r in rows:
    q = r[10] or 0; vt[r[4]] += q; vr[r[12]] += q; dia[r[3].strftime('%Y-%m-%d')] += q
tot = sum(vt.values())
# por_region
c100 = sum(x['cuota_100k'] for x in co['por_region']); cp = sum(x['cuota_prime'] for x in co['por_region'])
for x in co['por_region']:
    v = vr.get(x['region'], 0)
    x.update(ventas=v, faltante=max(x['cuota_100k']-v, 0), proyeccion=round(v/avt),
             alcance=round(v/x['cuota_100k'], 3), alc_proy=round(v/avt/x['cuota_100k'], 3))
# por_tienda
pt = {x['clave']: x for x in co['por_tienda']}
for k, x in pt.items():
    x['ventas'] = vt.get(k, 0); x['alcance'] = round(x['ventas']/x['cuota'], 3) if x['cuota'] else 0
proy = round(tot/avt)
co.update(fecha_corte=cut.strftime('%Y-%m-%d'), avance_tiempo=avt, total_ventas=tot, proyeccion=proy,
          alcance=round(tot/c100, 3), alc_proy=round(proy/c100, 3), faltante=max(c100-tot, 0),
          diario=dict(sorted(dia.items())))
# base_data (histórico) del mes en curso
bm = d['base_data']['byMonth'].get('Oct|2026')
if bm is not None:
    HM = {'BAJIO':'BAJIO-NORTE 2','NORTE 2':'BAJIO-NORTE 2'}
    br = collections.Counter(); bc = collections.Counter(); brc = collections.Counter()
    for r in rows:
        q = r[10] or 0; rg = HM.get(r[12], r[12]); ds = (r[1] or r[0]).strip()
        br[rg] += q; bc[ds] += q; brc[rg+'|'+ds] += q
    bm.update(total=tot, by_region=dict(br), by_cat=dict(bc), by_rc=dict(brc))
# share por categoría y región (7 regiones, mes en curso)
def cat_of(desc):
    t = (desc or '').upper()
    if 'LIQUID' in t: return 'Liquid'
    if 'MICA' in t or 'VIDRIO' in t or 'TEMPERED' in t: return 'Micas'
    if 'CABLE' in t: return 'Cables'
    if 'CARG' in t or 'CHARGER' in t: return 'Charger'
    if any(w in t for w in ('FUNDA', 'CASE', 'SPECK', 'SPK', 'BDLSP')): return 'Cases iPhone'
    return 'Accesorios'
CATS = ['Cables', 'Charger', 'Cases iPhone', 'Liquid', 'Micas', 'Accesorios']
sr = collections.defaultdict(collections.Counter); st = collections.Counter()
for r in rows:
    q = r[10] or 0; c = cat_of(r[1] or r[0]); sr[r[12]][c] += q; st[c] += q
d['share_cat'] = {'periodo': f"oct-26 · corte día {cut.day}", 'fecha_corte': cut.strftime('%Y-%m-%d'), 'categorias': CATS,
    'total': {c: st.get(c, 0) for c in CATS},
    'regiones': [{'region': rg, 'total': sum(sr[rg].values()), 'cats': {c: sr[rg].get(c, 0) for c in CATS}} for rg in sorted(sr)]}
# AR
pdv = [r for r in list(wb['Equipos por PDV'].iter_rows(values_only=True))[4:] if r[2]]
oar = {x['clave']: x for x in d['ar_clean']}; seen = set()
for r in pdv:
    a = int(r[5] or 0); e = int(r[4] or 0); seen.add(r[2])
    x = oar.get(r[2])
    if x is None:
        x = {'region': r[0], 'gerente': r[1] or '', 'clave': r[2], 'tienda': r[3], 'pct_alpha': 0, 'proy': 0}; d['ar_clean'].append(x)
    x.update(gerente=r[1] or x.get('gerente', ''), alphacomm=a, total_acc=a, equipo=e, ar_ratio=round(a/e, 3) if e else 0)
for x in d['ar_clean']:
    if x['clave'] not in seen: x.update(alphacomm=0, total_acc=0, equipo=0, ar_ratio=0)
ra = collections.defaultdict(lambda: [0, 0])
for r in pdv: ra[r[0]][0] += int(r[5] or 0); ra[r[0]][1] += int(r[4] or 0)
for x in d['ar_region']:
    a, e = ra[x['region']]; x.update(alphacomm=a, total_acc=a, equipo=e, ar_ratio=round(a/e, 3) if e else 0)
A = sum(v[0] for v in ra.values()); E = sum(v[1] for v in ra.values()); con = sum(1 for r in pdv if (r[5] or 0) > 0)
d['primemx_summary'].update(periodo=f"oct-26 · corte día {cut.day}", alpha_units=A, total_acc=A, equipo=E,
                            ar_ratio=round(A/E, 3), pdvs_con=con, pdvs_sin=len(pdv)-con)
json.dump(d, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
print(f"Corte {cut.date()} avt={avt} | ventas {tot} | proy {proy} | alc {tot/c100:.1%} | AR {A}/{E}={A/E:.3f} | pdvs {con}/{len(pdv)}")
