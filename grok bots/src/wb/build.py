"""Build GPU price history workbook + CSVs + charts from cached Wayback snapshots.
Usage: python build.py   (re-parses everything under /workspace/nebius/wayback_raw)"""
import sys, os, json, re, datetime as dt
sys.path.insert(0, os.path.dirname(__file__))
import pandas as pd, numpy as np
from run_parse import parse_all
from parse import PARSERS, money, oracle_api
from txt import text, raw

ROOT = '/workspace/nebius'
OUT_MAIN = f'{ROOT}/market/gpu_price_history.xlsx'
PARTS = f'{ROOT}/parts'; CHARTS = f'{ROOT}/market/charts'
os.makedirs(PARTS, exist_ok=True); os.makedirs(CHARTS, exist_ok=True)
PTYPE = {'Nebius': 'nebius', 'AWS': 'hyperscaler', 'Azure': 'hyperscaler', 'Google Cloud': 'hyperscaler', 'Oracle OCI': 'hyperscaler'}
SKU = {'H100': 'H100 SXM 80GB', 'H100 PCIe': 'H100 PCIe 80GB', 'H100 NVL': 'H100 NVL 94GB', 'H200': 'H200 141GB', 'B200': 'B200', 'B300': 'B300',
       'GB200': 'GB200 NVL72', 'GB300': 'GB300 NVL72', 'A100 80GB': 'A100 80GB', 'A100 40GB': 'A100 40GB', 'L40S': 'L40S', 'GH200': 'GH200',
       'RTX PRO 6000': 'RTX PRO 6000', 'MI300X': 'MI300X', 'MI355X': 'MI355X', 'L40': 'L40', 'A100': 'A100'}
H100_SPECS = {
    'Nebius': 'per-GPU VM slice of HGX H100: 16 vCPU, 200 GB RAM (per page); InfiniBand clusters',
    'CoreWeave': '8-GPU HGX H100 node: 128 vCPU, 2,048 GB RAM, 61.44 TB local NVMe, InfiniBand (per page)',
    'Lambda': '8x H100 SXM instance: 208 vCPU, 1,800 GiB RAM, 22-26 TiB SSD (per page)',
    'AWS': 'p5.48xlarge: 8x H100, 192 vCPU, 2 TiB RAM, 3,200 Gbps EFAv2, 8x3.84 TB NVMe',
    'Azure': 'ND96isr H100 v5: 8x H100, 96 vCPU, 1,900 GiB RAM, 8x400 Gb/s InfiniBand',
    'Oracle OCI': 'BM.GPU.H100.8 bare metal: 8x H100, 112 cores, 2 TB RAM, 16x3.84 TB NVMe, 8x2x200 Gb/s RDMA',
    'Crusoe': 'HGX H100 80GB SXM VM (per-GPU price)', 'Together AI': 'HGX H100 SXM GPU cluster (per-GPU price)',
}

def region_of(r):
    c = r['config'] or ''
    m = re.search(r'(US East \(N\. Virginia\)|US East \(Ohio\)|US West \(Oregon\))', c)
    if m: return m.group(1)
    if r['provider'] == 'CoreWeave': return 'US (page default region)'
    if r['provider'] == 'Azure': return r.get('region') or 'lowest region (tracker)'
    if r['provider'] == 'AWS': return 'lowest region (tracker; = us-east-1)'
    return 'not stated (single global list)'

def basis_of(r):
    q = (r['quoted_price'] or '') + ' ' + (r['term'] or '') + ' ' + (r['notes'] or '')
    if 'ANOMALY' in q: return 'anomalous (excluded)'
    if 'Promo' in q or 'promo' in q: return 'promotional'
    if re.search(r'Starting at|starting at|as low as|from \$|"from"', q): return 'list floor ("starting at"/"as low as"/"from")'
    if 'DATED FROM ANNOUNCEMENT' in q: return 'list (effective date from provider announcement)'
    if 'tracker' in q: return 'list (via third-party tracker)'
    if r['pricing_type'] == 'capacity-block': return 'list (EC2 Capacity Block published rate)'
    return 'list'

def live_records():
    """Live pages captured 2026-10-04 by the market workstream (/workspace/nebius/raw). Not Wayback."""
    out = []; ts = '20261004174800'
    live = [('Nebius', 'raw/nebius_com_prices.html', 'https://nebius.com/prices'),
            ('Crusoe', 'raw/www_crusoe_ai_cloud_pricing.html', 'https://www.crusoe.ai/cloud/pricing'),
            ('Together AI', 'raw/www_together_ai_pricing.html', 'https://www.together.ai/pricing'),
            ('RunPod', 'raw/www_runpod_io_pricing.html', 'https://www.runpod.io/pricing'),
            ('AWS', 'raw/aws_amazon_com_ec2_capacityblocks_pricing_.html', 'https://aws.amazon.com/ec2/capacityblocks/pricing/')]
    for prov, f, url in live:
        fp = f'{ROOT}/{f}'
        if not os.path.exists(fp): continue
        for r in PARSERS[prov](text(fp)):
            r.update(provider=prov, snapshot_timestamp=ts, snapshot_url='LIVE capture ' + fp, original_url=url, source='live page 2026-10-04 17:48 ET')
            out.append(r)
    fp = f'{ROOT}/raw/oci_products.json'
    if os.path.exists(fp):
        for r in oracle_api(open(fp).read()):
            r.update(provider='Oracle OCI', snapshot_timestamp=ts, snapshot_url='LIVE capture ' + fp, original_url='https://apexapps.oracle.com/pls/apex/cetools/api/v1/products/', source='live API 2026-10-04')
            out.append(r)
    # AWS P5 on-demand cut: effective date from AWS announcement, level from instances.vantage.sh 2025-09-30 snapshot
    out.append(dict(chip='H100', config='p5.48xlarge (8x H100)', pricing_type='on-demand', term='PAYG', quoted_price='$55.04', quoted_unit='per 8-GPU instance-hour',
                    gpus_per_unit=8, headline=True, provider='AWS', snapshot_timestamp='20250601000000',
                    snapshot_url='https://aws.amazon.com/about-aws/whats-new/2025/06/pricing-usage-model-ec2-instances-nvidia-gpus/',
                    original_url='https://aws.amazon.com/blogs/aws/announcing-up-to-45-price-reduction-for-amazon-ec2-nvidia-gpu-accelerated-instances/',
                    notes='DATED FROM ANNOUNCEMENT: AWS cut P5 On-Demand ~44% effective 2025-06-01 (posted 2025-06-05); $55.04 level = first post-cut tracker snapshot 2025-09-30 (98.32 x 0.56 = 55.06)',
                    source='AWS What\'s New / News Blog (press)', region='us-east-1'))
    # Azure Retail Prices API (current) -- effectiveStartDate gives the date the price took effect
    for fn, chip in [('az_ND96isr_H100_v5.json', 'H100'), ('az_ND96isr_H200_v5.json', 'H200'), ('az_ND128isr_NDR_GB200_v6.json', 'GB200')]:
        fp = f'{ROOT}/raw/{fn}'
        if not os.path.exists(fp): continue
        items = json.load(open(fp)).get('Items', [])
        for it in items:
            if it.get('armRegionName') != 'eastus' or 'Windows' in it.get('productName', '') or 'Low Priority' in it.get('skuName', ''): continue
            n = 4 if chip == 'GB200' else 8
            sku = it.get('skuName'); typ = it.get('type'); term = it.get('reservationTerm'); p = it.get('retailPrice')
            if typ == 'Consumption':
                pt, tm, unit, usd = ('spot', 'spot', f'per {n}-GPU VM-hour', p / n) if 'Spot' in sku else ('on-demand', 'PAYG', f'per {n}-GPU VM-hour', p / n)
                q = f'${p}'
            elif typ == 'Reservation':
                yrs = int(term.split()[0]); hrs = 8760 * yrs
                pt, tm, unit, usd = 'reserved', f'{yrs}-year', f'total upfront for {yrs}-yr {n}-GPU VM reservation', p / hrs / n
                q = f'${p:,.0f}'
            else: continue
            eff = it.get('effectiveStartDate', '')[:10]
            for when, label in [(eff.replace('-', '') + '000000', 'effective start'), (ts, 'live')]:
                r = dict(chip=chip, config=f"Azure {it.get('armSkuName')} ({n}x {chip})", pricing_type=pt, term=tm, quoted_price=q, quoted_unit=unit,
                         gpus_per_unit=n, headline=(pt != 'spot' or True), notes=f'Azure Retail Prices API, eastus Linux; effectiveStartDate {eff}' +
                         (f'; derived per GPU-hr = {q}/({8760*int(term.split()[0]) if term else 1} h x {n} GPUs)' if typ == 'Reservation' else f'; derived = {q}/{n} GPUs') +
                         ('' if label == 'live' else '; row dated at effectiveStartDate (price in force since then per API)'),
                         provider='Azure', snapshot_timestamp=when, snapshot_url='LIVE capture ' + fp, original_url='https://prices.azure.com/api/retail/prices', source=f'Azure Retail Prices API ({label})', usd_override=round(usd, 4), region='eastus')
                out.append(r)
    return out

def to_long(recs):
    rows = []
    for r in recs:
        q = money(r.get('quoted_price'))
        usd = r.get('usd_override')
        if usd is None and q is not None:
            unit = r['quoted_unit'] or ''
            if 'minute' in unit: usd = q * 60
            elif ('instance' in unit or 'VM-hour' in unit) and r['gpus_per_unit']: usd = q / r['gpus_per_unit']
            else: usd = q
        ts = r['snapshot_timestamp']
        derived = ''
        if usd is not None and q is not None and abs(usd - q) > 1e-9 and 'override' not in r:
            derived = f"derived: {r['quoted_price']} {r['quoted_unit']} -> /{r['gpus_per_unit']} GPUs" if 'minute' not in (r['quoted_unit'] or '') else f"derived: {r['quoted_price']}/min x 60"
        prov = r['provider']
        sku = SKU.get(r['chip'], r['chip'])
        specs = H100_SPECS.get(prov, '') if sku == 'H100 SXM 80GB' else ''
        nm = r['notes'] or ''
        if 'vCPUs per GPU' in nm: specs = nm.split('; ', 1)[-1]
        rows.append(dict(
            date=pd.Timestamp(dt.datetime.strptime(ts[:8], '%Y%m%d')), provider=prov, provider_type=PTYPE.get(prov, 'neocloud'), SKU=sku,
            pricing_type=r['pricing_type'], term=r['term'], price_basis=basis_of(r), usd_per_gpu_hr=None if usd is None else round(usd, 4),
            quoted_price=r['quoted_price'], quoted_unit=r['quoted_unit'], gpus_per_instance=r['gpus_per_unit'], bundled_specs=specs,
            region=r.get('region') or region_of(r), currency='USD', source_url=r['snapshot_url'], wayback_timestamp=ts if not r['snapshot_url'].startswith('LIVE') else '',
            original_url=r['original_url'], config=r['config'], headline=bool(r['headline']) and usd is not None and 'anomalous' not in basis_of(r),
            notes='; '.join(x for x in [nm, derived, r.get('source', '')] if x)))
    df = pd.DataFrame(rows)
    # Nebius: once the 'Effective October 1, 2026' column is in force, it is the on-demand price; the old column becomes 'prior price still displayed'
    eff = (df.term == 'PAYG (announced, effective 2026-10-01)') & (df.date >= '2026-10-01')
    for _, r in df[eff].iterrows():
        same = (df.provider == r.provider) & (df.SKU == r.SKU) & (df.date == r.date) & (df.term == 'PAYG') & (df.source_url == r.source_url)
        df.loc[same, 'term'] = 'PAYG (pre-Oct-1 price, still displayed)'; df.loc[same, 'headline'] = False
    df.loc[eff, 'notes'] = df.loc[eff, 'notes'] + '; column "GPU-hour (Effective October 1, 2026)" -> treated as current on-demand price on/after 2026-10-01'
    df.loc[eff, 'term'] = 'PAYG'
    df = df.drop_duplicates(subset=['date', 'provider', 'SKU', 'pricing_type', 'term', 'config', 'quoted_price', 'wayback_timestamp'])
    return df.sort_values(['SKU', 'provider', 'pricing_type', 'term', 'date']).reset_index(drop=True)

def series_key(r):
    return f"{r.provider} | {r.SKU} | {r.pricing_type} | {r.term}"

MONTHS = pd.period_range('2024-01', '2026-10', freq='M')
def monthly(df):
    h = df[df.headline].copy()
    h['month'] = h.date.dt.to_period('M')
    h = h[(h.month >= MONTHS[0]) & (h.month <= MONTHS[-1])]
    h['series'] = h.apply(series_key, axis=1)
    h = h.sort_values(['series', 'date', 'wayback_timestamp'])
    # several configs can map to one series key (e.g. two H200 SKUs); take min config-sorted first to keep one per series
    h = h.groupby(['series', 'month']).last().reset_index()
    piv = h.pivot(index='month', columns='series', values='usd_per_gpu_hr').reindex(MONTHS)
    piv.index = piv.index.astype(str)
    return piv, h

def seeded_ffill(df, piv):
    """Forward-fill monthly pivot; series observed before Jan 2024 (e.g. Azure effectiveStartDate 2023-12-01, Nebius Nov/Dec 2023) are seeded with that last pre-2024 value."""
    h = df[df.headline].copy(); h['series'] = h.apply(series_key, axis=1)
    pre = h[h.date < '2024-01-01'].sort_values('date').groupby('series').last()
    seed = piv.copy()
    for s_, r in pre.iterrows():
        if s_ in seed.columns and pd.isna(seed.iloc[0][s_]): seed.loc[seed.index[0], s_] = r.usd_per_gpu_hr
        elif s_ not in seed.columns: seed[s_] = np.nan; seed.loc[seed.index[0], s_] = r.usd_per_gpu_hr
    return seed.ffill()

def quarterly(df, piv_obs):
    qs = pd.period_range('2024Q1', '2026Q2', freq='Q')
    h = df[df.headline].copy(); h['series'] = h.apply(series_key, axis=1)
    h['q'] = h.date.dt.to_period('Q')
    rows = []
    for s, g in h.groupby('series'):
        g = g.sort_values('date')
        row = {'series': s, 'provider': g.provider.iloc[0], 'SKU': g.SKU.iloc[0], 'pricing_type': g.pricing_type.iloc[0], 'term': g.term.iloc[0]}
        for q in qs:
            inq = g[g.q == q]
            label = str(q)
            if len(inq):
                last = inq.iloc[-1]
                tag = 'Derived' if str(last.notes).find('derived') >= 0 else 'Reported'
                row[label] = last.usd_per_gpu_hr
                row[label + ' tag'] = tag + (f" ({last.notes[last.notes.find('derived'):].split(';')[0]})" if tag == 'Derived' else f" (snapshot {last.date.date()})")
            else:
                prior = g[g.q < q]; later = g[g.q > q]
                if len(prior):
                    lp = prior.iloc[-1]
                    row[label] = lp.usd_per_gpu_hr
                    unchanged = len(later) and later.iloc[0].usd_per_gpu_hr == lp.usd_per_gpu_hr
                    row[label + ' tag'] = f"Estimated (no snapshot in quarter; carried forward from {lp.date.date()}" + ("; same price observed again on " + str(later.iloc[0].date.date()) + ")" if unchanged else ")")
                else:
                    row[label] = None
                    row[label + ' tag'] = 'no snapshot / not disclosed (no earlier observation; SKU not yet listed or page not archived/parseable)'
        rows.append(row)
    return pd.DataFrame(rows)

LAUNCHES = [  # successor-chip cloud availability dates used for post-launch decline windows
    ('H200', '2024-09-01', 'AWS P5e (H200) GA Sep 2024 (aws.amazon.com/about-aws/whats-new/2024/09/amazon-ec2-p5e-instances-generally-available/); Nebius lists H200 from Nov 2024 snapshot'),
    ('B200', '2025-02-01', 'First public B200 cloud listings in our data: CoreWeave/Together pages Q1 2025; NVIDIA Blackwell announced 2024-03-18 (nvidianews.nvidia.com)'),
    ('GB200', '2025-02-04', 'CoreWeave first GA of GB200 NVL72 instances, press release 2025-02-04 (coreweave.com/news)'),
]
def decline_rates(df):
    h = df[df.headline & df.SKU.isin(['H100 SXM 80GB', 'H200 141GB', 'B200', 'GB200 NVL72', 'A100 80GB', 'H100 PCIe 80GB'])].copy()
    h['series'] = h.apply(series_key, axis=1)
    rows = []
    for s, g in h.groupby('series'):
        g = g.sort_values('date'); f, l = g.iloc[0], g.iloc[-1]
        days = (l.date - f.date).days
        ann = (l.usd_per_gpu_hr / f.usd_per_gpu_hr) ** (365 / days) - 1 if days >= 90 and f.usd_per_gpu_hr else None
        r = dict(series=s, provider=f.provider, SKU=f.SKU, pricing_type=f.pricing_type, term=f.term, first_date=f.date.date(), first_usd=f.usd_per_gpu_hr,
                 last_date=l.date.date(), last_usd=l.usd_per_gpu_hr, total_change_pct=round(100 * (l.usd_per_gpu_hr / f.usd_per_gpu_hr - 1), 1),
                 annualized_change_pct=None if ann is None else round(100 * ann, 1), n_obs=len(g))
        # peak-to-date and to Jun-2026 (pre-2026 increases)
        pre = g[g.date < '2026-05-01']
        if len(pre) and (pre.date.iloc[-1] - f.date).days >= 90:
            r['annualized_to_Apr2026_pct'] = round(100 * ((pre.iloc[-1].usd_per_gpu_hr / f.usd_per_gpu_hr) ** (365 / (pre.date.iloc[-1] - f.date).days) - 1), 1)
        for chip, d0, src in LAUNCHES:
            d0 = pd.Timestamp(d0)
            before = g[g.date <= d0]; after = g[(g.date > d0 + pd.DateOffset(months=6)) & (g.date <= d0 + pd.DateOffset(months=9))]
            if len(before) and len(after):
                r[f'post_{chip}_launch_6-9mo_change_pct'] = round(100 * (after.iloc[-1].usd_per_gpu_hr / before.iloc[-1].usd_per_gpu_hr - 1), 1)
                r[f'post_{chip}_basis'] = f"{before.iloc[-1].date.date()} ${before.iloc[-1].usd_per_gpu_hr} -> {after.iloc[-1].date.date()} ${after.iloc[-1].usd_per_gpu_hr}"
        rows.append(r)
    out = pd.DataFrame(rows)
    launches = pd.DataFrame(LAUNCHES, columns=['successor_chip', 'launch_date_used', 'source'])
    return out, launches

def recent_changes(df):
    """Compare each headline series: last observation before 2026-06-01, last before 2026-08-01, latest (incl. live)."""
    h = df[df.headline].copy(); h['series'] = h.apply(series_key, axis=1)
    rows = []
    for s, g in h.groupby('series'):
        g = g.sort_values('date')
        if g.date.max() < pd.Timestamp('2026-05-01'): continue
        a = g[g.date < '2026-06-01']; b = g[g.date < '2026-08-01']; c = g.iloc[-1]
        av = a.iloc[-1] if len(a) else None; bv = b.iloc[-1] if len(b) else None
        # dated change points after 2026-01-01
        chg = []
        g26 = g[g.date >= '2026-01-01']
        prev = None
        for _, x in g26.iterrows():
            if prev is not None and x.usd_per_gpu_hr != prev.usd_per_gpu_hr:
                chg.append(f"${prev.usd_per_gpu_hr:g} ({prev.date.date()}) -> ${x.usd_per_gpu_hr:g} ({x.date.date()}) {100*(x.usd_per_gpu_hr/prev.usd_per_gpu_hr-1):+.0f}%")
            prev = x
        base = av if av is not None else (g.iloc[0])
        rows.append(dict(provider=c.provider, SKU=c.SKU, pricing_type=c.pricing_type, term=c.term,
                         last_pre_Jun2026_date=None if av is None else av.date.date(), last_pre_Jun2026_usd=None if av is None else av.usd_per_gpu_hr,
                         last_pre_Aug2026_date=None if bv is None else bv.date.date(), last_pre_Aug2026_usd=None if bv is None else bv.usd_per_gpu_hr,
                         latest_date=c.date.date(), latest_usd=c.usd_per_gpu_hr, latest_source=c.source_url,
                         change_vs_pre_Jun_pct=None if av is None else round(100 * (c.usd_per_gpu_hr / av.usd_per_gpu_hr - 1), 1),
                         change_vs_pre_Aug_pct=None if bv is None else round(100 * (c.usd_per_gpu_hr / bv.usd_per_gpu_hr - 1), 1),
                         dated_changes_2026='; '.join(chg) or 'no change observed in 2026'))
    return pd.DataFrame(rows).sort_values(['SKU', 'provider', 'pricing_type', 'term'])

def nebius_wording():
    rows = []; prev = None
    logs = [json.loads(l) for l in open(f'{ROOT}/wayback_raw/fetchlog_nebius.jsonl') if json.loads(l).get('file')]
    for d in sorted(logs, key=lambda d: d['ts']):
        t = text(d['file'])
        w = re.findall(r'((?:Save|Pay|Cut|get)[^.]{0,40}?(?:up to|more than|as low as) [^.]{0,120}?\.|Commitment discounts [^.]{0,200}\.|Both amount of GPUs[^.]{0,200}\.)', t)
        w = ' | '.join(dict.fromkeys(w))
        rows.append(dict(snapshot_date=f"{d['ts'][:4]}-{d['ts'][4:6]}-{d['ts'][6:8]}", wording=w, changed='YES' if w != prev else '', snapshot_url=f"https://web.archive.org/web/{d['ts']}/{d['orig']}"))
        prev = w
    return pd.DataFrame(rows)

def charts(df, piv, ffs):
    import matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt
    plt.rcParams.update({'font.family': 'DejaVu Sans', 'font.size': 10, 'axes.spines.top': False, 'axes.spines.right': False})
    colors = {'Nebius': '#0B5FFF', 'CoreWeave': '#E4572E', 'Lambda': '#6A4C93', 'Crusoe': '#8B5E34', 'Together AI': '#29BF12', 'RunPod': '#8D99AE',
              'AWS': '#FF9900', 'Azure': '#008AD7', 'Oracle OCI': '#C74634', 'Google Cloud': '#34A853'}
    made = []
    def foot(fig, extra=''):
        fig.text(0.01, 0.01, 'Source: Internet Archive Wayback Machine snapshots of provider pricing pages (Jan 2024-Oct 2026); Azure Retail Prices API; OCI price-list API; '
                 'instances.vantage.sh. Steps = last observed list price, carried forward between snapshots. ' + extra, fontsize=7, color='#555', wrap=True)
    for sku, fn, title in [('H100 SXM 80GB', 'h100_ondemand_by_provider.png', 'NVIDIA H100 SXM: on-demand list price by provider'),
                           ('H200 141GB', 'h200_ondemand_by_provider.png', 'NVIDIA H200: on-demand list price by provider'),
                           ('B200', 'b200_ondemand_by_provider.png', 'NVIDIA B200: on-demand list price by provider')]:
        cols = [c for c in piv.columns if f'| {sku} | on-demand | ' in c and ('PAYG |' in c + ' |' or 'hourly' in c) and 'community' not in c and 'effective 2026-10-01' not in c]
        if not cols: continue
        fig, ax = plt.subplots(figsize=(11, 5.6))
        x = pd.PeriodIndex(piv.index, freq='M').to_timestamp()
        for c in cols:
            prov = c.split(' | ')[0]
            s = piv[c]
            ff = ffs[c].copy()
            ax.step(x, ff.values, where='post', color=colors.get(prov, 'grey'), lw=2.2 if prov == 'Nebius' else 1.5, label=prov)
            ax.plot(x, s.values, 'o', color=colors.get(prov, 'grey'), ms=3)
        ax.set_ylabel('USD per GPU-hour (list, on-demand)'); ax.set_xlabel('Month')
        ax.set_title(title, loc='left', fontsize=13, fontweight='bold')
        ax.yaxis.set_major_formatter(matplotlib.ticker.StrMethodFormatter('${x:,.2f}'))
        ax.grid(axis='y', alpha=.3); ax.legend(frameon=False, fontsize=9, loc='upper left', bbox_to_anchor=(1.01, 1))
        foot(fig, 'AWS H100 = p5.48xlarge on-demand / 8; CoreWeave = HGX H100 8-GPU instance / 8 (2024 values are a-la-carte GPU component only).')
        fig.tight_layout(rect=(0, 0.05, 1, 1))
        for d in (PARTS, CHARTS): fig.savefig(f'{d}/{fn}', dpi=200)
        plt.close(fig); made.append(fn)
    # Nebius vs neocloud median vs hyperscaler median (H100 on-demand, forward-filled monthly)
    cols = [c for c in piv.columns if '| H100 SXM 80GB | on-demand | ' in c and 'community' not in c and 'effective 2026-10-01' not in c and 'a-la-carte' not in c]
    ff = ffs[cols].copy()
    prov_of = {c: c.split(' | ')[0] for c in cols}
    neb = ff[[c for c in cols if prov_of[c] == 'Nebius']].mean(axis=1)
    neo = ff[[c for c in cols if PTYPE.get(prov_of[c], 'neocloud') == 'neocloud']].median(axis=1)
    hyp = ff[[c for c in cols if PTYPE.get(prov_of[c]) == 'hyperscaler']].median(axis=1)
    comp = pd.DataFrame({'Nebius': neb, 'Neocloud median': neo, 'Hyperscaler median': hyp,
                         'n_neoclouds': ff[[c for c in cols if PTYPE.get(prov_of[c], 'neocloud') == 'neocloud']].notna().sum(axis=1),
                         'n_hyperscalers': ff[[c for c in cols if PTYPE.get(prov_of[c]) == 'hyperscaler']].notna().sum(axis=1)})
    fig, ax = plt.subplots(figsize=(11, 5.6)); x = pd.PeriodIndex(comp.index, freq='M').to_timestamp()
    for c, col, lw in [('Nebius', '#0B5FFF', 2.6), ('Neocloud median', '#29BF12', 1.8), ('Hyperscaler median', '#FF9900', 1.8)]:
        ax.step(x, comp[c].values, where='post', color=col, lw=lw, label=c)
    ax.set_ylabel('USD per GPU-hour (list, on-demand)'); ax.set_xlabel('Month')
    ax.set_title('H100 SXM on-demand: Nebius vs. neocloud and hyperscaler medians', loc='left', fontsize=13, fontweight='bold')
    ax.yaxis.set_major_formatter(matplotlib.ticker.StrMethodFormatter('${x:,.2f}')); ax.grid(axis='y', alpha=.3); ax.legend(frameon=False)
    foot(fig, 'Neoclouds: CoreWeave, Lambda, Crusoe, Together AI, RunPod (Secure Cloud). Hyperscalers: AWS p5 on-demand, Azure ND H100 v5, OCI. Monthly forward-filled between observations.')
    fig.tight_layout(rect=(0, 0.05, 1, 1))
    for d in (PARTS, CHARTS): fig.savefig(f'{d}/h100_nebius_vs_medians.png', dpi=200)
    plt.close(fig); made.append('h100_nebius_vs_medians.png')
    return made, comp

def autosize(ws):
    from openpyxl.utils import get_column_letter
    for i, col in enumerate(ws.columns, 1):
        w = max((len(str(c.value)) if c.value is not None else 0) for c in list(col)[:300])
        ws.column_dimensions[get_column_letter(i)].width = min(max(10, w + 2), 60)
    ws.freeze_panes = 'B2'

def main():
    recs, cov = parse_all()
    for r in recs: r.setdefault('source', 'Wayback snapshot')
    recs += live_records()
    df = to_long(recs)
    piv, hobs = monthly(df)
    ffill = seeded_ffill(df, piv)
    q = quarterly(df, piv)
    dec, launches = decline_rates(df)
    rc = recent_changes(df)
    word = nebius_wording()
    covdf = pd.DataFrame(cov)
    covdf = pd.concat([covdf, pd.DataFrame([
        dict(provider='Nscale', url='nscale.com/pricing, www.nscale.com/pricing, nscale.com/* prefix', month=None, snapshot_timestamp=None, status='no_snapshots', n_prices=0, note='Wayback CDX has no 200 snapshots of any Nscale pricing URL (2024-2026); live page is a JS shell'),
        dict(provider='Vast.ai', url='vast.ai/pricing', month=None, snapshot_timestamp=None, status='see fetchlog', n_prices=0, note='Marketplace prices are loaded client-side from an API; static snapshots carry no prices'),
        dict(provider='Google Cloud', url='cloud.google.com/compute/gpus-pricing', month=None, snapshot_timestamp=None, status='ok', n_prices=0, note='A3 (H100) prices are not on gpus-pricing page (it points to accelerator-optimized VM pricing, which loads tables via JS); no GCP H100 history captured yet'),
        dict(provider='Azure', url='azure.microsoft.com/.../virtual-machines/linux/', month=None, snapshot_timestamp=None, status='ok', n_prices=0, note='Price tables JS-rendered; used Azure Retail Prices API (effectiveStartDate) + instances.vantage.sh snapshots instead'),
        dict(provider='AWS', url='aws.amazon.com/ec2/instance-types/p5/', month=None, snapshot_timestamp=None, status='ok', n_prices=0, note='On-demand price column JS-rendered; used instances.vantage.sh (ec2instances.info) Wayback snapshots for p5.48xlarge on-demand'),
        dict(provider='Oracle OCI', url='oracle.com/cloud/compute/gpu/pricing/', month=None, snapshot_timestamp=None, status='ok', n_prices=0, note='Prices injected by JS from price-list API (part B98415 = H100); API itself archived only Aug/Sep 2026'),
    ])], ignore_index=True)
    # coverage matrix provider x month (H100 SXM)
    h1 = df[(df.SKU == 'H100 SXM 80GB') & df.headline].copy(); h1['month'] = h1.date.dt.to_period('M').astype(str)
    cm = pd.DataFrame(index=[str(m) for m in MONTHS])
    for prov in ['Nebius', 'CoreWeave', 'Lambda', 'Crusoe', 'Together AI', 'RunPod', 'Vast.ai', 'Nscale', 'AWS', 'Azure', 'Google Cloud', 'Oracle OCI']:
        got = set(h1[h1.provider == prov].month)
        snaps = set(covdf[(covdf.provider == prov) & covdf.month.notna()].month)
        cm[prov] = ['H100 price' if m in got else ('snapshot, no H100 price' if m in snaps else 'no snapshot') for m in cm.index]
    made, comp = charts(df, piv, ffill)
    cols = ['date', 'provider', 'provider_type', 'SKU', 'pricing_type', 'term', 'price_basis', 'usd_per_gpu_hr', 'quoted_price', 'quoted_unit', 'gpus_per_instance',
            'bundled_specs', 'region', 'currency', 'source_url', 'wayback_timestamp', 'original_url', 'config', 'headline', 'notes']
    df = df[cols]
    readme = pd.DataFrame({'item': [
        'Scope', 'Price_History_Long', 'Price_History_Monthly', 'Monthly_FFill', 'Price_History_Quarterly', 'Decline_Rates', 'Recent_Changes', 'Nebius_Discount_Wording',
        'H100_Nebius_vs_Medians', 'Coverage', 'Coverage_H100_Matrix', 'Method', 'Caveats', 'Built'],
        'description': [
        'Monthly $/GPU-hr list prices, Jan 2024-Oct 2026, from Wayback snapshots of public pricing pages (+ live 2026-10-04 captures and APIs, labelled). STAGE 1 = H100 SXM fully validated; other SKUs parsed but not yet fully QA-ed (see PROGRESS.md).',
        'LONG raw data: one row per price literally present in a snapshot. No interpolation. headline=TRUE rows feed the pivots (one representative config per provider/SKU/pricing type/term).',
        'Pivot: month x "provider | SKU | pricing_type | term"; value = last observed headline price in that month; blank = no observation.',
        'SEPARATE forward-filled version of the monthly pivot (carry last observation forward). Use for charts/medians only; not raw data.',
        'Q1 2024-Q2 2026; each value tagged Reported / Derived (unit conversion, math shown) / Estimated (carried forward, method stated) / no snapshot.',
        'Annualized % change per series; change 6-9 months after successor-chip availability (dates + sources in lower table).',
        'Jun-Oct 2026: latest pre-Jun and pre-Aug observation vs latest snapshot/live page, with dated 2026 change points.',
        'History of Nebius commitment-discount wording on its pricing page.',
        'Monthly Nebius vs neocloud median vs hyperscaler median, H100 SXM on-demand (forward-filled).',
        'Every snapshot fetched, its status, and why it yielded no price (JS-rendered, no archive, etc.).',
        'Provider x month matrix for H100 SXM.',
        'CDX API (collapse=timestamp:6, statuscode 200) per URL + prefix-variant gap fill; fetch web.archive.org/web/<ts>id_/<url>; parse visible text with per-provider regex parsers (scrapers in /workspace/nebius/market/scrapers/wayback).',
        'Instance prices divided by GPU count (AWS p5 /8, CoreWeave HGX /8, GB200 instance /4). CoreWeave 2024 = a-la-carte GPU component only (vCPU/RAM extra) -> not like-for-like with 2025+ instance price. "Starting at"/"as low as"/"from" prices are floors (price_basis). AWS Capacity Blocks are dynamic short-term reserved rates. Vantage mid-2025 $3.86 p5 value flagged anomalous.',
        dt.datetime.now().strftime('%Y-%m-%d %H:%M ET')]})
    for path in [OUT_MAIN, f'{PARTS}/gpu_price_history.xlsx', f'{PARTS}/price_history.xlsx']:
        with pd.ExcelWriter(path, engine='openpyxl') as xw:
            readme.to_excel(xw, sheet_name='README', index=False)
            d2 = df.copy(); d2['date'] = d2.date.dt.date
            d2.to_excel(xw, sheet_name='Price_History_Long', index=False)
            if path.endswith('price_history.xlsx') and 'gpu_' not in path:
                d2.to_excel(xw, sheet_name='Price_History_Raw', index=False)
            piv.to_excel(xw, sheet_name='Price_History_Monthly')
            ffill.to_excel(xw, sheet_name='Monthly_FFill')
            q.to_excel(xw, sheet_name='Price_History_Quarterly', index=False)
            dec.to_excel(xw, sheet_name='Decline_Rates', index=False)
            launches.to_excel(xw, sheet_name='Decline_Rates', index=False, startrow=len(dec) + 3)
            rc.to_excel(xw, sheet_name='Recent_Changes', index=False)
            word.to_excel(xw, sheet_name='Nebius_Discount_Wording', index=False)
            comp.to_excel(xw, sheet_name='H100_Nebius_vs_Medians')
            covdf.to_excel(xw, sheet_name='Coverage', index=False)
            cm.to_excel(xw, sheet_name='Coverage_H100_Matrix')
            for ws in xw.book.worksheets: autosize(ws)
    d2 = df.copy(); d2['date'] = d2.date.dt.date
    for d in (PARTS, f'{ROOT}/market'):
        d2.to_csv(f'{d}/price_history_raw.csv', index=False); piv.to_csv(f'{d}/price_history_monthly.csv')
        ffill.to_csv(f'{d}/price_history_monthly_ffill.csv'); q.to_csv(f'{d}/price_history_quarterly.csv', index=False)
        rc.to_csv(f'{d}/price_history_recent_changes.csv', index=False)
    print('rows', len(df), 'headline', int(df.headline.sum()), 'series', piv.shape[1], 'charts', made)
    return df, piv, comp, rc, dec, cm

if __name__ == '__main__':
    main()
