"""Per-provider parsers for archived GPU pricing pages.
Each parser takes the visible text of a snapshot and returns a list of dicts:
chip, config, pricing_type, term, quoted_price, quoted_unit, gpus_per_unit, headline(bool), notes
Only prices literally present in the text are emitted."""
import re
VAL = r'(?:from \$ ?[\d,]+(?:\.\d+)?|\$ ?[\d,]+(?:\.\d+)?|Contact (?:sales|Sales|us)|N/A|—|––|-)'
def money(s):
    m = re.search(r'\$ ?([\d,]+(?:\.\d+)?)', s or '')
    return float(m.group(1).replace(',', '')) if m else None

def norm_chip(name):
    n = name.upper()
    for k in ['GB300', 'GB200', 'GH200', 'B300', 'B200', 'H200', 'MI355X', 'MI300X', 'RTX PRO 6000', 'L40S', 'A100']:
        if k in n:
            if k == 'A100': return 'A100 80GB' if '80' in n else ('A100 40GB' if '40' in n else 'A100')
            return k
    if 'H100' in n:
        return 'H100 PCIe' if 'PCIE' in n else 'H100'
    if 'L40' in n: return 'L40'
    return None

def rec(chip, config, ptype, term, qp, unit, gpus=1, headline=True, notes=''):
    return dict(chip=chip, config=config, pricing_type=ptype, term=term, quoted_price=qp,
                quoted_unit=unit, gpus_per_unit=gpus, headline=headline, notes=notes)

def vals_after(t, pos, n):
    out = []
    for m in re.finditer(r'\s*(' + VAL + r')(?:\s*/(?:GPU-hr|hr|GPU / hr| GPU / hr))?', t[pos:]):
        if m.start() != 0 and t[pos:pos + m.start()].strip():
            break
        out.append(m.group(1)); pos2 = m.end(); 
        t2 = t[pos:]; 
        if len(out) >= n: break
        # continue matching contiguous tokens
        t = t; pos = pos + m.end()
        return_list = None
        # restart scanning from new pos
        rest = vals_after(t, pos, n - len(out))
        return out + rest
    return out

# ---------------- NEBIUS ----------------
NEB_ROW = re.compile(r'NVIDIA (?:HGX |GB)?(?P<name>(?:GB)?(?:H100|H200|B200|B300|GB200 NVL72|GB300 NVL72|200 NVL72|300 NVL72|RTX PRO 6000|L40S(?: GPU)? with (?:AMD|Intel)(?: CPU)?))\*?(?: GPU)? (?P<specs>(?:[\d]+(?:-\d+)? ){1,3})')
def nebius(t):
    out = []
    # Layout A (nebius.ai 2023-24): GPU type: H100 SXM5. PAYG price – $4.85/h ... 3/6/12-month reserve
    for m in re.finditer(r'GPU type: (?P<g>[A-Z0-9 ]+?)\. PAYG price – \$(?P<p>[\d.]+)/h(?P<rest>.{0,400}?)(?=GPU type:|Default|Need more|Choose a service|About reserving)', t):
        chip = norm_chip(m.group('g')); cfg = m.group('g')
        out.append(rec(chip, cfg, 'on-demand', 'PAYG', '$' + m.group('p'), 'per GPU-hour', notes='PAYG price'))
        rest = m.group('rest')
        hdr_terms = re.findall(r'(\d+)-(month|year) reserve', rest[:200])
        terms = [f"{int(n)*12 if u == 'year' else int(n)}-month" for n, u in hdr_terms][:3] if len(hdr_terms) >= 3 else ['3-month', '6-month', '12-month']
        r = re.search(r'(up to 8|1 GPU/h) \$([\d.]+) \$([\d.]+) \$([\d.]+)', rest)
        if r:
            tier = 'volume tier up to 8 GPUs' if r.group(1) == 'up to 8' else 'per GPU'
            for term, v in zip(terms, r.groups()[1:]):
                out.append(rec(chip, cfg, 'reserved', term, '$' + v, 'per GPU-hour', notes=f'reserve price, {tier}'))
        r = re.search(r'up to 128 \$([\d.]+) \$([\d.]+) \$([\d.]+)', rest)
        if r:
            for term, v in zip(terms, r.groups()):
                out.append(rec(chip, cfg, 'reserved', term, '$' + v, 'per GPU-hour', headline=False, notes='reserve price, volume tier up to 128 GPUs'))
    # Layout B (nebius.com Nov24-Jan25): 1 × H100 SXM 80 200 16 $2.95 $2.00 (On demand, Reserve)
    for m in re.finditer(r'(?P<n>[18]) × (?P<g>H100 SXM|H200 SXM|L40s PCIe) [\d ]+? \$(?P<od>[\d.]+) \$(?P<rs>[\d.]+)', t):
        n = int(m.group('n')); chip = norm_chip(m.group('g'))
        hl = n == 1
        out.append(rec(chip, f"{n}x {m.group('g')}", 'on-demand', 'PAYG', '$' + m.group('od'), f'per {n}-GPU VM-hour', n, hl))
        out.append(rec(chip, f"{n}x {m.group('g')}", 'reserved', 'commit ("as low as", large volume)', '$' + m.group('rs'), f'per {n}-GPU VM-hour', n, hl,
                       'Reserve column; page: commitment period from 3 months to 3 years'))
    # "Reserve your best price" cards (2025+): NVIDIA H100 GPU $2.00 / hour ... commitment of hundreds of units
    for m in re.finditer(r'NVIDIA (H100|H200) GPU \$([\d.]+) / hour (?:Access|Receive)[^.]*?commitment of hundreds of units(?P<min> for at least 3 months)?', t):
        wording = 'commitment of hundreds of units for at least 3 months' if m.group('min') else 'commitment of hundreds of units (no minimum term stated)'
        out.append(rec(m.group(1), f'NVIDIA {m.group(1)} reserve card', 'reserved', 'commit ("as low as", large volume)', '$' + m.group(2), 'per GPU-hour', notes='page wording: ' + wording))
    # Instance tables (2025-03 onward)
    for sec in re.finditer(r'(On-demand GPU pricing|NVIDIA GPU Instances|Preemptible NVIDIA GPU Instances)(?P<body>.*?)(?=On-demand CPU pricing|CPU-only instances|Preemptible NVIDIA GPU Instances|Storage Item|$)', t):
        title, body = sec.group(1), sec.group('body')
        hdr = body[:body.find('NVIDIA')] if 'NVIDIA' in body else body
        if title.startswith('Preemptible') or 'Spot price' in hdr:
            cols = [('spot', 'preemptible')]
        elif 'Preemptible, GPU-hour On-demand, GPU-hour' in hdr:
            cols = [('spot', 'preemptible'), ('on-demand', 'PAYG')]
        elif 'Effective October 1, 2026' in hdr:
            cols = [('on-demand', 'PAYG'), ('on-demand', 'PAYG (announced, effective 2026-10-01)')]
        else:
            cols = [('on-demand', 'PAYG')]
        for m in NEB_ROW.finditer(body):
            name = m.group('name'); chip = norm_chip('GB' + name if name.startswith(('200 N', '300 N')) else name)
            if chip is None: continue
            vs = re.findall(VAL, body[m.end():m.end() + 80])
            seq = []
            for v in re.finditer(r'\s*(' + VAL + r')\*{0,2}', body[m.end():]):
                if v.start() != 0: break
                seq.append(v.group(1))
                if len(seq) == len(cols): break
                body_rest = None
            # contiguous scan
            seq = []; pos = m.end()
            while len(seq) < len(cols):
                v = re.match(r'\s*(' + VAL + r')\*{0,2}', body[pos:])
                if not v: break
                seq.append(v.group(1)); pos += v.end()
            for (pt, term), v in zip(cols, seq):
                cfg = 'NVIDIA ' + m.group('name')
                hl = 'L40S' not in name or 'Intel' in name
                if money(v) is None:
                    out.append(rec(chip, cfg, pt, term, v, 'per GPU-hour', headline=False, notes='no list price shown'))
                else:
                    note = '"from" price (floor)' if v.startswith('from') else ''
                    out.append(rec(chip, cfg, pt, term, v, 'per GPU-hour', headline=hl, notes=note))
    return out

# ---------------- COREWEAVE ----------------
CW_NAMES = r'(NVIDIA GB300 NVL72|NVIDIA GB200 NVL72|NVIDIA HGX B300|NVIDIA HGX B200|NVIDIA HGX H100|NVIDIA HGX H200|NVIDIA H100 PCIe|NVIDIA GH200|NVIDIA L40S|NVIDIA L40|NVIDIA A100|A100 80GB NVLINK|NVIDIA RTX PRO 6000 Blackwell Server Edition \(High Memory\))'
def coreweave(t):
    out = []
    # 2024 a-la-carte GPU component table
    if 'GPU Component Cost Per Hour' in t:
        body = t[t.find('GPU Component Cost Per Hour'):]
        body = body[:body.find('*Compute instances')]
        for m in re.finditer(r'(NVIDIA HGX H100|NVIDIA H100 PCIe|A100 80GB NVLINK|A100 80GB PCIe|A100 40GB NVLINK) (\d+) (\d+) (\d+) \$([\d.]+)', body):
            chip = norm_chip(m.group(1))
            out.append(rec(chip, m.group(1), 'on-demand', 'PAYG (a-la-carte GPU component only)', '$' + m.group(5), 'per GPU-hour (GPU component only)', 1, True,
                           'a-la-carte GPU component; excludes vCPU ($0.01/hr) and RAM ($0.005/GB-hr) which are billed separately'))
        return out
    us = t.split('REGION: EUROPE')[0]
    seen = set()
    for m in re.finditer(CW_NAMES + r' (?P<cnt>\d+|4\^1) (?P<vram>\d+) (?P<vcpu>\d+) (?P<ram>[\d,]+) (?P<ls>[\d.]+) (?P<p>\$[\d,.]+|Contact sales)', us):
        name = m.group(1)
        if name in seen: continue
        seen.add(name)
        cnt = 4 if m.group('cnt') == '4^1' else int(m.group('cnt'))
        chip = norm_chip(name)
        note = 'instance price; GB200/GB300 instance = 4 GPUs (1 tray of NVL72), per page footnote' if cnt == 4 else 'instance price (8-GPU node), US region'
        p = m.group('p')
        out.append(rec(chip, f'{name} x{cnt}', 'on-demand', 'PAYG', p, f'per {cnt}-GPU instance-hour', cnt, money(p) is not None and 'PCIe' not in name, note))
        sm = re.search(re.escape(name) + r' On-Demand Price: (?:\$[\d,.]+ )?/ Hour Spot Price: (\$[\d,.]+|N/A) / Hour', us)
        if sm and money(sm.group(1)) is not None:
            out.append(rec(chip, f'{name} x{cnt}', 'spot', 'spot', sm.group(1), f'per {cnt}-GPU instance-hour', cnt, True, 'spot price, US region'))
    return out

# ---------------- LAMBDA ----------------
LG = r'(B200 SXM6|B300 SXM6|B300|H200 SXM|H200|H100 SXM|H100 PCIe|GH200|A100 SXM|A100 PCIe)'
def _lchip(g, vram):
    return norm_chip(g + (f' {vram}GB' if 'A100' in g else ''))
def lambda_(t):
    out = []
    def add(chip, cfg, pt, term, qp, unit, gpus, hl, note):
        if any(o['config'] == cfg and o['pricing_type'] == pt and o['term'] == term for o in out): return
        out.append(rec(chip, cfg, pt, term, qp, unit, gpus, hl, note))
    # 2022-2023 layout (lambdalabs.com): '1x NVIDIA A100 40 GB 30 200 GiB 512 GiB $1.10 / hr' (no SXM/PCIe label, no 'SSD');
    # on-demand table first, then 'Reserved pricing' (3-month minimum) table
    if re.search(r'\dx NVIDIA A100 \d+ GB \d+ [\d.]+ [GT]iB [\d.]+ [GT]iB \$', t):
        cut = re.search(r'Reserved (?:pricing|Cloud)', t); od, rs = (t[:cut.start()], t[cut.start():]) if cut else (t, '')
        for part, pt, term in [(od, 'on-demand', 'PAYG'), (rs[:1500], 'reserved', '3-month minimum (Reserved pricing)')]:
            for m in re.finditer(r'(\d)x NVIDIA (A100|H100 PCIe|H100 SXM)(?: New)? (\d+) GB \d+ [\d.]+ [GT]iB [\d.]+ [GT]iB \$([\d.]+) / hr', part):
                n = int(m.group(1)); g = m.group(2); vram = m.group(3)
                chip = norm_chip(g + (f' {vram}GB' if g == 'A100' else ''))
                if chip is None: continue
                hl = (n == 8) or (g == 'A100' and vram == '40' and n == 1 and pt == 'on-demand' and not re.search(r'8x NVIDIA A100 40 GB', part)) or chip == 'H100 PCIe'
                add(chip, f'{n}x NVIDIA {g} {vram}GB', pt, term, '$' + m.group(4), f'per {n}-GPU instance-hour', n, hl,
                    f'{n}x instance, instance price (2022-23 lambdalabs layout)' + ('; reserved = minimum 3-month commitment' if pt == 'reserved' else ''))
        for m in re.finditer(r'NVIDIA A100s now available on-demand for \$([\d.]+)/hr', t):
            pass
        if out: return out
    # Jan 2024: 2023 Price / 2024 Price (effective 1/1/2024), per-instance
    if '2024 Price (effective 1/1/2024)' in t:
        for m in re.finditer(r'(\d)x NVIDIA ' + LG + r'(?: New)? (\d+) GB .*? SSD \$([\d.]+) / hr \$([\d.]+) / hr', t):
            n = int(m.group(1)); chip = _lchip(m.group(2), m.group(3))
            add(chip, f'{n}x NVIDIA {m.group(2)} {m.group(3)}GB', 'on-demand', 'PAYG', '$' + m.group(5), f'per {n}-GPU instance-hour', n,
                n == 8 or chip == 'H100 PCIe', f'2024 price (effective 1/1/2024); page also shows 2023 price ${m.group(4)}/hr')
        return out
    # 2024-2025 rows with explicit Nx prefix (per-instance '/ hr' or per-GPU '/ GPU / hr' or bare per-GPU)
    for m in re.finditer(r'(?:(On-demand|Reserved) )?(\d)x NVIDIA ' + LG + r'(?: New(?: Price)?)? (\d+) GB \d+ [\d.]+ GiB [\d.]+ [TG]iB SSD (\$[\d.]+|CONTACT SALES)( / GPU / hr| / hr)?', t):
        kind, n, g, vram, p, unit = m.groups(); n = int(n)
        if money(p) is None or kind == 'Reserved': continue
        chip = _lchip(g, vram)
        if unit == ' / hr':
            add(chip, f'{n}x NVIDIA {g} {vram}GB', 'on-demand', 'PAYG', p, f'per {n}-GPU instance-hour', n, n == 8 or chip in ('H100 PCIe', 'GH200'), f'{n}x instance, instance price')
        else:
            add(chip, f'{n}x NVIDIA {g} {vram}GB', 'on-demand', 'PAYG', p, 'per GPU-hour', 1, n == 8 or chip in ('H100 PCIe', 'GH200'), f'{n}x instance')
    # lambda.ai 2025-26 tabbed tables
    hdr = 'VRAM/GPU vCPUs RAM STORAGE PRICE/GPU/HR*'
    if hdr in t and not out:
        tabs = re.search(r'((?:\d+x ){2,4})(?:Plan )?VRAM/GPU', t)
        sizes = [int(x[:-1]) for x in tabs.group(1).split()] if tabs else []
        tables = t.split(hdr)[1:]
        for n, tb in zip(sizes, tables):
            tb = tb.split('* plus')[0]
            for m in re.finditer(r'NVIDIA ' + LG + r' (\d+) GB \d+ [\d.]+ GiB [\d.]+ [TG]iB SSD (\$[\d.]+)', tb):
                g, vram, p = m.groups(); chip = _lchip(g, vram)
                add(chip, f'{n}x NVIDIA {g} {vram}GB', 'on-demand', 'PAYG', p, 'per GPU-hour', 1, n == 8 or chip in ('H100 PCIe', 'GH200'), f'{n}x instance')
    # 1-Click Clusters, layout 1: NVIDIA H100 2 weeks – 1 year 16 $6.16
    for m in re.finditer(r'NVIDIA (HGX B200|H100|HGX B300|B300|HGX H200) (\d+ (?:weeks?|days?) – \d+ (?:year|months?|weeks?)|1 year\+) (\d+\+?) (\$[\d.]+|—)', t):
        g, dur, cnt, p = m.groups()
        if money(p) is None: continue
        add(norm_chip(g), f'1-Click Cluster NVIDIA {g}, {cnt} GPUs', 'reserved', f'1-Click Cluster {dur}', p, 'per GPU-hour', 1,
            cnt.startswith('16'), '1-Click Clusters')
    # 1-Click Clusters, layout 2: Commitment / As low as (per GPU-hour)
    for sm in re.finditer(r'NVIDIA (HGX B200|H100|HGX B300|B300|HGX H200) systems[^*]*?Commitment As low as \(per GPU-hour\)(.*?)\* plus', t):
        g = sm.group(1)
        for m in re.finditer(r'(On-Demand|Reserved) ((?:\d+ (?:weeks?|months?|years?)(?:-\d+ (?:weeks?|months?|years?))?)) (\$[\d.]+)', sm.group(2)):
            kind, dur, p = m.groups()
            pt = 'on-demand' if kind == 'On-Demand' else 'reserved'
            add(norm_chip(g), f'1-Click Cluster NVIDIA {g}', pt if pt == 'reserved' else 'reserved', f'1-Click Cluster {kind} {dur}', p, 'per GPU-hour', 1, True,
                '1-Click Clusters "as low as"; on-demand cluster still requires min. commitment term shown')
    return out

# ---------------- TOGETHER ----------------
TG_HEADERS = {
 'Hourly 1 - 6 Days 1 Week - 3 Months': [('on-demand','hourly'),('reserved','1-6 days'),('reserved','1 week-3 months')],
 '1 Week - 3 Months 1 - 6 Days Hourly': [('reserved','1 week-3 months'),('reserved','1-6 days'),('on-demand','hourly')],
 '1 Week - 3 Months* 1 - 6 Days* Hourly*': [('reserved','1 week-3 months'),('reserved','1-6 days'),('on-demand','hourly')],
 'Hourly': [('on-demand','hourly')],
 '1 Week - 1 Month 2 - 3 Months 4 - 6 Months 6+ Months': [('reserved','1 week-1 month'),('reserved','2-3 months'),('reserved','4-6 months'),('reserved','6+ months')],
 '7-30 days 31-90 days 91-180 days 181+ days': [('reserved','7-30 days'),('reserved','31-90 days'),('reserved','91-180 days'),('reserved','181+ days')],
 'ON-Demand Reserved Pay as you go 7-30 days 31-90 days 91-180 days 181+ days': [('on-demand','hourly'),('reserved','7-30 days'),('reserved','31-90 days'),('reserved','91-180 days'),('reserved','181+ days')],
 'Preemptible Compute ON-Demand Reserved Pay as you go Pay as you go 7-30 days 31-90 days 91-180 days 181+ days': [('spot','preemptible'),('on-demand','hourly'),('reserved','7-30 days'),('reserved','31-90 days'),('reserved','91-180 days'),('reserved','181+ days')],
}
def together(t):
    out = []
    # GPU cluster tables
    for h in re.finditer(r'Hardware (?P<hdr>(?:(?!NVIDIA|Hardware).){1,140}?) (?=NVIDIA)', t):
        hdr = h.group('hdr').strip()
        if hdr.startswith('All prices per gpu per hour On-demand Pay as you go Reserved'):
            cols = [('dedicated-endpoint', 'on-demand hourly'), ('dedicated-endpoint', 'reserved')]
        elif hdr.startswith('GPU Memory Price'):
            cols = 'starting'
        else:
            cols = TG_HEADERS.get(hdr)
        if cols is None:
            out.append(rec(None, None, 'PARSE_WARNING', hdr, None, None, notes='unknown Together header')); continue
        body = t[h.end():h.end() + 1500]
        stop = re.search(r'(Reserved (Reserve|Clusters)|On-demand hourly rates|Sandbox|Frontier AI|GPU Clusters|Instant Clusters|\* plus|Storage|Hardware )', body)
        body = body[:stop.start()] if stop else body
        if cols == 'starting':
            for m in re.finditer(r'NVIDIA (?:HGX )?(GB200 NVL72|B200|H200|H100|A100) \d+GB HBM\w+ (Starting at \$[\d.]+|Contact us)', body):
                if money(m.group(2)) is None: continue
                out.append(rec(norm_chip(m.group(1) + (' 80GB' if m.group(1) == 'A100' else '')), f'Reserved cluster NVIDIA {m.group(1)}', 'reserved', 'reserved cluster, term unspecified ("starting at")',
                               m.group(2), 'per GPU-hour', 1, True, 'Reserved Clusters "Starting at" price'))
            continue
        for m in re.finditer(r'NVIDIA (?:HGX )?(H100 SXM|H100 Inference|H100|H200|B200|B300|GB200 NVL72|GB300 NVL72)', body):
            pos = m.end(); seq = []
            while len(seq) < len(cols):
                v = re.match(r'\s*(' + VAL + r')', body[pos:])
                if not v: break
                seq.append(v.group(1)); pos += v.end()
            g = m.group(1); chip = norm_chip(g)
            if len(seq) < len(cols) and len(seq) >= 1:
                # rows like 'GB200 NVL72 — Contact us' are shorter; only keep if numeric prices align from the left
                pass
            for (pt, term), v in zip(cols, seq):
                if money(v) is None: continue
                cfg = f'GPU Clusters NVIDIA HGX {g}' if pt != 'dedicated-endpoint' else f'Dedicated Inference NVIDIA HGX {g}'
                out.append(rec(chip, cfg, pt, term, v, 'per GPU-hour', 1, 'Inference' not in g, 'Instant/GPU Clusters' if pt != 'dedicated-endpoint' else 'Dedicated Inference'))
    # 2024-early 2025 GPU Clusters cards: 'H100 80GB price 1k tokens ... Starting at $1.99/hr'
    for m in re.finditer(r'(A100 SXM 80GB|A100 PCIe 80GB|H100 80GB|H200 141GB|NVIDIA H100|NVIDIA H200|NVIDIA A100) price 1k tokens (?:(?!Starting|NVIDIA|H100 80GB|H200 141GB).){0,120}?Starting at (\$[\d.]+)/hr', t):
        g = m.group(1); chip = norm_chip(g + (' 80GB' if 'A100' in g and '80' not in g else ''))
        out.append(rec(chip, f'Together GPU Clusters {g}', 'reserved', 'reserved cluster, term unspecified ("starting at")', m.group(2), 'per GPU-hour', 1, True, 'GPU Clusters "Starting at" price'))
    # Dedicated endpoints per-hour list: 1x H100 80GB $3.36
    for m in re.finditer(r'1x (?:HGX )?(H100|H200|B200) \d+ ?GB (?:price/minute \$[\d.]+ price/hour )?(\$[\d.]+)(?: (\$[\d.]+))?', t):
        g, a, b = m.groups()
        if b:  # per-minute then per-hour columns
            out.append(rec(g, f'Dedicated endpoint 1x {g}', 'dedicated-endpoint', 'on-demand', b, 'per GPU-hour', 1, False, f'dedicated inference endpoint; also {a}/min'))
        elif float(a.strip('$')) < 0.5:
            out.append(rec(g, f'Dedicated endpoint 1x {g}', 'dedicated-endpoint', 'on-demand', a, 'per GPU-minute', 1, False, 'dedicated inference endpoint, per-minute price'))
        else:
            out.append(rec(g, f'Dedicated endpoint 1x {g}', 'dedicated-endpoint', 'on-demand', a, 'per GPU-hour', 1, False, 'dedicated inference endpoint'))
    # dedupe
    u = []; seen = set()
    for o in out:
        k = (o['chip'], o['config'], o['pricing_type'], o['term'], o['quoted_price'])
        if k not in seen: seen.add(k); u.append(o)
    return u

# ---------------- CRUSOE ----------------
def crusoe(t):
    out = []
    for h in re.finditer(r'(?:GPU Compute|GPU Pricing|GPU VM pricing) (?P<hdr>On-Demand(?: Current)?(?: Spot)? 6-month reserved 1-year reserved 3-year reserved)', t):
        hdr = h.group('hdr')
        cols = [('on-demand', 'PAYG')] + ([('spot', 'spot')] if 'Spot' in hdr else []) + [('reserved', '6-month'), ('reserved', '1-year'), ('reserved', '3-year')]
        body = t[h.end():h.end() + 3000]
        body = re.split(r'(?:CPU Compute|CPU VM pricing|GPU Pricing NVIDIA)', body)[0]
        for m in re.finditer(r'(NVIDIA|AMD) ((?:[A-Z0-9]+ )?(?:Tensor Core )?(?:GPU )?[\w ]*?\d+GB(?: (?:SXM|PCIe|OAM))?)', body):
            pos = m.end(); seq = []
            while len(seq) < len(cols):
                v = re.match(r'\s*(' + VAL + r')(?: /hr)?', body[pos:])
                if not v: break
                seq.append(v.group(1)); pos += v.end()
            if len(seq) < len(cols): continue
            name = m.group(1) + ' ' + m.group(2); chip = norm_chip(name)
            if chip is None: continue
            for (pt, term), v in zip(cols, seq):
                if money(v) is None: continue
                out.append(rec(chip, name, pt, term, v.replace('$ ', '$'), 'per GPU-hour', 1, True, ''))
    for h in re.finditer(r'GPU model On-demand Current spot', t):
        body = t[h.end():h.end() + 1500].split('CPU instances')[0]
        for m in re.finditer(r'(NVIDIA|AMD) ([\w ]+?\d+GB(?: (?:SXM|PCIe|HGX|NVL72|OAM))?) (' + VAL + r')(?:/GPU-hr)? (' + VAL + r')(?:/GPU-hr)?', body):
            name = m.group(1) + ' ' + m.group(2); chip = norm_chip(name)
            for pt, v in [('on-demand', m.group(3)), ('spot', m.group(4))]:
                if money(v) is None: continue
                out.append(rec(chip, name, pt, 'PAYG' if pt == 'on-demand' else 'spot', v, 'per GPU-hour', 1, True, ''))
    for h in re.finditer(r'Self-Serve Deployments pricing.*?GPU type Price per hour', t):
        body = t[h.end():h.end() + 400].split('Tailored')[0]
        for m in re.finditer(r'(NVIDIA|AMD) ([\w ]+?\d+GB(?: HGX)?) (\$[\d.]+)', body):
            name = m.group(1) + ' ' + m.group(2)
            out.append(rec(norm_chip(name), 'Self-serve dedicated endpoint ' + name, 'dedicated-endpoint', 'on-demand', m.group(3), 'per GPU-hour', 1, False, 'Crusoe Managed Inference self-serve deployment'))
    u = []; seen = set()
    for o in out:
        k = (o['chip'], o['config'], o['pricing_type'], o['term'])
        if k not in seen: seen.add(k); u.append(o)
    return u

# ---------------- AWS EC2 Capacity Blocks ----------------
def aws(t):
    out = []; seen = set()
    for m in re.finditer(r'(p5\.48xlarge|p5e\.48xlarge|p5en\.48xlarge|p6-b200\.48xlarge|p6e-gb200\.36xlarge|p6-b300\.48xlarge|p4d\.24xlarge|p4de\.24xlarge|p5\.4xlarge|u-p6e-gb200x72|u-p6e-gb200x36) (US East \(N\. Virginia\)|US East \(Ohio\)|US West \(Oregon\)) \$([\d,.]+) USD \(?\$([\d,.]+) USD\)? (\d+) x (\w+)', t):
        inst, region, ip, gp, n, g = m.groups()
        if (inst, region) in seen: continue
        seen.add((inst, region))
        chip = norm_chip(g + (' 40GB' if inst == 'p4d.24xlarge' else ' 80GB' if inst == 'p4de.24xlarge' else ''))
        out.append(rec(chip, f'{inst} ({n}x {g}), {region}', 'capacity-block', 'EC2 Capacity Block (1 day-6 months, prepaid)', '$' + gp, 'per GPU-hour (accelerator hourly rate)', 1,
                       region == 'US East (Ohio)' or region == 'US East (N. Virginia)', f'instance rate ${ip}/hr'))
    # keep one headline per instance: prefer N. Virginia, else Ohio, else Oregon
    by = {}
    for o in out:
        inst = o['config'].split(' ')[0]; by.setdefault(inst, []).append(o)
    for inst, lst in by.items():
        order = ['US East (N. Virginia)', 'US East (Ohio)', 'US West (Oregon)']
        lst.sort(key=lambda o: min([i for i, r in enumerate(order) if r in o['config']] or [9]))
        for i, o in enumerate(lst): o['headline'] = (i == 0) and inst not in ('p5.4xlarge', 'p5en.48xlarge', 'p4d.24xlarge')
    return out

PARSERS = {'Nebius': nebius, 'CoreWeave': coreweave, 'Lambda': lambda_, 'Together AI': together, 'Crusoe': crusoe, 'AWS': aws}

# ---------------- RUNPOD ----------------
RP_OLD = {'A100 80GB': 'A100 80GB PCIe', 'A100 SXM 80GB': 'A100 80GB', 'H100 80GB SXM5': 'H100', 'H100 80GB PCIe': 'H100 PCIe', 'H100 80GB HBM3': 'H100', 'L40S': 'L40S', 'L40': 'L40'}
def runpod_old(t):
    """2022-Apr 2024 www.runpod.io/gpu-instance/pricing layouts. First table = Secure Cloud, second = Community Cloud."""
    out = []; cnt = {}
    names = r'(A100 SXM 80GB|A100 80GB|H100 80GB SXM5|H100 80GB PCIe|H100 80GB HBM3|L40S|L40)'
    pats = [r'(?P<n>\d)x ' + names + r' \d+ (?:GB )?(?:\d+ )?\d* ?GB (?P<spot>[\d.]+) (?P<od>[\d.]+)',                     # 2022-07/08: vCPU RAM spot od
            r'(?P<n>\d)x ' + names + r' \d+ GB - - (?P<od>[\d.]+)',                                                      # 2022-10..2023-07
            r'(?P<n>)' + names + r' \d+ GB VRAM (?:\d+ )?GB RAM (?:\d+ )?vCPU Spot (?P<spot>[\d.]+|-) On-Demand (?P<od>[\d.]+)']  # 2023-08..2024-04
    for pat in pats:
        for m in re.finditer(pat, t):
            n = int(m.group('n') or 1)
            if n != 1: continue
            g = m.group(2); chip = RP_OLD[g]
            k = cnt.get(g, 0); cnt[g] = k + 1
            cloud = 'Secure Cloud' if k == 0 else 'Community Cloud'
            if k > 1: continue
            od = m.group('od'); spot = m.groupdict().get('spot')
            if float(od) <= 0.5 and chip in ('L40S', 'H100'): continue
            out.append(rec(chip, f'RunPod {g} ({cloud})', 'on-demand', 'PAYG' if k == 0 else 'PAYG (community cloud)', '$' + od, 'per GPU-hour', 1, k == 0,
                           f'{cloud}; legacy gpu-instance/pricing page layout'))
            if spot and spot not in ('-',) and float(spot) > 0:
                out.append(rec(chip, f'RunPod {g} ({cloud})', 'spot', 'spot (interruptible)' + ('' if k == 0 else ', community cloud'), '$' + spot, 'per GPU-hour', 1, k == 0,
                               f'{cloud} spot/interruptible; legacy layout'))
        if out: break
    return out
def runpod(t):
    out = []; seen = set()
    if 'GB VRAM' not in t or re.search(r'Spot (?:[\d.]+|-) On-Demand', t) or re.search(r'\dx A100 (?:SXM )?80GB \d+ GB', t):
        o = runpod_old(t)
        if o: return o
    for m in re.finditer(r'(H100 SXM|H100 PCIe|H100 NVL|H200 SXM|H200 NVL|H200|B200|B300|A100 SXM|A100 PCIe|L40S|MI300X) (\d+) GB VRAM(?: (\d+) GB RAM (\d+) vCPUs)? \$ ([\d.]+) /hr(?: Secure Cloud)?(?: \$ ([\d.]+) /hr Community Cloud)?', t):
        g, vram, ram, vcpu, sec, com = m.groups()
        if g in seen or float(sec) == 0: continue
        seen.add(g)
        chip = norm_chip(g + (' 80GB' if 'A100' in g else ''))
        if g == 'H100 NVL': chip = 'H100 NVL'
        if g == 'A100 PCIe': chip = 'A100 80GB PCIe'
        specs = f'{vram}GB VRAM, {ram}GB RAM, {vcpu} vCPUs per GPU' if ram else f'{vram}GB VRAM'
        out.append(rec(chip, f'RunPod {g} (Secure Cloud)', 'on-demand', 'PAYG', '$' + sec, 'per GPU-hour', 1, True, 'Secure Cloud; ' + specs))
        if com:
            out.append(rec(chip, f'RunPod {g} (Community Cloud)', 'on-demand', 'PAYG (community cloud)', '$' + com, 'per GPU-hour', 1, False, 'Community Cloud (third-party hosts); ' + specs))
    return out

# ---------------- ORACLE (cetools price-list API JSON) ----------------
def oracle_api(rawjson):
    import json
    out = []
    try: d = json.loads(rawjson)
    except Exception: return out
    for it in d.get('items', []):
        n = it.get('displayName', '')
        if 'GPU' not in n or 'Per Hour' not in (it.get('metricName') or ''): continue
        chip = None
        for k, c in [('GB300', 'GB300'), ('GB200', 'GB200'), ('B300', 'B300'), ('B200', 'B200'), ('H200', 'H200'), ('H100T', None), ('H100', 'H100'), ('A100 - v2', 'A100 80GB'), ('MI300X', 'MI300X')]:
            if k in n: chip = c; break
        if not chip: continue
        for loc in it.get('currencyCodeLocalizations', []):
            if loc.get('currencyCode') != 'USD': continue
            for pr in loc.get('prices', []):
                if pr.get('model') == 'PAY_AS_YOU_GO' or 'value' in pr:
                    out.append(rec(chip, f"OCI {it.get('partNumber')} {n}", 'on-demand', 'PAYG', f"${pr['value']}", 'per GPU-hour', 1, True, 'OCI price-list API, list price (bare metal shapes, e.g. BM.GPU.H100.8)'))
                    break
    return out

PARSERS.update({'RunPod': runpod})

# ---------------- Vantage instances.vantage.sh (ec2instances.info) ----------------
VANTAGE_INST = {'p5.48xlarge': ('H100', 8, 'AWS', '192 vCPU, 2048 GiB RAM, 3200 Gbps EFAv2, 8x3.84TB NVMe'),
                'nd96isr-h100-v5': ('H100', 8, 'Azure', '96 vCPU, 1900 GiB RAM, 8x400 Gb/s InfiniBand (ND H100 v5)')}
def vantage(t, inst):
    out = []
    chip, n, prov, specs = VANTAGE_INST[inst]
    m = re.search(r'starting at \$([\d,.]+) per hour', t)
    if not m: return out
    od = float(m.group(1).replace(',', ''))
    anomalous = od < 1.25 * n  # e.g. an 8-GPU p5.48xlarge at <$10/hr is not a real on-demand list price (mid-2025 data glitch on vantage)
    out.append(rec(chip, f'{inst} ({n}x {chip})', 'on-demand', 'PAYG', f'${m.group(1)}', f'per {n}-GPU instance-hour', n, not anomalous,
                   ('ANOMALY: implausible value shown by tracker, excluded from headline; ' if anomalous else '') + f'tracker "starting at" = lowest-region Linux on-demand; {specs}'))
    m2 = re.search(r'Pricing \$([\d,.]+) On Demand (\$[\d,.]+|N/A) Spot (\$[\d,.]+|N/A) 1-Year Reserved (\$[\d,.]+|N/A) 3-Year Reserved', t)
    if m2:
        for pt, term, v in [('spot', 'spot (snapshot)', m2.group(2)), ('reserved', '1-year', m2.group(3)), ('reserved', '3-year', m2.group(4))]:
            if money(v) is None: continue
            out.append(rec(chip, f'{inst} ({n}x {chip})', pt, term, v, f'per {n}-GPU instance-hour', n, not anomalous,
                           f'tracker headline figure (lowest region); reserved = no-upfront effective hourly as displayed by tracker; {specs}'))
    return out

# ---------------- Google Cloud ----------------
GCP_MT = {'a3-highgpu-8g': ('H100', 8, 'A3 High: 8x H100 80GB, 208 vCPU, 1,872 GB RAM, 6 TB local SSD, 8x200 Gbps GPUDirect-TCPX'),
          'a3-megagpu-8g': ('H100', 8, 'A3 Mega: 8x H100 80GB, 208 vCPU, 1,872 GB RAM, 6 TB local SSD, 8x200 Gbps GPUDirect-TCPXO'),
          'a3-ultragpu-8g': ('H200', 8, 'A3 Ultra: 8x H200 141GB, 224 vCPU, 2,952 GB RAM, 12 TB local SSD, 3.2 Tbps RDMA'),
          'a4-highgpu-8g': ('B200', 8, 'A4: 8x B200, 224 vCPU, 3,968 GB RAM, 12 TB local SSD, 3.2 Tbps RDMA'),
          'a4x-highgpu-4g': ('GB200', 4, 'A4X: 4x GB200 (NVL72 rack slice), 140 vCPU, 884 GB RAM'),
          'a4x-maxgpu-4g': ('GB300', 4, 'A4X Max: 4x GB300'),
          'a2-ultragpu-8g': ('A100 80GB', 8, 'A2 Ultra: 8x A100 80GB, 96 vCPU, 1,360 GB RAM, 3 TB local SSD'),
          'a2-highgpu-8g': ('A100 40GB', 8, 'A2 Standard: 8x A100 40GB, 96 vCPU, 680 GB RAM'),
          'g2-standard-48': ('L4', 4, 'G2: 4x L4')}
GCP_COLS = [(r'DWS Flex-start', ('reserved', 'DWS Flex-start (dynamic)')), (r'DWS Calendar', ('reserved', 'DWS Calendar mode (dynamic)')),
            (r'Spot', ('spot', 'spot')), (r'(?:1 year|1 Year|CUDs - 1 Year)', ('reserved', '1-year CUD')), (r'(?:3 year|3 Year|CUDs - 3 Year)', ('reserved', '3-year CUD')),
            (r'On-[Dd]emand', ('on-demand', 'PAYG'))]
def gcp_vm(t):
    """cloud.google.com/compute/vm-instance-pricing static tables (2025+). Default region in static HTML = Iowa (us-central1)."""
    out = []; seen = set()
    for mt, (chip, n, specs) in GCP_MT.items():
        for m in re.finditer(re.escape(mt) + r' ', t):
            hs = t.rfind('Machine type', 0, m.start())
            hdr = t[hs:hs + 400]
            hdr = hdr[:hdr.find('$')] if '$' in hdr else hdr
            # header columns in order of appearance
            cols = []
            for mm in re.finditer(r'On-[Dd]emand|Default \*|DWS Flex-start|DWS Calendar|(?:Current )?Spot|Flexible CUD - [13] Year|(?:Resource )?CUDs - [13] Year|[13] year commitment', hdr):
                lab = mm.group(0)
                if lab.startswith(('On-', 'Default')): cols.append(('on-demand', 'PAYG'))
                elif 'Flex-start' in lab: cols.append(('reserved', 'DWS Flex-start (dynamic)'))
                elif 'Calendar' in lab: cols.append(('reserved', 'DWS Calendar mode (dynamic)'))
                elif 'Spot' in lab: cols.append(('spot', 'spot'))
                elif 'Flexible' in lab: cols.append(('reserved', ('1' if '1 Year' in lab else '3') + '-year flexible CUD'))
                else: cols.append(('reserved', ('1' if ('1 Year' in lab or '1 year' in lab) else '3') + '-year CUD'))
            seq = re.findall(r'(\$[\d.,]+ / 1 hour|N/A|\*\*)', t[m.end():m.end() + 420].split(' a')[0] if False else t[m.end():m.end() + 420])
            # stop at next machine type token
            nxt = re.search(r' [a-z]\d[a-z]?-[a-z]+-\d+g? ', t[m.end():m.end() + 420])
            seg = t[m.end():m.end() + (nxt.start() if nxt else 420)]
            seq = re.findall(r'(\$[\d.,]+ / 1 hour|N/A|\*\*)', seg)
            if not seq or not cols: continue
            for (pt, term), v in zip(cols, seq):
                if not v.startswith('$'): continue
                key = (mt, pt, term)
                if key in seen: continue
                seen.add(key)
                od = [money(o['quoted_price']) for o in out if o['config'].startswith(mt + ' ') and o['term'] == 'PAYG']
                if 'CUD' in term and 'flexible' not in term and od and money(v.split(' /')[0]) > 0.95 * od[0]:
                    term = term.replace('CUD', 'flexible CUD (inferred: ~0% discount, header unlabeled)')
                out.append(rec(chip, f'{mt} ({n}x {chip})', pt, term, v.split(' /')[0], f'per {n}-GPU VM-hour', n, mt != 'a3-megagpu-8g' and 'flexible' not in term,
                               f'GCP list, us-central1 (Iowa, default static table); {specs}' + ('; DWS prices are variable' if 'DWS' in term else '') + ('; spot is dynamic' if pt == 'spot' else '')))
    return out

def gcp_gpu_frame(html):
    """gpus-pricing iframe (2020-2023): <td cloud-pricer io-hourly='$2.934 per GPU'> cells; GPU component price only (vCPU/RAM extra)."""
    from bs4 import BeautifulSoup
    out = []; soup = BeautifulSoup(html, 'lxml')
    for tb in soup.find_all('table'):
        hdr = [th.get_text(' ', strip=True) for th in tb.find_all('th')]
        pcols = [h for h in hdr if 'price' in h.lower() or 'commitment' in h.lower()]
        for tr in tb.find_all('tr'):
            tds = tr.find_all('td')
            if not tds: continue
            name = tds[0].get_text(' ', strip=True)
            if 'A100' not in name and 'H100' not in name and not re.search(r'\bL4\b', name): continue
            chip = 'A100 80GB' if '80GB' in name.replace(' ', '') else ('A100 40GB' if 'A100' in name else norm_chip(name))
            if chip == 'A100 40GB' and '40GB' not in name.replace(' ', ''): chip = 'A100 40GB'  # 2020-22: only the 40GB A100 existed on GCP
            vals = [td.get('io-hourly') for td in tds if td.has_attr('cloud-pricer')]
            for h, v in zip(pcols, vals):
                if not v or not v.startswith('$'): continue
                hl = h.lower()
                pt, term = ('spot', 'preemptible/spot') if ('spot' in hl or 'preemptible' in hl) else (('reserved', '1-year CUD') if '1 year' in hl else (('reserved', '3-year CUD') if '3 year' in hl else ('on-demand', 'PAYG (GPU component only)')))
                out.append(rec(chip, f'GCP {name} GPU (attached to A2 VM)', pt, term, v.split(' per')[0], 'per GPU-hour', 1, True,
                               'GCP GPU component price, us-central1 (Iowa); A2 vCPU/RAM billed separately (not bundled)'))
    return out

# ---------------- RunPod per-model pages (www.runpod.io/gpu-models/<slug>, 2025-26) ----------------
RUNPOD_SLUG = {'h100-sxm': 'H100', 'h100-pcie': 'H100 PCIe', 'h100-nvl': 'H100 NVL', 'h200': 'H200', 'h200-sxm': 'H200', 'b200': 'B200',
               'a100-sxm': 'A100 80GB', 'a100-pcie': 'A100 80GB PCIe', 'l40s': 'L40S'}
def runpod_model(t, url):
    out = []
    slug = url.rstrip('/').split('/')[-1].split('?')[0]
    chip = RUNPOD_SLUG.get(slug)
    if not chip: return out
    m = re.search(r'Community Cloud \$ ?([\d.]*) /hr(?: N/A)? Secure Cloud \$ ?([\d.]+) /hr', t)
    if not m: return out
    com, sec = m.groups()
    cfg = f'RunPod {slug} model page'
    if sec and float(sec) > 0:
        out.append(rec(chip, cfg + ' (Secure Cloud)', 'on-demand', 'PAYG', '$' + sec, 'per GPU-hour', 1, chip != 'A100 80GB PCIe', 'Secure Cloud; from per-GPU-model page (www.runpod.io/gpu-models/' + slug + ')'))
    if com and float(com) > 0:
        out.append(rec(chip, cfg + ' (Community Cloud)', 'on-demand', 'PAYG (community cloud)', '$' + com, 'per GPU-hour', 1, False, 'Community Cloud (third-party hosts); per-GPU-model page'))
    return out

# ---------------- Vast.ai per-GPU pricing pages (marketplace) ----------------
VAST_SLUG = {'H100-SXM': 'H100', 'H100-PCIE': 'H100 PCIe', 'H200': 'H200', 'B200': 'B200', 'A100-SXM4': 'A100 80GB', 'A100-PCIE': 'A100 80GB PCIe', 'L40S': 'L40S'}
def vast_gpu(t, url):
    out = []
    slug = url.rstrip('/').split('/')[-1].split('?')[0]
    chip = VAST_SLUG.get(slug)
    if not chip: return out
    m = re.search(r'prices ranging from \$([\d.]+)/hr to \$([\d.]+)/hr', t)
    if m:
        lo, hi = m.groups()
        out.append(rec(chip, f'Vast.ai {slug} marketplace', 'marketplace', 'marketplace low (on-demand offers)', '$' + lo, 'per GPU-hour', 1, True, f'Vast.ai marketplace: range ${lo}-${hi}/hr across live host offers at capture time (not a list price)'))
        out.append(rec(chip, f'Vast.ai {slug} marketplace', 'marketplace', 'marketplace high (on-demand offers)', '$' + hi, 'per GPU-hour', 1, False, 'Vast.ai marketplace range high'))
        return out
    m = re.search(r'by the hour at a price of \$([\d.]+)/hr', t)
    if m:
        out.append(rec(chip, f'Vast.ai {slug} marketplace', 'marketplace', 'marketplace low (on-demand offers)', '$' + m.group(1), 'per GPU-hour', 1, True, 'Vast.ai marketplace: single offer price at capture time (not a list price)'))
    return out

VANTAGE_INST.update({
    'p4d.24xlarge': ('A100 40GB', 8, 'AWS', '96 vCPU, 1,152 GiB RAM, 400 Gbps EFA, 8x1 TB NVMe'),
    'p4de.24xlarge': ('A100 80GB', 8, 'AWS', '96 vCPU, 1,152 GiB RAM, 400 Gbps EFA, 8x1 TB NVMe'),
    'p5e.48xlarge': ('H200', 8, 'AWS', '192 vCPU, 2 TiB RAM, 3,200 Gbps EFA, 8x3.84 TB NVMe (p5e)'),
    'p5en.48xlarge': ('H200', 8, 'AWS', '192 vCPU, 2 TiB RAM, 3,200 Gbps EFAv3, 8x3.84 TB NVMe (p5en)'),
    'p6-b200.48xlarge': ('B200', 8, 'AWS', '192 vCPU, 2 TiB RAM, 3,200 Gbps EFAv4, 8x3.84 TB NVMe'),
    'nd96isrh100-v5': ('H100', 8, 'Azure', '96 vCPU, 1,900 GiB RAM, 8x400 Gb/s InfiniBand (ND H100 v5)'),
    'nd96isrh200-v5': ('H200', 8, 'Azure', '96 vCPU, 1,850 GiB RAM, 8x400 Gb/s InfiniBand (ND H200 v5)'),
    'nd96amsr': ('A100 80GB', 8, 'Azure', 'ND96amsr A100 v4: 96 vCPU, 1,900 GiB RAM, 8x200 Gb/s HDR InfiniBand'),
    'nd96asr': ('A100 40GB', 8, 'Azure', 'ND96asr v4: 96 vCPU, 900 GiB RAM, 8x200 Gb/s HDR InfiniBand'),
    'nc40adsh100-v5': ('H100 NVL', 1, 'Azure', 'NC40ads H100 v5: 1x H100 NVL 94GB, 40 vCPU, 320 GiB RAM'),
    'nc24ads-v4': ('A100 80GB PCIe', 1, 'Azure', 'NC24ads A100 v4: 1x A100 80GB PCIe, 24 vCPU, 220 GiB RAM'),
})

# ---------------- AWS EC2 P4 instance-type page (aws.amazon.com/ec2/instance-types/p4/) ----------------
def aws_p4(t):
    """Row: p4d.24xlarge 96 1152 8 400 Gbps ... 19 Gbps $32.77 $19.22 $11.57 = On-Demand / 1-yr RI effective hourly / 3-yr RI effective hourly (us-east-1 Linux)."""
    out = []
    for m in re.finditer(r'(p4de?\.24xlarge) (\d+) (\d+) 8 400[^$]{0,160}?(\$[\d.]+|N/A) (\$[\d.]+|N/A) (\$[\d.]+|N/A)', t):
        inst = m.group(1); chip = 'A100 80GB' if inst == 'p4de.24xlarge' else 'A100 40GB'
        specs = f'{inst}: 8x {chip}, {m.group(2)} vCPU, {m.group(3)} GiB RAM, 400 Gbps EFA, 8x1 TB NVMe'
        for (pt, term), v in zip([('on-demand', 'PAYG'), ('reserved', '1-year RI (effective hourly)'), ('reserved', '3-year RI (effective hourly)')], m.groups()[3:]):
            if money(v) is None: continue
            out.append(rec(chip, f'{inst} (8x {chip})', pt, term, v, 'per 8-GPU instance-hour', 8, True, 'AWS P4 instance-type page; prices for Linux/Unix, US East (N. Virginia); ' + specs))
    return out

CONTACT_RE = re.compile(r'(GB300 NVL72|GB200 NVL72|GB300|GB200|B300|B200|H200|H100|Vera Rubin|VR200|Rubin)(?:[^$]{0,60}?)(Contact (?:sales|Sales|us|Us)|CONTACT SALES|Request access|Reserve now|Talk to sales)')
def contact_rows(t, priced_chips):
    """Explicit 'no public price' rows: chip name followed (within 60 chars, no $ in between) by Contact sales/us etc., for chips with no $ price in this snapshot."""
    out = []; seen = set()
    for m in CONTACT_RE.finditer(t):
        chip = norm_chip(m.group(1)) or ('VR200 NVL144' if 'Rubin' in m.group(1) or 'VR' in m.group(1) else None)
        if not chip or chip in priced_chips or chip in seen or 'Search' in m.group(0) or 'Log in' in m.group(0): continue
        seen.add(chip)
        out.append(rec(chip, f'{m.group(1)} (no public price)', 'on-demand', 'contact sales (no public price)', m.group(2), 'n/a', 1, False,
                       f'NO PUBLIC PRICE: page text "{m.group(0)[:90]}"'))
    return out
