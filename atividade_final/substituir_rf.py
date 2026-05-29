"""
Remove os slides de RF do apresentacao_compilada.pptx e substitui
pelos slides de RF_modelo.pptx.

Identifica os slides de RF automaticamente: são todos a partir do
primeiro slide cujo texto começa com "Random Forest".
"""
import zipfile, re, os
from lxml import etree

BASE     = '/home/matheus-sales/Documents/atividade_reconhecimento/relatorio'
COMPILED = f'{BASE}/apresentacao_compilada.pptx'
RF_MODEL = f'{BASE}/RF_modelo.pptx'
PEDRO    = f'{BASE}/Classificação de áreas queimadas-REC PADRÕES_pedro.pptx'

NS_CT   = 'http://schemas.openxmlformats.org/package/2006/content-types'
NS_PRES = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_R    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

RT_SLIDE  = f'{NS_R}/slide'
RT_LAYOUT = f'{NS_R}/slideLayout'

SKIP_RELTYPES = {
    RT_LAYOUT,
    f'{NS_R}/slideMaster', f'{NS_R}/theme', f'{NS_R}/tableStyles',
    f'{NS_R}/viewProps',   f'{NS_R}/presProps', f'{NS_R}/notesSlide',
    f'{NS_R}/notesMaster', f'{NS_R}/handoutMaster', f'{NS_R}/slideLayout',
    f'{NS_R}/comments',    f'{NS_R}/commentAuthors',
}
MEDIA_RELTYPES = {
    f'{NS_R}/image', f'{NS_R}/video', f'{NS_R}/audio', f'{NS_R}/media',
}


def parse(data):
    return etree.fromstring(data)

def serialize(root):
    return etree.tostring(root, xml_declaration=True,
                          encoding='UTF-8', standalone=True)

def read_rels(z, path):
    dir_, name = path.rsplit('/', 1) if '/' in path else ('', path)
    rp = f"{dir_}/_rels/{name}.rels" if dir_ else f"_rels/{name}.rels"
    if rp not in z.namelist():
        return {}
    root = parse(z.read(rp))
    return {r.get('Id'): (r.get('Type'), r.get('Target'),
                          r.get('TargetMode') == 'External')
            for r in root}

def build_ct_map(z):
    root = parse(z.read('[Content_Types].xml'))
    ct = {}
    for c in root:
        if c.tag == f'{{{NS_CT}}}Default':
            ct['.' + c.get('Extension')] = c.get('ContentType')
        elif c.tag == f'{{{NS_CT}}}Override':
            ct[c.get('PartName')] = c.get('ContentType')
    return ct

def resolve_target(base_path, target):
    if target.startswith('/'):
        return target.lstrip('/')
    dir_ = base_path.rsplit('/', 1)[0] if '/' in base_path else ''
    parts = (dir_ + '/' + target).split('/')
    result = []
    for p in parts:
        if p == '..':
            if result: result.pop()
        elif p and p != '.':
            result.append(p)
    return '/'.join(result)

def slide_text(z, slide_path):
    """Extrai texto plano do slide (para identificação)."""
    try:
        root = parse(z.read(slide_path))
        texts = root.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}t')
        return ' '.join(t.text for t in texts if t.text)
    except Exception:
        return ''

def build_pedro_layout_map():
    mapping = {}
    with zipfile.ZipFile(PEDRO) as z:
        layouts = sorted(
            [n for n in z.namelist()
             if re.match(r'ppt/slideLayouts/slideLayout\d+\.xml$', n)],
            key=lambda x: int(re.search(r'\d+', x).group())
        )
        for lpath in layouts:
            root = parse(z.read(lpath))
            csld = root.find(f'{{{NS_PRES}}}cSld')
            name = csld.get('name', '') if csld is not None else ''
            if name:
                mapping[name] = '../slideLayouts/' + lpath.split('/')[-1]
    mapping.setdefault('BLANK', mapping.get('TITLE_AND_BODY',
                                             '../slideLayouts/slideLayout3.xml'))
    return mapping

def get_layout_name(src_z, slide_path, layout_rel_target):
    abs_path = resolve_target(slide_path, layout_rel_target)
    if abs_path not in src_z.namelist():
        return 'TITLE_AND_BODY'
    root = parse(src_z.read(abs_path))
    csld = root.find(f'{{{NS_PRES}}}cSld')
    return (csld.get('name', 'TITLE_AND_BODY') if csld is not None
            else 'TITLE_AND_BODY')

def ordered_slides_from_zip(z):
    prs_xml  = parse(z.read('ppt/presentation.xml'))
    prs_rels = read_rels(z, 'ppt/presentation.xml')
    sld_id_lst = prs_xml.find(f'.//{{{NS_PRES}}}sldIdLst')
    result = []
    if sld_id_lst is not None:
        for elem in sld_id_lst:
            rid = elem.get(f'{{{NS_R}}}id')
            if rid in prs_rels:
                _, target, _ = prs_rels[rid]
                result.append(resolve_target('ppt/presentation.xml', target))
    return result


# ─────────────────────────────────────────────────────────────
# 1. Analisa o arquivo compilado — identifica slides de RF
# ─────────────────────────────────────────────────────────────
print('=== Analisando apresentacao_compilada.pptx ===')
with zipfile.ZipFile(COMPILED) as z:
    comp_ordered = ordered_slides_from_zip(z)
    print(f'Total de slides: {len(comp_ordered)}')

    rf_start_idx = None
    for i, sp in enumerate(comp_ordered):
        txt = slide_text(z, sp)
        marker = 'Random Forest' in txt and i > 0
        tag = ' ← RF START' if (marker and rf_start_idx is None) else ''
        print(f'  [{i+1:02d}] {sp.split("/")[-1]}  {txt[:60]}{tag}')
        if marker and rf_start_idx is None:
            rf_start_idx = i

if rf_start_idx is None:
    print('ERRO: nenhum slide de Random Forest encontrado.')
    exit(1)

keep_slides = comp_ordered[:rf_start_idx]
drop_slides = comp_ordered[rf_start_idx:]
print(f'\nManter : slides 1–{rf_start_idx}  ({len(keep_slides)} slides)')
print(f'Remover: slides {rf_start_idx+1}–{len(comp_ordered)}  ({len(drop_slides)} slides)')

# ─────────────────────────────────────────────────────────────
# 2. Carrega arquivo compilado em memória, remove slides de RF
# ─────────────────────────────────────────────────────────────
with zipfile.ZipFile(COMPILED) as z:
    out_files = {n: z.read(n) for n in z.namelist()}

out_prs_xml  = parse(out_files['ppt/presentation.xml'])
out_prs_rels = parse(out_files['ppt/_rels/presentation.xml.rels'])
out_ct_xml   = parse(out_files['[Content_Types].xml'])

sld_id_lst = out_prs_xml.find(f'.//{{{NS_PRES}}}sldIdLst')

# Constrói mapa rId → slide path (para identificar rIds a remover)
rid2target = {r.get('Id'): r.get('Target')
              for r in out_prs_rels
              if r.get('Type') == RT_SLIDE}
target2rid = {v: k for k, v in rid2target.items()}

# Remove entradas de RF do sldIdLst
drop_paths_set = set(drop_slides)
drop_rids = set()
for elem in list(sld_id_lst):
    rid  = elem.get(f'{{{NS_R}}}id')
    tgt  = rid2target.get(rid, '')
    full = 'ppt/' + tgt if not tgt.startswith('ppt') else tgt
    if full in drop_paths_set:
        sld_id_lst.remove(elem)
        drop_rids.add(rid)

# Remove rId entries de presentation.xml.rels
for rel in list(out_prs_rels):
    if rel.get('Id') in drop_rids:
        out_prs_rels.remove(rel)

# Remove arquivos de slide e .rels dos slides removidos
drop_files = set()
for sp in drop_slides:
    drop_files.add(sp)
    base = sp.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels'
    drop_files.add(base)

for f in drop_files:
    out_files.pop(f, None)

# Remove Override de Content_Types
drop_parts = {('/' + f) for f in drop_files}
for elem in list(out_ct_xml):
    if (elem.tag == f'{{{NS_CT}}}Override' and
            elem.get('PartName') in drop_parts):
        out_ct_xml.remove(elem)

print(f'\nRemovidos {len(drop_slides)} slides de RF do compilado.')

# ─────────────────────────────────────────────────────────────
# 3. Adiciona slides de RF_modelo.pptx
# ─────────────────────────────────────────────────────────────
print('\n=== Adicionando slides de RF_modelo.pptx ===')

pedro_layouts = build_pedro_layout_map()
print('Layouts do pedro:', list(pedro_layouts.keys()))

# Contador de slides e media existentes pós-remoção
existing_slides = sorted(
    [n for n in out_files if re.match(r'ppt/slides/slide\d+\.xml$', n)],
    key=lambda x: int(re.search(r'\d+', x.split('/')[-1]).group())
)
slide_count = max(
    (int(re.search(r'\d+', n.split('/')[-1]).group()) for n in existing_slides),
    default=0
)
media_count = len([n for n in out_files if n.startswith('ppt/media/')])

# rIds existentes no presentation.xml.rels
existing_rids = {r.get('Id') for r in out_prs_rels}
def next_prs_rid():
    i = 1
    while f'rId{i}' in existing_rids:
        i += 1
    existing_rids.add(f'rId{i}')
    return f'rId{i}'

# Id numérico máximo no sldIdLst
existing_ids = {int(s.get('id', 256)) for s in sld_id_lst}
next_slide_id = max(existing_ids) + 1 if existing_ids else 256

rf_src = zipfile.ZipFile(RF_MODEL, 'r')
rf_ct  = build_ct_map(rf_src)

# Ordem dos slides do RF_modelo
rf_prs_xml  = parse(rf_src.read('ppt/presentation.xml'))
rf_prs_rels = read_rels(rf_src, 'ppt/presentation.xml')
rf_sld_lst  = rf_prs_xml.find(f'.//{{{NS_PRES}}}sldIdLst')
rf_ordered  = []
if rf_sld_lst is not None:
    for elem in rf_sld_lst:
        rid = elem.get(f'{{{NS_R}}}id')
        if rid in rf_prs_rels:
            _, tgt, _ = rf_prs_rels[rid]
            rf_ordered.append(resolve_target('ppt/presentation.xml', tgt))

added = 0
for src_slide_path in rf_ordered:
    if src_slide_path not in rf_src.namelist():
        continue

    slide_count += 1
    new_slide_path = f'ppt/slides/slide{slide_count}.xml'
    new_rels_path  = f'ppt/slides/_rels/slide{slide_count}.xml.rels'

    slide_xml_bytes = rf_src.read(src_slide_path)
    slide_rels      = read_rels(rf_src, src_slide_path)

    new_slide_rels = {}
    rid_map        = {}
    rel_counter    = 1

    for old_rid, (reltype, target, is_ext) in slide_rels.items():
        if reltype == RT_LAYOUT:
            layout_name  = get_layout_name(rf_src, src_slide_path, target)
            pedro_target = pedro_layouts.get(
                layout_name,
                pedro_layouts.get('TITLE_AND_BODY',
                                  '../slideLayouts/slideLayout3.xml')
            )
            new_rid = f'rId{rel_counter}'; rel_counter += 1
            new_slide_rels[new_rid] = (reltype, pedro_target, False)
            rid_map[old_rid] = new_rid
            continue

        if reltype in SKIP_RELTYPES:
            continue

        if is_ext:
            new_rid = f'rId{rel_counter}'; rel_counter += 1
            new_slide_rels[new_rid] = (reltype, target, True)
            rid_map[old_rid] = new_rid
            continue

        if reltype not in MEDIA_RELTYPES:
            continue

        abs_target = resolve_target(src_slide_path, target)
        if abs_target not in rf_src.namelist():
            continue

        media_count += 1
        ext = os.path.splitext(abs_target)[1]
        new_media_name = f'ppt/media/image{media_count}{ext}'
        out_files[new_media_name] = rf_src.read(abs_target)

        ct = rf_ct.get(abs_target) or rf_ct.get(ext, 'image/png')
        etree.SubElement(out_ct_xml, f'{{{NS_CT}}}Override',
                         PartName=f'/{new_media_name}', ContentType=ct)

        new_rid = f'rId{rel_counter}'; rel_counter += 1
        new_slide_rels[new_rid] = (reltype,
                                   f'../media/image{media_count}{ext}',
                                   False)
        rid_map[old_rid] = new_rid

    # Substitui rIds no XML
    slide_xml_str = slide_xml_bytes.decode('utf-8')
    for old, new in sorted(rid_map.items(), key=lambda x: -len(x[0])):
        slide_xml_str = slide_xml_str.replace(f'"{old}"', f'"{new}"')
    out_files[new_slide_path] = slide_xml_str.encode('utf-8')

    # .rels do novo slide
    rels_root = etree.Element(
        'Relationships',
        xmlns='http://schemas.openxmlformats.org/package/2006/relationships'
    )
    for rid, (reltype, tgt, is_ext) in new_slide_rels.items():
        rel_elem = etree.SubElement(rels_root, 'Relationship',
                                    Id=rid, Type=reltype, Target=tgt)
        if is_ext:
            rel_elem.set('TargetMode', 'External')
    out_files[new_rels_path] = serialize(rels_root)

    # Registra no presentation.xml.rels
    new_prs_rid = next_prs_rid()
    etree.SubElement(out_prs_rels, 'Relationship',
                     Id=new_prs_rid, Type=RT_SLIDE,
                     Target=f'slides/slide{slide_count}.xml')

    # Registra no sldIdLst
    elem = etree.SubElement(sld_id_lst, f'{{{NS_PRES}}}sldId')
    elem.set('id', str(next_slide_id))
    elem.set(f'{{{NS_R}}}id', new_prs_rid)
    next_slide_id += 1

    # Content-Type do slide
    ct_part = f'/{new_slide_path}'
    if not any(c.get('PartName') == ct_part for c in out_ct_xml
               if c.tag == f'{{{NS_CT}}}Override'):
        etree.SubElement(out_ct_xml, f'{{{NS_CT}}}Override',
                         PartName=ct_part,
                         ContentType=(
                             'application/vnd.openxmlformats-officedocument'
                             '.presentationml.slide+xml'))
    added += 1

rf_src.close()
print(f'Adicionados {added} slides do RF_modelo.pptx')

# ─────────────────────────────────────────────────────────────
# 4. Serializa e escreve ZIP final
# ─────────────────────────────────────────────────────────────
out_files['ppt/presentation.xml']            = serialize(out_prs_xml)
out_files['ppt/_rels/presentation.xml.rels'] = serialize(out_prs_rels)
out_files['[Content_Types].xml']             = serialize(out_ct_xml)

tmp = COMPILED + '.tmp'
with zipfile.ZipFile(tmp, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for name, data in out_files.items():
        zout.writestr(name, data)
os.replace(tmp, COMPILED)

# ─────────────────────────────────────────────────────────────
# 5. Validação final
# ─────────────────────────────────────────────────────────────
print(f'\n✓ Salvo: {COMPILED}')
with zipfile.ZipFile(COMPILED) as zf:
    names = zf.namelist()
    slides = sorted([n for n in names if re.match(r'ppt/slides/slide\d+\.xml$', n)],
                    key=lambda x: int(re.search(r'\d+', x.split('/')[-1]).group()))
    prs_check = parse(zf.read('ppt/presentation.xml'))
    sll = prs_check.find(f'.//{{{NS_PRES}}}sldIdLst')
    registered = len(list(sll)) if sll is not None else 0
    dupes = len(names) - len(set(names))
    media = len([n for n in names if n.startswith('ppt/media/')])

    missing = 0
    for s in slides:
        rp = s.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels'
        if rp in names:
            for r in parse(zf.read(rp)):
                if (r.get('Type') in MEDIA_RELTYPES and
                        r.get('TargetMode') != 'External'):
                    mn = re.sub(r'\.\./media/', 'ppt/media/', r.get('Target', ''))
                    if mn not in names:
                        missing += 1

print(f'  Total slides (ZIP)     : {len(slides)}')
print(f'  Total slides (sldIdLst): {registered}')
print(f'  Total media            : {media}')
print(f'  Duplicatas             : {dupes}')
print(f'  Media faltando         : {missing}')
