"""
Compila 3 PPTX em um único arquivo manipulando diretamente o ZIP.
Single-pass: lê todos os fontes de uma vez, escreve o ZIP final uma única vez.
Ordem: pedro (base) → REC_PADRÕES → RF_modelo
"""
import zipfile, re, shutil, os
from lxml import etree

BASE   = '/home/matheus-sales/Documents/atividade_reconhecimento/relatorio'
PEDRO  = f'{BASE}/Classificação de áreas queimadas-REC PADRÕES_pedro.pptx'
SOURCES = [
    f'{BASE}/Classificação de áreas queimadas - REC PADRÕES.pptx',
    f'{BASE}/RF_modelo.pptx',
]
OUTPUT = f'{BASE}/apresentacao_compilada.pptx'

NS_PKG  = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS_CT   = 'http://schemas.openxmlformats.org/package/2006/content-types'
NS_PRES = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_R    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

RT_SLIDE  = f'{NS_R}/slide'
RT_LAYOUT = f'{NS_R}/slideLayout'

SKIP_RELTYPES = {
    RT_LAYOUT,
    f'{NS_R}/slideMaster',
    f'{NS_R}/theme',
    f'{NS_R}/tableStyles',
    f'{NS_R}/viewProps',
    f'{NS_R}/presProps',
    f'{NS_R}/notesSlide',
    f'{NS_R}/notesMaster',
    f'{NS_R}/handoutMaster',
    f'{NS_R}/slideLayout',
    f'{NS_R}/comments',
    f'{NS_R}/commentAuthors',
}

MEDIA_RELTYPES = {
    f'{NS_R}/image',
    f'{NS_R}/video',
    f'{NS_R}/audio',
    f'{NS_R}/media',
}


def parse(data):
    return etree.fromstring(data)


def serialize(root):
    return etree.tostring(root, xml_declaration=True,
                          encoding='UTF-8', standalone=True)


def read_rels(z, path):
    dir_, name = path.rsplit('/', 1) if '/' in path else ('', path)
    rels_path = f"{dir_}/_rels/{name}.rels" if dir_ else f"_rels/{name}.rels"
    if rels_path not in z.namelist():
        return {}
    root = parse(z.read(rels_path))
    return {rel.get('Id'): (rel.get('Type'), rel.get('Target'),
                             rel.get('TargetMode') == 'External')
            for rel in root}


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


def build_pedro_layout_map(z):
    """Retorna dict: layout_name → '../slideLayouts/slideLayoutN.xml'."""
    mapping = {}
    layouts = sorted(
        [n for n in z.namelist() if re.match(r'ppt/slideLayouts/slideLayout\d+\.xml$', n)],
        key=lambda x: int(re.search(r'\d+', x).group())
    )
    for lpath in layouts:
        root = parse(z.read(lpath))
        csld = root.find(f'{{{NS_PRES}}}cSld')
        name = csld.get('name', '') if csld is not None else ''
        rel_path = '../slideLayouts/' + lpath.split('/')[-1]
        if name:
            mapping[name] = rel_path
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


def ordered_slides(src_z):
    """Retorna lista de slide paths na ordem do sldIdLst."""
    src_prs_xml  = parse(src_z.read('ppt/presentation.xml'))
    src_prs_rels = read_rels(src_z, 'ppt/presentation.xml')
    sld_id_lst   = src_prs_xml.find(f'.//{{{NS_PRES}}}sldIdLst')

    result = []
    if sld_id_lst is not None:
        for sld_id in sld_id_lst:
            rid = sld_id.get(f'{{{NS_R}}}id')
            if rid in src_prs_rels:
                _, target, _ = src_prs_rels[rid]
                result.append(resolve_target('ppt/presentation.xml', target))

    if not result:
        result = sorted(
            [n for n in src_z.namelist()
             if re.match(r'ppt/slides/slide\d+\.xml$', n)],
            key=lambda x: int(re.search(r'\d+', x.split('/')[-1]).group())
        )
    return result


# ─────────────────────────────────────────────────────────────────────────────
# Passo 1: carrega pedro como base (todos os arquivos em memória)
# ─────────────────────────────────────────────────────────────────────────────
print("Carregando base (pedro)...")
with zipfile.ZipFile(PEDRO) as z_pedro:
    pedro_names  = z_pedro.namelist()
    out_files    = {n: z_pedro.read(n) for n in pedro_names}
    pedro_layouts = build_pedro_layout_map(z_pedro)

print("Layouts do pedro:", list(pedro_layouts.keys()))

pedro_slides = len([n for n in pedro_names
                    if re.match(r'ppt/slides/slide\d+\.xml$', n)])
pedro_media  = len([n for n in pedro_names if n.startswith('ppt/media/')])
print(f"  {pedro_slides} slides, {pedro_media} media")

# Contadores globais (continuam a partir do pedro)
slide_count = pedro_slides
media_count = pedro_media

# ─────────────────────────────────────────────────────────────────────────────
# Passo 2: parse dos XMLs que vamos modificar (feito UMA VEZ sobre o pedro)
# ─────────────────────────────────────────────────────────────────────────────
out_prs_xml      = parse(out_files['ppt/presentation.xml'])
out_prs_rels_xml = parse(out_files['ppt/_rels/presentation.xml.rels'])
out_ct_xml       = parse(out_files['[Content_Types].xml'])

# Conjunto de rIds já existentes no presentation.xml.rels
existing_rids = {rel.get('Id') for rel in out_prs_rels_xml}

def next_prs_rid():
    i = 1
    while f'rId{i}' in existing_rids:
        i += 1
    existing_rids.add(f'rId{i}')
    return f'rId{i}'

sld_id_lst_out = out_prs_xml.find(f'.//{{{NS_PRES}}}sldIdLst')
existing_ids   = {int(s.get('id', 256)) for s in sld_id_lst_out}
next_slide_id  = max(existing_ids) + 1 if existing_ids else 256

# ─────────────────────────────────────────────────────────────────────────────
# Passo 3: processa cada arquivo fonte (single pass, tudo em memória)
# ─────────────────────────────────────────────────────────────────────────────
for src_path in SOURCES:
    fname  = src_path.split('/')[-1]
    src_z  = zipfile.ZipFile(src_path, 'r')
    src_ct = build_ct_map(src_z)

    slides_ordered = ordered_slides(src_z)
    added = 0

    for src_slide_path in slides_ordered:
        if src_slide_path not in src_z.namelist():
            continue

        slide_count += 1
        new_slide_path = f'ppt/slides/slide{slide_count}.xml'
        new_rels_path  = f'ppt/slides/_rels/slide{slide_count}.xml.rels'

        slide_xml_bytes = src_z.read(src_slide_path)
        slide_rels      = read_rels(src_z, src_slide_path)

        new_slide_rels = {}
        rid_map        = {}
        rel_counter    = 1

        for old_rid, (reltype, target, is_ext) in slide_rels.items():

            # Layout: mapeia para o equivalente do pedro por nome
            if reltype == RT_LAYOUT:
                layout_name  = get_layout_name(src_z, src_slide_path, target)
                pedro_target = pedro_layouts.get(
                    layout_name,
                    pedro_layouts.get('TITLE_AND_BODY',
                                      '../slideLayouts/slideLayout3.xml')
                )
                new_rid = f'rId{rel_counter}'; rel_counter += 1
                new_slide_rels[new_rid] = (reltype, pedro_target, False)
                rid_map[old_rid] = new_rid
                continue

            # Tipos não-media: pular
            if reltype in SKIP_RELTYPES:
                continue

            # Hyperlink externo
            if is_ext:
                new_rid = f'rId{rel_counter}'; rel_counter += 1
                new_slide_rels[new_rid] = (reltype, target, True)
                rid_map[old_rid] = new_rid
                continue

            # Media interno: copia com nome único
            if reltype not in MEDIA_RELTYPES:
                continue

            abs_target = resolve_target(src_slide_path, target)
            if abs_target not in src_z.namelist():
                continue

            media_count += 1
            ext = os.path.splitext(abs_target)[1]
            new_media_name = f'ppt/media/image{media_count}{ext}'
            out_files[new_media_name] = src_z.read(abs_target)

            ct = (src_ct.get(abs_target) or src_ct.get(ext, 'image/png'))
            etree.SubElement(out_ct_xml, f'{{{NS_CT}}}Override',
                             PartName=f'/{new_media_name}',
                             ContentType=ct)

            new_rid = f'rId{rel_counter}'; rel_counter += 1
            new_slide_rels[new_rid] = (reltype,
                                       f'../media/image{media_count}{ext}',
                                       False)
            rid_map[old_rid] = new_rid

        # Substitui rIds no XML do slide
        slide_xml_str = slide_xml_bytes.decode('utf-8')
        for old, new in sorted(rid_map.items(), key=lambda x: -len(x[0])):
            slide_xml_str = slide_xml_str.replace(f'"{old}"', f'"{new}"')
        out_files[new_slide_path] = slide_xml_str.encode('utf-8')

        # Cria .rels do novo slide
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

        # Registra slide no presentation.xml.rels
        new_prs_rid = next_prs_rid()
        etree.SubElement(
            out_prs_rels_xml, 'Relationship',
            Id=new_prs_rid, Type=RT_SLIDE,
            Target=f'slides/slide{slide_count}.xml'
        )

        # Registra no sldIdLst
        elem = etree.SubElement(sld_id_lst_out, f'{{{NS_PRES}}}sldId')
        elem.set('id', str(next_slide_id))
        elem.set(f'{{{NS_R}}}id', new_prs_rid)
        next_slide_id += 1

        # Registra content type do slide
        ct_part = f'/{new_slide_path}'
        if not any(c.get('PartName') == ct_part for c in out_ct_xml
                   if c.tag == f'{{{NS_CT}}}Override'):
            etree.SubElement(out_ct_xml, f'{{{NS_CT}}}Override',
                             PartName=ct_part,
                             ContentType=(
                                 'application/vnd.openxmlformats-officedocument'
                                 '.presentationml.slide+xml'))

        added += 1

    src_z.close()
    print(f"  +{added:2d} slides  ← {fname}")

# Garante tipos Default de imagem comuns
for ext2, ct2 in [('png', 'image/png'), ('jpg', 'image/jpeg'),
                  ('jpeg', 'image/jpeg'), ('gif', 'image/gif')]:
    if not any(c.get('Extension') == ext2 for c in out_ct_xml
               if c.tag == f'{{{NS_CT}}}Default'):
        etree.SubElement(out_ct_xml, f'{{{NS_CT}}}Default',
                         Extension=ext2, ContentType=ct2)

# ─────────────────────────────────────────────────────────────────────────────
# Passo 4: serializa XMLs modificados e escreve ZIP final (uma única vez)
# ─────────────────────────────────────────────────────────────────────────────
out_files['ppt/presentation.xml']            = serialize(out_prs_xml)
out_files['ppt/_rels/presentation.xml.rels'] = serialize(out_prs_rels_xml)
out_files['[Content_Types].xml']             = serialize(out_ct_xml)

tmp = OUTPUT + '.tmp'
with zipfile.ZipFile(tmp, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for name, data in out_files.items():
        zout.writestr(name, data)
os.replace(tmp, OUTPUT)

# ─────────────────────────────────────────────────────────────────────────────
# Validação final
# ─────────────────────────────────────────────────────────────────────────────
print(f"\n✓ Salvo: {OUTPUT}")
RT_LAYOUT_CHK = f'{NS_R}/slideLayout'
with zipfile.ZipFile(OUTPUT) as zfinal:
    names  = zfinal.namelist()
    slides = sorted([n for n in names if re.match(r'ppt/slides/slide\d+\.xml$', n)],
                    key=lambda x: int(re.search(r'\d+', x.split('/')[-1]).group()))
    dupes  = len(names) - len(set(names))
    media  = len([n for n in names if n.startswith('ppt/media/')])
    missing_media = 0
    layout_usage  = {}
    for s in slides:
        rp = s.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels'
        if rp in names:
            root = etree.fromstring(zfinal.read(rp))
            for rel in root:
                if rel.get('Type') == RT_LAYOUT_CHK:
                    tgt = rel.get('Target', '')
                    layout_usage[tgt] = layout_usage.get(tgt, 0) + 1
                if (rel.get('Type') in MEDIA_RELTYPES and
                        rel.get('TargetMode') != 'External'):
                    mn = re.sub(r'\.\./media/', 'ppt/media/', rel.get('Target', ''))
                    if mn not in names:
                        missing_media += 1

    # Verifica sldIdLst
    prs = etree.fromstring(zfinal.read('ppt/presentation.xml'))
    sld_id_lst_check = prs.find(f'.//{{{NS_PRES}}}sldIdLst')
    registered = len(list(sld_id_lst_check)) if sld_id_lst_check is not None else 0

print(f"  Total slides (ZIP)     : {len(slides)}")
print(f"  Total slides (sldIdLst): {registered}")
print(f"  Total media            : {media}")
print(f"  Duplicatas             : {dupes}")
print(f"  Media faltando         : {missing_media}")
print(f"  Uso de layouts:")
for lyt, cnt in sorted(layout_usage.items()):
    print(f"    {lyt} → {cnt} slides")
