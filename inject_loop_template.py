"""
inject_loop_template.py
========================
Memperbaiki dan menyuntikkan loop docxtpl ke template pkkm.docx.

Kasus yang ditangani:
  KASUS A (pkkm.docx root) - Template dengan variabel FLAT:
    - Collapse split-run {{no}}, {{waktu}}, {{kegiatan}} yang terpecah
    - Ganti variabel flat -> item.xxx
    - Ganti hardcoded '1 JP' -> {{item.jp}}
    - Inject {%tr for item in item_materi %} dan {%tr endfor %}

  KASUS B (templates/pkkm.docx) - Template yang sudah pakai {{ item.xxx }}:
    - Fix {% endfor %} yang terpecah menjadi: {% end + for + %}
    - Tidak perlu inject loop baru

Jalankan: python inject_loop_template.py
"""

import zipfile
import re
import os
import sys
import shutil
from pathlib import Path

# Fix encoding untuk Windows terminal
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')


BASE_DIR = Path(__file__).parent.resolve()
TEMPLATE_PATH = BASE_DIR / "pkkm.docx"
BACKUP_PATH = BASE_DIR / "pkkm_sebelum_inject.docx"


# ─── STEP 1: Collapse split-run  ────────────────────────────────────────────
def collapse_split_vars(xml: str) -> str:
    """
    Gabungkan split-run: 
      <w:r>...<w:t>{{</w:t></w:r>[proofErr?]<w:r>...<w:t>VARNAME</w:t></w:r>[proofErr?]<w:r>...<w:t>}}</w:t></w:r>
    menjadi:
      <w:r>...<w:t>{{VARNAME}}</w:t></w:r>
    
    Strategi: hapus semua XML tags di antara {{ dan }} yang merupakan varname,
    lalu pastikan hanya ada satu <w:r> per variabel.
    """
    def collapse_match(m):
        full = m.group(0)
        varname = m.group('var')
        # Ambil run pertama (run yang berisi '{{')
        run1_end = full.find('</w:r>') + len('</w:r>')
        run1 = full[:run1_end]
        # Modifikasi <w:t>{{</w:t> menjadi <w:t>{{VARNAME}}</w:t>
        new_run1 = re.sub(
            r'(<w:t[^>]*>)\{\{(</w:t>)',
            f'\\1{{{{{varname}}}}}\\2',
            run1,
            count=1
        )
        return new_run1

    # Pattern: run dengan {{ -> optional stuff -> run dengan varname -> optional stuff -> run dengan }}
    # Tangani variasi dengan/tanpa proofErr, dengan/tanpa rPr, dengan/tanpa rsid attrs
    pattern = (
        # Run 1: berisi {{
        r'<w:r(?:\s[^>]*)?>(?:<w:rPr>.*?</w:rPr>)?'
        r'<w:t[^>]*>\{\{</w:t>'
        r'</w:r>'
        # Opsional: proofErr
        r'(?:<w:proofErr[^/]*/>\s*)?'
        # Run 2: berisi varname (capture)
        r'<w:r(?:\s[^>]*)?>(?:<w:rPr>.*?</w:rPr>)?'
        r'<w:t[^>]*>(?P<var>[a-zA-Z_]\w*)</w:t>'
        r'</w:r>'
        # Opsional: proofErr
        r'(?:<w:proofErr[^/]*/>\s*)?'
        # Run 3: berisi }}
        r'<w:r(?:\s[^>]*)?>(?:<w:rPr>.*?</w:rPr>)?'
        r'<w:t[^>]*>\}\}</w:t>'
        r'</w:r>'
    )

    before_count = len(re.findall(r'<w:t[^>]*>\{\{</w:t>', xml))
    xml = re.sub(pattern, collapse_match, xml, flags=re.DOTALL)
    after_count = len(re.findall(r'<w:t[^>]*>\{\{</w:t>', xml))
    print(f"  Collapsed: {before_count - after_count} split-vars "
          f"({before_count} -> {after_count} remaining open-{{{{)")
    
    # Jika masih ada split, coba variasi lain (tanpa proofErr, runs langsung berurutan)
    if after_count > 0:
        # Coba pattern yang lebih longgar — hapus semua XML di antara {{ ... }}
        def loose_collapse(m):
            inner = m.group(1)  # semua XML antara {{ dan }}
            # Ekstrak nama variabel dari inner
            varname_m = re.search(r'>([a-zA-Z_]\w*)<', inner)
            if varname_m:
                return f'<w:t>{{{{{varname_m.group(1)}}}}}</w:t>'
            return m.group(0)  # tidak diubah
        
        xml = re.sub(
            r'<w:t[^>]*>\{\{</w:t>((?:(?!</w:tr>)(?!</w:p>).)*?)<w:t[^>]*>\}\}</w:t>',
            loose_collapse,
            xml,
            flags=re.DOTALL
        )
        final_count = len(re.findall(r'<w:t[^>]*>\{\{</w:t>', xml))
        print(f"  After loose collapse: {final_count} remaining")
    
    return xml


# ─── STEP 2 & 3: Replace vars ────────────────────────────────────────────────
def replace_vars(xml: str) -> str:
    """Ganti variabel flat -> item.xxx dan hardcoded '1 JP' -> {{item.jp}}"""
    pairs = [
        (r'\{\{no\}\}', '{{item.no}}'),
        (r'\{\{waktu\}\}', '{{item.waktu}}'),
        (r'\{\{kegiatan\}\}', '{{item.kegiatan}}'),
        (r'\{\{jp\}\}', '{{item.jp}}'),
    ]
    for pat, rep in pairs:
        n = len(re.findall(pat, xml))
        if n:
            xml = re.sub(pat, rep, xml)
            print(f"  [VAR] {pat} -> {rep}  ({n}x)")
    
    # Ganti hardcoded '1 JP' di <w:t> yang ada di baris data
    # (baris yang berisi {{item.no}} atau {{item.kegiatan}})
    rows = list(re.finditer(r'<w:tr[ >].*?</w:tr>', xml, re.DOTALL))
    for rm in rows:
        row_text = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', rm.group(0), re.DOTALL))
        if ('{{item.' in row_text or '{{no}}' in row_text) and '1 JP' in row_text:
            new_row = re.sub(
                r'(<w:t[^>]*>)1 JP(</w:t>)',
                r'\1{{item.jp}}\2',
                rm.group(0)
            )
            if new_row != rm.group(0):
                xml = xml[:rm.start()] + new_row + xml[rm.end():]
                print("  [JP]  '1 JP' -> {{item.jp}}")
                break
    return xml


# ─── STEP FIX ENDFOR ──────────────────────────────────────────────────────────
def fix_split_endfor(xml: str) -> str:
    """
    Fix {% endfor %} yang terpecah menjadi 3 run terpisah:
      {% end</w:t> ... <w:t>for</w:t> ... <w:t> %}
    Digabung menjadi satu run: {% endfor %}
    """
    pattern = (
        r'(<w:t[^>]*>)'
        r'(\{%-?\s*end)'
        r'(</w:t></w:r>)'
        r'(?:<w:proofErr[^/]*/>\s*)?'
        r'<w:r[^>]*>(?:<w:rPr>.*?</w:rPr>)?'
        r'<w:t[^>]*>for</w:t>'
        r'</w:r>'
        r'(?:<w:proofErr[^/]*/>\s*)?'
        r'<w:r[^>]*>(?:<w:rPr>.*?</w:rPr>)?'
        r'<w:t[^>]*>(\s*-?%\})'
        r'</w:t></w:r>'
    )
    n_before = len(re.findall(r'\{%-?\s*end</w:t>', xml))
    xml = re.sub(pattern, r'\1{% endfor %}</w:t></w:r>', xml, flags=re.DOTALL)
    n_after = len(re.findall(r'\{%-?\s*end</w:t>', xml))
    if n_before > 0:
        print(f"  [FIX] endfor split: {n_before} -> {n_after} (fixed {n_before - n_after})")
    return xml


# ─── STEP 4: Inject {%tr for/endfor %} ─────────────────────────────────────
def inject_tr_tags(xml: str) -> str:
    """
    Temukan <w:tr> yang berisi variabel item.xxx, 
    inject {%tr for item in item_materi %} di paragraf pertama,
    tambah baris <w:tr> baru berisi {%tr endfor %}.
    """
    rows = list(re.finditer(r'<w:tr[ >].*?</w:tr>', xml, re.DOTALL))
    print(f"  Total baris tabel: {len(rows)}")

    loop_row_m = None
    for rm in rows:
        texts = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', rm.group(0), re.DOTALL))
        # Deteksi variabel item dengan atau tanpa spasi: {{item. atau {{ item.
        has_item_var = bool(re.search(r'\{\{[\s]*item\.', texts))
        if has_item_var:
            loop_row_m = rm
            print(f"  [FOUND] Baris loop: '{re.sub(chr(60)+'[^>]+'+chr(62), '', texts[:80]).strip()}'")
            break

    if loop_row_m is None:
        print("  [ERROR] Baris loop tidak ditemukan!")
        return xml

    loop_row_xml = loop_row_m.group(0)

    # Jangan inject ganda
    if '{%tr for' in loop_row_xml:
        print("  [SKIP]  {%tr for} sudah ada")
        return xml

    # Inject {%tr for %} di awal paragraf pertama di sel pertama
    # Cari <w:p ...> pertama di dalam <w:tc> pertama
    first_p_m = re.search(
        r'(<w:tc>(?:<w:tcPr>.*?</w:tcPr>)?\s*)(<w:p[^>]*>)',
        loop_row_xml, re.DOTALL
    )
    if first_p_m:
        # Ambil rPr dari run pertama untuk styling yang konsisten
        rpr_m = re.search(r'<w:r>(<w:rPr>.*?</w:rPr>)', loop_row_xml, re.DOTALL)
        rpr_xml = rpr_m.group(1) if rpr_m else ''

        for_run = (
            f'<w:r>{rpr_xml}'
            f'<w:t xml:space="preserve">{{%tr for item in item_materi %}}</w:t>'
            f'</w:r>'
        )
        insert_pos = first_p_m.end()  # setelah <w:p ...>
        loop_row_xml = loop_row_xml[:insert_pos] + for_run + loop_row_xml[insert_pos:]
        print("  [INJ]  {%tr for item in item_materi %} ✓")

    # Buat baris endfor
    trpr_m = re.search(r'<w:trPr>.*?</w:trPr>', loop_row_xml, re.DOTALL)
    trpr_xml = trpr_m.group(0) if trpr_m else ''

    tcpr_m = re.search(r'<w:tcPr>.*?</w:tcPr>', loop_row_xml, re.DOTALL)
    tcpr_xml = tcpr_m.group(0) if tcpr_m else ''

    n_cols = loop_row_xml.count('<w:tc>')

    tc_cells = []
    for i in range(n_cols):
        content = (
            '<w:r><w:t xml:space="preserve">{%tr endfor %}</w:t></w:r>'
            if i == 0 else ''
        )
        tc_cells.append(
            f'<w:tc>{tcpr_xml}<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
            f'{content}</w:p></w:tc>'
        )

    endfor_row = f'<w:tr>{trpr_xml}{"".join(tc_cells)}</w:tr>'
    print("  [INJ]  {%tr endfor %} baris baru ✓")

    # Rekonstruksi XML
    new_xml = (
        xml[:loop_row_m.start()]
        + loop_row_xml
        + '\n'
        + endfor_row
        + xml[loop_row_m.end():]
    )
    return new_xml


# ─── VERIFY ─────────────────────────────────────────────────────────────────
def verify(path: Path):
    """Tampilkan semua tag di template untuk verifikasi."""
    with zipfile.ZipFile(str(path), 'r') as z:
        xml = z.read('word/document.xml').decode('utf-8')
    tags = re.findall(r'\{%[^%]+%\}|\{\{[^{}]+\}\}', xml)
    print("Tag di template baru:")
    for t in tags[:30]:
        print(f"  {t}")
    
    # Cek apakah ada split yang tersisa
    splits = re.findall(r'<w:t[^>]*>\{\{</w:t>', xml)
    if splits:
        print(f"\n  [WARN] Masih ada {len(splits)} split '{{{{' yang belum ter-collapse!")


# ─── PROCESS ONE TEMPLATE ────────────────────────────────────────────────────
def process_template(tmpl_path: Path):
    """Proses satu file template: collapse, replace, inject, simpan."""
    backup = tmpl_path.with_name(tmpl_path.stem + "_backup.docx")

    if not backup.exists():
        shutil.copy2(str(tmpl_path), str(backup))
        print(f"  [BAK] Backup: {backup.name}")
    else:
        shutil.copy2(str(backup), str(tmpl_path))
        print(f"  [RST] Restored dari: {backup.name}")

    with zipfile.ZipFile(str(tmpl_path), 'r') as z:
        xml = z.read('word/document.xml').decode('utf-8')
    print(f"  [XML] {len(xml)} chars")

    xml = collapse_split_vars(xml)
    xml = fix_split_endfor(xml)   # Fix {% endfor %} yang terpecah
    xml = replace_vars(xml)
    xml = inject_tr_tags(xml)

    temp_path = tmpl_path.with_suffix('.new.docx')
    with zipfile.ZipFile(str(tmpl_path), 'r') as zin:
        with zipfile.ZipFile(str(temp_path), 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'word/document.xml':
                    zout.writestr(item, xml.encode('utf-8'))
                else:
                    zout.writestr(item, zin.read(item.filename))

    os.replace(str(temp_path), str(tmpl_path))
    print(f"  [OK]  Disimpan: {tmpl_path.name}")

    # Test compile
    try:
        from docxtpl import DocxTemplate
        doc = DocxTemplate(str(tmpl_path))
        doc.init_docx()
        xml_t = doc.get_xml()
        patched = doc.patch_xml(xml_t)
        import re as re2
        patched2 = re2.sub(r'<w:p([ >])', r'\n<w:p\1', patched)
        from jinja2 import Template as J2T
        J2T(patched2)
        print("  [OK]  Jinja2 compile SUKSES")
    except Exception as e:
        print(f"  [ERR] {type(e).__name__}: {e}")


# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print(" INJECT LOOP TAG KE SEMUA TEMPLATE DOCX")
    print("=" * 60)
    print()

    templates_dir = BASE_DIR / "templates"

    # Kumpulkan semua template yang perlu diproses (bukan backup)
    candidates = []
    if (templates_dir / "pkkm.docx").exists():
        candidates.append(templates_dir / "pkkm.docx")
    for f in sorted(templates_dir.glob("pkkm_*.docx")):
        if "_backup" not in f.name and ".bak" not in f.name:
            candidates.append(f)
    if TEMPLATE_PATH.exists():
        candidates.append(TEMPLATE_PATH)  # root pkkm.docx

    if not candidates:
        print("ERROR: Tidak ada template ditemukan!")
        return

    for tmpl in candidates:
        print(f"\n>> {tmpl.relative_to(BASE_DIR)}")
        try:
            process_template(tmpl)
        except PermissionError:
            print(f"  [ERR] PERMISSION DENIED - tutup file di Word lalu coba lagi!")
        except Exception as e:
            print(f"  [ERR] {type(e).__name__}: {e}")

    print()
    print("=" * 60)
    print(" SELESAI! Jalankan: python singlesheet.py")
    print("=" * 60)


if __name__ == "__main__":
    main()
