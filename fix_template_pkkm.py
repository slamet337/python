"""
fix_template_pkkm.py (v5)
==========================
Ganti baris header dan data di templates/pkkm.docx:
  No | Tanggal | Ruangan | Judul Materi | Jam Pertemuan
menggunakan {%tr for/endfor %} untuk row-level iteration.
"""
import sys, zipfile, re, os, shutil
from pathlib import Path

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

BASE_DIR = Path(__file__).parent.resolve()
TMPL = BASE_DIR / "templates" / "pkkm.docx"
BAK  = BASE_DIR / "templates" / "pkkm_backup.docx"

# Lebar kolom (total ~9016 dxa) untuk 5 kolom:
# No | Tanggal | Ruangan | Judul Materi | Jam Pertemuan
COL_W = {
    'no':       '600',
    'tanggal':  '2200',
    'ruangan':  '2216',
    'kegiatan': '2500',
    'jp':       '1500',
}


def make_rpr(bold=False):
    """Format rPr Times New Roman 12pt."""
    b = '<w:b/><w:bCs/>' if bold else ''
    return (
        '<w:rPr>'
        + b
        + '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        '<w:color w:val="000000" w:themeColor="text1"/>'
        '<w:sz w:val="24"/><w:szCs w:val="24"/>'
        '</w:rPr>'
    )


def make_cell(width, content, align='center', bold=False):
    """Buat satu sel tabel Word."""
    jc = f'<w:jc w:val="{align}"/>' if align else ''
    return (
        f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>'
        f'<w:p><w:pPr><w:spacing w:after="0"/>{jc}</w:pPr>'
        f'<w:r>{make_rpr(bold)}<w:t xml:space="preserve">{content}</w:t></w:r>'
        f'</w:p></w:tc>'
    )


def build_header_row() -> str:
    """Buat baris header: No | Tanggal | Ruangan | Judul Materi | Jam Pertemuan."""
    return (
        '<w:tr>'
        + make_cell(COL_W['no'],       'No',            'center', bold=True)
        + make_cell(COL_W['tanggal'],  'Tanggal',       'center', bold=True)
        + make_cell(COL_W['ruangan'],  'Ruangan',       'center', bold=True)
        + make_cell(COL_W['kegiatan'], 'Judul Materi',  'center', bold=True)
        + make_cell(COL_W['jp'],       'Jam Pertemuan', 'center', bold=True)
        + '</w:tr>'
    )


def build_loop_rows() -> str:
    """
    Buat 2 baris untuk loop {%tr for/endfor %}:
    1. Baris FOR  → diganti {% for item in item_materi %} oleh docxtpl
    2. Baris DATA → diulang Jinja2 untuk setiap item
    """
    # Baris FOR (akan dihapus oleh docxtpl)
    for_row = (
        '<w:tr>'
        '<w:tc><w:tcPr><w:tcW w:w="' + COL_W['no'] + '" w:type="dxa"/></w:tcPr>'
        '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
        '<w:r><w:t xml:space="preserve">{%tr for item in item_materi %}</w:t></w:r>'
        '</w:p></w:tc>'
        '</w:tr>'
    )

    # Baris DATA — diulang per item
    data_row = (
        '<w:tr>'
        + make_cell(COL_W['no'],       '{{ item.no }}',       'center')
        + make_cell(COL_W['tanggal'],  '{{ item.tanggal }}',  'center')
        + make_cell(COL_W['ruangan'],  '{{ item.ruangan }}',  'center')
        + make_cell(COL_W['kegiatan'], '{{ item.kegiatan }}', 'both')
        + make_cell(COL_W['jp'],       '{{ item.jp }}',       'center')
        + '</w:tr>'
    )

    return for_row + '\n' + data_row + '\n'


def build_endfor_row() -> str:
    """Baris penutup loop {%tr endfor %}."""
    return (
        '<w:tr>'
        '<w:tc><w:tcPr><w:tcW w:w="' + COL_W['no'] + '" w:type="dxa"/></w:tcPr>'
        '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
        '<w:r><w:t xml:space="preserve">{%tr endfor %}</w:t></w:r>'
        '</w:p></w:tc>'
        '</w:tr>\n'
    )


def build_totaljp_row() -> str:
    """Baris Total JP dengan gridSpan (merge kolom No s.d. Ruangan+Kegiatan)."""
    # Merge 4 kolom kiri menjadi 1, lalu kolom jp
    total_w = int(COL_W['no']) + int(COL_W['tanggal']) + int(COL_W['ruangan']) + int(COL_W['kegiatan'])
    return (
        '<w:tr>'
        f'<w:tc><w:tcPr><w:tcW w:w="{total_w}" w:type="dxa"/>'
        f'<w:gridSpan w:val="4"/></w:tcPr>'
        f'<w:p><w:pPr><w:spacing w:after="0"/><w:jc w:val="right"/></w:pPr>'
        f'<w:r>{make_rpr(bold=True)}<w:t xml:space="preserve">Total JP</w:t></w:r>'
        f'</w:p></w:tc>'
        + make_cell(COL_W['jp'], '{{total_jp}}', 'center', bold=True)
        + '</w:tr>'
    )


def build_new_data_row(header_row_xml: str) -> str:
    """Alias lama - sekarang memanggil build_loop_rows tanpa parameter."""
    return build_loop_rows()


def main():
    print("=" * 60)
    print(" FIX TEMPLATE - 5 Kolom: No|Tanggal|Ruangan|Materi|JP")
    print("=" * 60)
    print()

    shutil.copy2(str(BAK), str(TMPL))
    print("[RST] Restored dari backup")

    with zipfile.ZipFile(str(TMPL), 'r') as z:
        xml = z.read('word/document.xml').decode('utf-8')
    print(f"[XML] {len(xml)} chars")

    rows = list(re.finditer(r'<w:tr[ >].*?</w:tr>', xml, re.DOTALL))
    print(f"[INF] Total rows: {len(rows)}")

    # Identifikasi baris header, data, dan total_jp
    header_rm = data_rm = total_rm = None

    for rm in rows:
        content = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', rm.group(0), re.DOTALL))
        clean = re.sub(r'<[^>]+>', '', content).strip()

        if re.search(r'No|Judul Materi|Jam Pertemuan', clean):
            header_rm = rm
            print(f"[HDR] Header: '{clean[:60]}'")
        elif re.search(r'\{\{[\s]*item\.', content) or re.search(r'\{%\s*for\s+item', content):
            data_rm = rm
            print(f"[DAT] Data row ditemukan")
        elif 'total_jp' in content or 'Total JP' in clean:
            total_rm = rm
            print(f"[TOT] Total JP row ditemukan")

    if not header_rm or not data_rm:
        print("[ERR] Baris penting tidak ditemukan!")
        return

    # ── Ganti seluruh isi tabel (header s.d. total_jp) sekaligus ──────────────
    tbl_start = header_rm.start()
    tbl_end   = total_rm.end() if total_rm else data_rm.end()

    new_rows = (
        build_header_row()  + '\n'
        + build_loop_rows()
        + build_endfor_row()
        + build_totaljp_row() + '\n'
    )

    xml = xml[:tbl_start] + new_rows + xml[tbl_end:]
    print("\n[OK]  Tabel diganti: header + loop + endfor + total_jp")


    # Simpan
    tmp = TMPL.with_suffix('.new.docx')
    with zipfile.ZipFile(str(TMPL), 'r') as zin:
        with zipfile.ZipFile(str(tmp), 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = xml.encode('utf-8') if item.filename == 'word/document.xml' else zin.read(item.filename)
                zout.writestr(item, data)
    os.replace(str(tmp), str(TMPL))
    print(f"\n[OK]  Disimpan: {TMPL.name}")

    # Verifikasi
    print()
    rows3 = list(re.finditer(r'<w:tr[ >].*?</w:tr>', xml, re.DOTALL))
    for i, rm in enumerate(rows3):
        tags = re.findall(r'\{%[^%]+%\}|\{\{[^{}]+\}\}', rm.group(0))
        if tags:
            print(f"  Row {i}: {tags}")

    # Test Jinja2 compile
    print()
    try:
        from docxtpl import DocxTemplate
        doc = DocxTemplate(str(TMPL))
        doc.init_docx()
        xml_t = doc.get_xml()
        patched = doc.patch_xml(xml_t)
        patched2 = re.sub(r'<w:p([ >])', r'\n<w:p\1', patched)
        from jinja2 import Template as J2T
        J2T(patched2)
        print("[OK]  Jinja2 compile SUKSES")
    except Exception as e:
        print(f"[ERR] Jinja2: {type(e).__name__}: {e}")
        return

    # Test render lengkap
    print()
    from docxtpl import DocxTemplate
    from docx import Document as DocxDoc

    ctx = {
        'nama': 'Dedy Farhamsa S.Si. M.T.',
        'mhs': '12345',
        'gugus': 'Gugus 4 FAHUM',
        'kode': 'Gugus 4',
        'kelas': 'BT 1',
        'tanggal': '27 Juli 2026',
        'ruangan': 'Ruangan Gugus 4',
        'item_materi': [
            {'no': '1', 'waktu': '08.00-10.00', 'kegiatan': 'Sistem Pendidikan Tinggi', 'jp': '1 JP'},
            {'no': '2', 'waktu': '10.00-12.00', 'kegiatan': 'Kehidupan Berbangsa dan Bernegara', 'jp': '1 JP'},
        ],
        'total_jp': '2 JP',
        'jumlah_materi': 2,
    }

    out = BASE_DIR / 'output_pdf' / 'test_final4.docx'
    doc2 = DocxTemplate(str(TMPL))
    try:
        doc2.render(ctx)
        doc2.save(str(out))
        print(f"[OK]  Render: {out.name}")

        d = DocxDoc(str(out))
        print("\n=== ISI TABEL ===")
        for tbl in d.tables:
            for j, row in enumerate(tbl.rows):
                cells = [c.text.strip() for c in row.cells]
                if any(cells):
                    print(f"  Row {j}: {cells}")
    except Exception as e:
        print(f"[ERR] Render: {type(e).__name__}: {e}")

    print()
    print("=" * 60)
    print(" SELESAI! Jalankan: python singlesheet.py")
    print("=" * 60)


if __name__ == '__main__':
    main()
