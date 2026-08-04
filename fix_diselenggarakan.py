"""
fix_diselenggarakan.py
======================
Ganti teks '{{tanggal}} di ruangan {{gugus}} {{ruangan}}' di template
menjadi '{{detail_pelaksanaan}}' agar semua tanggal+ruangan per materi
tampil di bagian "Diselenggarakan pada:" sertifikat.

Juga kembalikan tabel ke 3 kolom: No | Judul Materi | Jam Pertemuan
"""
import sys, zipfile, re, os, shutil
from pathlib import Path

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

BASE_DIR = Path(__file__).parent.resolve()
TMPL = BASE_DIR / "templates" / "pkkm.docx"
BAK  = BASE_DIR / "templates" / "pkkm_backup.docx"

# ── Konstan lebar kolom tabel 3-kolom (sama seperti template asli) ────────────
COL_W = {'no': '917', 'kegiatan': '5599', 'jp': '2500'}


def make_rpr(bold=False):
    b = '<w:b/><w:bCs/>' if bold else ''
    return (
        '<w:rPr>' + b
        + '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        '<w:color w:val="000000" w:themeColor="text1"/>'
        '<w:sz w:val="24"/><w:szCs w:val="24"/>'
        '</w:rPr>'
    )


def make_cell(width, content, align='center', bold=False):
    jc = f'<w:jc w:val="{align}"/>' if align else ''
    return (
        f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>'
        f'<w:p><w:pPr><w:spacing w:after="0"/>{jc}</w:pPr>'
        f'<w:r>{make_rpr(bold)}<w:t xml:space="preserve">{content}</w:t></w:r>'
        f'</w:p></w:tc>'
    )


def build_header_row():
    return (
        '<w:tr>'
        + make_cell(COL_W['no'],       'No',            'center', bold=True)
        + make_cell(COL_W['kegiatan'], 'Judul Materi',  'center', bold=True)
        + make_cell(COL_W['jp'],       'Jam Pertemuan', 'center', bold=True)
        + '</w:tr>'
    )


def build_loop_rows():
    for_row = (
        '<w:tr>'
        '<w:tc><w:tcPr><w:tcW w:w="' + COL_W['no'] + '" w:type="dxa"/></w:tcPr>'
        '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
        '<w:r><w:t xml:space="preserve">{%tr for item in item_materi %}</w:t></w:r>'
        '</w:p></w:tc>'
        '</w:tr>'
    )
    data_row = (
        '<w:tr>'
        + make_cell(COL_W['no'],       '{{ item.no }}',       'center')
        + make_cell(COL_W['kegiatan'], '{{ item.kegiatan }}', 'both')
        + make_cell(COL_W['jp'],       '{{ item.jp }}',       'center')
        + '</w:tr>'
    )
    return for_row + '\n' + data_row + '\n'


def build_endfor_row():
    return (
        '<w:tr>'
        '<w:tc><w:tcPr><w:tcW w:w="' + COL_W['no'] + '" w:type="dxa"/></w:tcPr>'
        '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
        '<w:r><w:t xml:space="preserve">{%tr endfor %}</w:t></w:r>'
        '</w:p></w:tc>'
        '</w:tr>\n'
    )


def build_totaljp_row():
    total_w = int(COL_W['no']) + int(COL_W['kegiatan'])
    return (
        '<w:tr>'
        f'<w:tc><w:tcPr><w:tcW w:w="{total_w}" w:type="dxa"/>'
        f'<w:gridSpan w:val="2"/></w:tcPr>'
        f'<w:p><w:pPr><w:spacing w:after="0"/><w:jc w:val="right"/></w:pPr>'
        f'<w:r>{make_rpr(bold=True)}<w:t xml:space="preserve">Total JP</w:t></w:r>'
        f'</w:p></w:tc>'
        + make_cell(COL_W['jp'], '{{total_jp}}', 'center', bold=True)
        + '</w:tr>'
    )


def main():
    print("=" * 60)
    print(" FIX TEMPLATE: detail_pelaksanaan + Tabel 3 Kolom")
    print("=" * 60)
    print()

    # Restore dari backup
    shutil.copy2(str(BAK), str(TMPL))
    print("[RST] Restored dari backup")

    with zipfile.ZipFile(str(TMPL), 'r') as z:
        xml = z.read('word/document.xml').decode('utf-8')
    print(f"[XML] {len(xml)} chars")

    # ── FIX 1: Ganti {{tanggal}} di ruangan {{gugus}} {{ruangan}} ────────────
    # dengan {{detail_pelaksanaan}}
    # Cari paragraf yang mengandung tanggal/ruangan flat variable
    paras = list(re.finditer(r'<w:p[ >].*?</w:p>', xml, re.DOTALL))
    replaced_para = False
    for p in paras:
        text = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', p.group(0), re.DOTALL))
        if '{{tanggal}}' in text or '{{ruangan}}' in text:
            # Ambil properti paragraf (pPr) dari paragraf asli untuk pertahankan styling
            ppr_m = re.search(r'<w:pPr>.*?</w:pPr>', p.group(0), re.DOTALL)
            ppr = ppr_m.group(0) if ppr_m else '<w:pPr/>'

            # Ambil rPr dari run pertama untuk pertahankan font/warna
            rpr_m = re.search(r'<w:rPr>.*?</w:rPr>', p.group(0), re.DOTALL)
            rpr = rpr_m.group(0) if rpr_m else make_rpr()

            # Buat paragraf baru dengan {{detail_pelaksanaan}}
            new_para = (
                f'<w:p>{ppr}'
                f'<w:r>{rpr}<w:t xml:space="preserve">{{{{detail_pelaksanaan}}}}</w:t></w:r>'
                f'</w:p>'
            )
            xml = xml[:p.start()] + new_para + xml[p.end():]
            print(f"[FIX] Ganti paragraf: '...{text[:60]}' → '{{{{detail_pelaksanaan}}}}'")
            replaced_para = True
            break

    if not replaced_para:
        print("[WARN] Paragraf {{tanggal}}/{{ruangan}} tidak ditemukan — periksa backup")

    # ── FIX 2: Ganti seluruh baris tabel kembali ke 3 kolom ─────────────────
    rows = list(re.finditer(r'<w:tr[ >].*?</w:tr>', xml, re.DOTALL))
    print(f"[INF] Total rows tabel: {len(rows)}")

    header_rm = data_rm = total_rm = None
    for rm in rows:
        content = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', rm.group(0), re.DOTALL))
        clean = re.sub(r'<[^>]+>', '', content).strip()
        if re.search(r'No|Judul Materi|Jam Pertemuan', clean):
            header_rm = rm
        elif re.search(r'\{\{[\s]*item\.', content) or re.search(r'\{%\s*for\s+item', content):
            data_rm = rm
        elif 'total_jp' in content or 'Total JP' in clean:
            total_rm = rm

    if header_rm and data_rm:
        tbl_start = header_rm.start()
        tbl_end   = total_rm.end() if total_rm else data_rm.end()

        new_rows = (
            build_header_row() + '\n'
            + build_loop_rows()
            + build_endfor_row()
            + build_totaljp_row() + '\n'
        )
        xml = xml[:tbl_start] + new_rows + xml[tbl_end:]
        print("[OK]  Tabel 3-kolom diterapkan")
    else:
        print("[WARN] Struktur tabel tidak ditemukan, skip")

    # ── Simpan ─────────────────────────────────────────────────────────────────
    tmp = TMPL.with_suffix('.new.docx')
    with zipfile.ZipFile(str(TMPL), 'r') as zin:
        with zipfile.ZipFile(str(tmp), 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = xml.encode('utf-8') if item.filename == 'word/document.xml' else zin.read(item.filename)
                zout.writestr(item, data)
    os.replace(str(tmp), str(TMPL))
    print(f"[OK]  Disimpan: {TMPL.name}")

    # ── Verifikasi tags ────────────────────────────────────────────────────────
    print()
    # Cek paragraf detail_pelaksanaan
    with zipfile.ZipFile(str(TMPL), 'r') as z:
        xml2 = z.read('word/document.xml').decode('utf-8')
    if 'detail_pelaksanaan' in xml2:
        print("[OK]  {{detail_pelaksanaan}} ada di template")
    else:
        print("[WARN] {{detail_pelaksanaan}} TIDAK ditemukan!")

    for i, rm in enumerate(re.finditer(r'<w:tr[ >].*?</w:tr>', xml2, re.DOTALL)):
        tags = re.findall(r'\{%[^%]+%\}|\{\{[^{}]+\}\}', rm.group(0))
        if tags:
            print(f"  Row {i}: {tags}")

    # ── Test Jinja2 ────────────────────────────────────────────────────────────
    print()
    try:
        from docxtpl import DocxTemplate
        doc = DocxTemplate(str(TMPL))
        doc.init_docx()
        patched = doc.patch_xml(doc.get_xml())
        from jinja2 import Template as J2T
        J2T(re.sub(r'<w:p([ >])', r'\n<w:p\1', patched))
        print("[OK]  Jinja2 compile SUKSES")
    except Exception as e:
        print(f"[ERR] Jinja2: {type(e).__name__}: {e}")
        return

    # ── Test render ────────────────────────────────────────────────────────────
    print()
    from docxtpl import DocxTemplate
    from docx import Document as DocxDoc

    ctx = {
        'nama': 'Herman, S.KM.,M.Med.Ed',
        'mhs': '', 'gugus': 'GUGUS UTAMA', 'kode': 'GUGUS UTAMA', 'kelas': '',
        'detail_pelaksanaan': (
            'Senin, 27 Juli 2026 di ruangan Aula Kedokteran Baru '
            'dan Selasa, 28 Juli 2026 di ruangan Aula Kedokteran Baru'
        ),
        'item_materi': [
            {'no': '1', 'waktu': '10.30-11.30', 'tanggal': 'Senin, 27 Juli 2026',
             'ruangan': 'Aula Kedokteran Baru',
             'kegiatan': 'Pengembangan Karakter Mahasiswa, Growth Mindset, Etika Akademik',
             'jp': '1 JP'},
            {'no': '2', 'waktu': '10.30-11.30', 'tanggal': 'Selasa, 28 Juli 2026',
             'ruangan': 'Aula Kedokteran Baru',
             'kegiatan': 'Materi 4. Pengembangan Karakter Mahasiswa, Growth Mindset',
             'jp': '1 JP'},
        ],
        'total_jp': '2 JP', 'jumlah_materi': 2,
    }

    out = BASE_DIR / 'output_pdf' / 'test_detail.docx'
    doc2 = DocxTemplate(str(TMPL))
    try:
        doc2.render(ctx)
        doc2.save(str(out))
        print(f"[OK]  Render: {out.name}")
        d = DocxDoc(str(out))
        print("\n=== ISI PARAGRAF (cari detail_pelaksanaan) ===")
        for para in d.paragraphs:
            if 'Senin' in para.text or 'Selasa' in para.text or 'Aula' in para.text:
                print(f"  >> {repr(para.text[:120])}")
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
