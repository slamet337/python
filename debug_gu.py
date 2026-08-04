"""Test render 1 dosen GUGUS UTAMA dan tangkap error-nya."""
import sys, pandas as pd, re, subprocess, shutil
from pathlib import Path
from docxtpl import DocxTemplate

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

BASE = Path('.')
df = pd.read_excel('datapkkm.xlsx')
df.columns = df.columns.str.strip().str.lower()
df_data = df[df['nama'].notnull()].copy()

def hitung_jp(waktu_str):
    if not waktu_str or ('-' not in str(waktu_str) and '–' not in str(waktu_str)):
        return 1
    try:
        clean = str(waktu_str).replace('–', '-')
        match = re.findall(r'(\d{1,2})[\.:](\ d{2})', clean)
        if len(match) >= 2:
            return max(1, round(((int(match[1][0])*60+int(match[1][1]))-(int(match[0][0])*60+int(match[0][1])))/60))
    except:
        pass
    return 1

grouped = df_data.groupby('nama', sort=False)
for nama, grp in grouped:
    gugus = str(grp.iloc[0].get('gugus', '')).strip()
    if 'UTAMA' not in gugus.upper():
        continue

    first_row = grp.iloc[0]
    print(f"=== Testing: {nama} ===")

    # Build context
    item_materi = []
    rincian = []
    total_jp = 0
    for i, (_, row) in enumerate(grp.iterrows(), 1):
        waktu = str(row.get('waktu', '')).strip() if pd.notnull(row.get('waktu')) else ''
        kegiatan = re.sub(r'\s+', ' ', str(row.get('kegiatan', '')) if pd.notnull(row.get('kegiatan')) else '').strip()
        tanggal = str(row.get('tanggal', '')).strip() if pd.notnull(row.get('tanggal')) and str(row.get('tanggal')).strip().lower() != 'nan' else ''
        ruangan = str(row.get('ruangan', '')).strip() if pd.notnull(row.get('ruangan')) and str(row.get('ruangan')).strip().lower() != 'nan' else ''
        jp = 1
        total_jp += jp
        item_materi.append({'no': str(i), 'waktu': waktu, 'kegiatan': kegiatan,
                            'tanggal': tanggal, 'ruangan': ruangan, 'jp': f'{jp} JP'})
        if tanggal and ruangan:
            rincian.append(f'{tanggal} di ruangan {ruangan}')
        elif tanggal:
            rincian.append(tanggal)

    if len(rincian) > 1:
        detail = ' dan '.join([', '.join(rincian[:-1]), rincian[-1]])
    else:
        detail = rincian[0] if rincian else ''

    mhs  = str(first_row.get('nim', '')) if pd.notnull(first_row.get('nim')) else ''
    kode = str(first_row.get('kode', '')) if pd.notnull(first_row.get('kode')) else ''

    ctx = {
        'nama': nama, 'mhs': mhs, 'gugus': gugus,
        'kode': kode, 'kelas': '',
        'detail_pelaksanaan': detail,
        'item_materi': item_materi,
        'total_jp': f'{total_jp} JP',
        'jumlah_materi': len(item_materi),
    }

    print(f"  Context: detail='{detail[:60]}'")
    print(f"  item_materi[0]: {item_materi[0]}")

    tmpl_path = BASE / 'templates' / 'pkkm.docx'
    try:
        doc = DocxTemplate(str(tmpl_path))
        doc.render(ctx)
        clean_nama  = re.sub(r'[\\/*?"<>|]', '', str(nama)).strip()
        clean_gugus = re.sub(r'[\\/*?"<>|]', '', str(gugus)).strip()
        out_docx = BASE / 'output_pdf' / f'Sertifikat_{clean_gugus}_{clean_nama}.docx'
        doc.save(str(out_docx))
        print(f"  [OK] DOCX saved: {out_docx.name}")

        # Konversi ke PDF
        lo = r"C:\Program Files\LibreOffice\program\soffice.exe"
        prof = (BASE / 'output_pdf' / '.libreoffice_tmp').resolve().as_uri()
        cmd = [lo, f'-env:UserInstallation={prof}', '--headless', '--convert-to', 'pdf',
               '--outdir', str((BASE / 'output_pdf').resolve()), str(out_docx.resolve())]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        pdf = out_docx.with_suffix('.pdf')
        if pdf.exists():
            out_docx.unlink()
            print(f"  [OK] PDF: {pdf.name}")
        else:
            print(f"  [ERR] PDF tidak dibuat!")
            print(f"  STDOUT: {result.stdout[:200]}")
            print(f"  STDERR: {result.stderr[:200]}")
    except Exception as e:
        print(f"  [ERR] {type(e).__name__}: {e}")

    print()
    break  # hanya test 1 dosen dulu
