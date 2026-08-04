import shutil
import subprocess
import sys
from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate


def find_libreoffice():
    """
    Mencari LibreOffice/soffice.exe di lokasi umum Windows atau PATH.
    """
    possible_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        r"C:\Program Files\LibreOffice\program\soffice.com",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.com",
    ]

    for path in possible_paths:
        if Path(path).exists():
            return path

    for command in ["soffice.exe", "soffice", "libreoffice"]:
        found = shutil.which(command)
        if found:
            return found

    return None


def convert_docx_to_pdf(docx_path, output_dir):
    """
    Mengubah DOCX menjadi PDF menggunakan LibreOffice.
    """
    docx_path = Path(docx_path)
    output_dir = Path(output_dir)

    output_dir.mkdir(parents=True, exist_ok=True)

    libreoffice = find_libreoffice()

    if not libreoffice:
        print("\n" + "=" * 60)
        print("ERROR: LibreOffice tidak ditemukan.")
        print("=" * 60)
        print("\nSilakan periksa lokasi instalasi LibreOffice Anda.")
        sys.exit(1)

    temp_profile_dir = output_dir / ".libreoffice_tmp"
    profile_url = temp_profile_dir.resolve().as_uri()

    cmd = [
        libreoffice,
        f"-env:UserInstallation={profile_url}",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(output_dir.resolve()),
        str(docx_path.resolve()),
    ]

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=180,
            check=False
        )
    except subprocess.TimeoutExpired:
        print("\nERROR: Proses konversi melebihi batas waktu (timeout 180 detik).")
        sys.exit(1)
    except Exception as e:
        print(f"\nGagal menjalankan LibreOffice: {e}")
        sys.exit(1)

    pdf_path = output_dir / f"{docx_path.stem}.pdf"

    if not pdf_path.exists():
        print("\nPDF tidak berhasil dibuat.")
        print("Output LibreOffice (stdout):", result.stdout)
        print("Output LibreOffice (stderr):", result.stderr)
        sys.exit(1)

    return pdf_path


def proses_mail_merge_sk():
    base_dir = Path(__file__).parent.resolve()
    
    excel_path = base_dir / "namatotdosen.xlsx"
    template_path = base_dir / "Sertifikat_tot.docx"  # Sesuaikan dengan nama template Word sertifikat kamu
    output_folder = base_dir / "output_sertifikat_tot"

    if not excel_path.exists():
        print(f"ERROR: File Excel tidak ditemukan di: {excel_path}")
        return
    if not template_path.exists():
        print(f"ERROR: File Template Word tidak ditemukan di: {template_path}")
        return

    output_folder.mkdir(parents=True, exist_ok=True)

    # 1. Baca data dari Excel SK.xlsx
    df = pd.read_excel(excel_path)

    # Bersihkan baris yang tidak memiliki Nama
    df_peserta = df[df['nama'].notnull()].copy()

    if df_peserta.empty:
        print("Tidak ada data nama peserta di file Excel SK.xlsx.")
        return

    print("=" * 60)
    print(f"Ditemukan {len(df_peserta)} data peserta dari namatotdosen.xlsx. Memulai Mail Merge...")
    print("=" * 60)

    # 2. Iterasi untuk setiap peserta
    for idx, row in enumerate(df_peserta.itertuples(), 1):
        nama = str(getattr(row, 'nama', '')).strip()
        
        # Penanganan nilai kosong/NaN pada kolom sebagai, nim, seksi, dan gugus
        sebagai_val = getattr(row, 'sebagai', '')
        sebagai = str(sebagai_val).strip() if pd.notnull(sebagai_val) and str(sebagai_val).lower() != 'nan' else ''

        # nim_val = getattr(row, 'nim', '')
        # nim = str(nim_val).strip() if pd.notnull(nim_val) and str(nim_val).lower() != 'nan' else ''

        # seksi_val = getattr(row, 'seksi', '')
        # seksi = str(seksi_val).strip() if pd.notnull(seksi_val) and str(seksi_val).lower() != 'nan' else ''

        fakultas_val = getattr(row, 'fakultas', '')
        fakultas = str(fakultas_val).strip() if pd.notnull(fakultas_val) and str(fakultas_val).lower() != 'nan' else ''

        # gugus_val = getattr(row, 'gugus', '')
        # gugus = str(gugus_val).strip() if pd.notnull(gugus_val) and str(gugus_val).lower() != 'nan' else ''

        print(f"\n[{idx}/{len(df_peserta)}] Memproses: {nama}")

        # Load Template Word
        doc = DocxTemplate(template_path)

        # Context variabel sesuai placeholder yang ada di file Word:
        # {{ nama }}, {{ nim }}, {{ seksi }}, {{ sebagai }}, {{ gugus }}
        context = {
            'nama': nama,
            # 'nim': nim,
            # 'seksi': seksi,
            # 'sebagai': sebagai
            'fakultas' : fakultas
            
        }

        doc.render(context)

        # Sanitasi teks agar aman dijadikan nama file Windows
        clean_nama = "".join(c for c in nama if c.isalnum() or c in (' ', '_', '-')).strip()
        # clean_gugus = "".join(c for c in gugus if c.isalnum() or c in (' ', '_', '-')).strip()
        clean_fakultas = "".join(c for c in fakultas if c.isalnum() or c in (' ', '_', '-')).strip()
        
        # Buat penamaan file dinamis
        # Format: Sertifikat_[Gugus]_[NIM]_[Nama].pdf
        # prefix_gugus = f"{clean_gugus}_" if clean_gugus else ""
        # prefix_seksi = f"{clean_seksi}_" if clean_seksi else ""
        # prefix_nim = f"{nim}_" if nim else ""
        prefix_fakultas = f"{clean_fakultas}_" if clean_fakultas else ""

        # filename = f"Sertifikat_{prefix_seksi}{prefix_nim}{clean_nama}.docx"
        filename = f"Sertifikat_{prefix_fakultas}{clean_nama}.docx"
        docx_out = output_folder / filename
        
        doc.save(docx_out)
        print(f"  -> File Word sementara dibuat: {docx_out.name}")

        # Konversi ke PDF
        pdf_out = convert_docx_to_pdf(docx_out, output_folder)
        print(f"  -> File PDF berhasil dibuat : {pdf_out.name}")

        # Menghapus file DOCX sementara setelah PDF terbentuk
        if pdf_out.exists():
            docx_out.unlink()

    print("\n" + "=" * 60)
    print(f"SELESAI! Semua ({len(df_peserta)}) sertifikat berhasil diproses di folder 'output_sertifikat_tot'.")
    print("=" * 60)


if __name__ == "__main__":
    proses_mail_merge_sk()