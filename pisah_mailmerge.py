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
        print(r"Contoh lokasi: C:\Program Files\LibreOffice\program\soffice.exe")
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


def proses_mail_merge_semua():
    base_dir = Path(__file__).parent.resolve()
    
    excel_path = base_dir / "datapkkm.xlsx"
    template_path = base_dir / "pkkm.docx"
    output_folder = base_dir / "output_sertifikat"

    if not excel_path.exists():
        print(f"ERROR: File Excel tidak ditemukan di: {excel_path}")
        return
    if not template_path.exists():
        print(f"ERROR: File Template Word tidak ditemukan di: {template_path}")
        return

    output_folder.mkdir(parents=True, exist_ok=True)

    # 1. Baca data dari Excel
    df = pd.read_excel(excel_path)

    # Pastikan baris kosong di nama diabaikan
    df_peserta = df[df['nama'].notnull()].copy()

    if df_peserta.empty:
        print("Tidak ada data nama peserta di file Excel.")
        return

    print("=" * 60)
    print(f"Ditemukan {len(df_peserta)} peserta. Memulai proses Mail Merge...")
    print("=" * 60)

    # 2. Iterasi untuk SETIAP peserta (dengan atau tanpa NIM)
    for idx, row in enumerate(df_peserta.itertuples(), 1):
        nama_peserta = str(row.nama).strip()
        
        # Cek apakah kolom 'mhs' terisi / ada NIM
        nim_val = getattr(row, 'mhs', None)
        if pd.notnull(nim_val) and str(nim_val).strip() != "" and str(nim_val).lower() != "nan":
            nim_peserta = str(nim_val).strip()
            text_info = f"{nama_peserta} ({nim_peserta})"
            file_prefix = f"Sertifikat_{nim_peserta}"
        else:
            nim_peserta = ""  # Kosongkan jika tidak ada NIM
            text_info = f"{nama_peserta} (Tanpa NIM)"
            file_prefix = "Sertifikat"

        print(f"\n[{idx}/{len(df_peserta)}] Memproses: {text_info}")

        # Load Template Word
        doc = DocxTemplate(template_path)

        # Context yang dikirim ke Word template
        context = {
            'nama': nama_peserta,
            'mhs': nim_peserta
        }

        doc.render(context)

        # Sanitasi nama agar aman dijadikan nama file Windows
        clean_nama = "".join(c for c in nama_peserta if c.isalnum() or c in (' ', '_', '-')).strip()
        docx_out = output_folder / f"{file_prefix}_{clean_nama}.docx"
        
        doc.save(docx_out)
        print(f"  -> File Word sementara dibuat: {docx_out.name}")

        # Konversi ke PDF
        pdf_out = convert_docx_to_pdf(docx_out, output_folder)
        print(f"  -> File PDF berhasil dibuat : {pdf_out.name}")

        # Menghapus file DOCX sementara setelah PDF terbentuk
        if pdf_out.exists():
            docx_out.unlink()

    print("\n" + "=" * 60)
    print(f"SELESAI! Semua ({len(df_peserta)}) sertifikat berhasil diproses di 'output_sertifikat'.")
    print("=" * 60)


if __name__ == "__main__":
    proses_mail_merge_semua()