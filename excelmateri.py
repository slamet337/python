import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate

# =========================================================
# HELPER: BERSIHKAN NAMA UNTUK MATCHING AKURAT
# =========================================================


def bersihkan_nama(nama):
    if not nama or pd.isna(nama):
        return ""
    nama_str = str(nama)
    gelar_pattern = r"\b(dr|drh|prof|ir|drs|dra|h|hj|s\.h|m\.h|s\.si|m\.t|m\.si|s\.pd|m\.pd|s\.kom|m\.kom|s\.e|m\.m|ll\.m|ph\.d)\b"
    clean = re.sub(gelar_pattern, "", nama_str, flags=re.IGNORECASE)
    clean = re.sub(r"[^a-zA-Z\s]", " ", clean)
    return re.sub(r"\s+", " ", clean).strip().lower()


def is_nama_cocok(nama_excel, nama_jadwal):
    n1 = bersihkan_nama(nama_excel)
    n2 = bersihkan_nama(nama_jadwal)
    if not n1 or not n2:
        return False
    if n1 in n2 or n2 in n1:
        return True
    words1 = set(w for w in n1.split() if len(w) > 2)
    words2 = set(w for w in n2.split() if len(w) > 2)
    if words1 and words2:
        overlap = words1.intersection(words2)
        if len(overlap) >= min(len(words1), len(words2)):
            return True
    return False


def hitung_jp_dari_waktu(waktu_str):
    """Menghitung durasi JP secara otomatis dari string 'HH.MM-HH.MM'."""
    if not waktu_str or "-" not in str(waktu_str):
        return 1

    try:
        match = re.findall(r"(\d{1,2})[\.:](\d{2})", str(waktu_str))
        if len(match) >= 2:
            jam_mulai, menit_mulai = int(match[0][0]), int(match[0][1])
            jam_selesai, menit_selesai = int(match[1][0]), int(match[1][1])

            total_menit = (jam_selesai * 60 + menit_selesai) - (
                jam_mulai * 60 + menit_mulai
            )

            jp = round(total_menit / 60)
            return max(1, jp)
    except Exception:
        pass

    return 1


# =========================================================
# KONVERSI LIBREOFFICE DOCX KE PDF
# =========================================================


def find_libreoffice():
    possible_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
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
    docx_path = Path(docx_path)
    output_dir = Path(output_dir)
    libreoffice = find_libreoffice()

    if not libreoffice:
        print("ERROR: LibreOffice tidak ditemukan!")
        sys.exit(1)

    temp_profile_dir = output_dir / ".libreoffice_tmp"
    profile_url = temp_profile_dir.resolve().as_uri()

    cmd = [
        libreoffice,
        f"-env:UserInstallation={profile_url}",
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(output_dir.resolve()),
        str(docx_path.resolve()),
    ]

    try:
        subprocess.run(
            cmd, capture_output=True, text=True, timeout=180, check=False
        )
    except Exception as e:
        print(f"Gagal Konversi: {e}")
        sys.exit(1)

    return output_dir / f"{docx_path.stem}.pdf"


# =========================================================
# EKSEKUSI UTAMA (SEMUA BACA EXCEL)
# =========================================================


def proses_mail_merge_full_excel():
    base_dir = Path(__file__).parent.resolve()

    excel_path = base_dir / "datapkkm.xlsx"
    folder_templates = base_dir / "templates"
    output_folder = base_dir / "output_pdf"

    output_folder.mkdir(parents=True, exist_ok=True)

    if not excel_path.exists():
        print(f"ERROR: File {excel_path.name} tidak ditemukan!")
        return

    print("=" * 65)
    print(" 1. MEMBACA DATA DARI EXCEL...")
    print("=" * 65)

    # Baca sheet Dosen dan sheet Jadwal dari file Excel yang sama
    excel_file = pd.ExcelFile(excel_path)

    # Jika sheet terpisah, sesuaikan nama sheet-nya (contoh: 'Dosen' dan 'Jadwal')
    # Jika hanya ada 1 sheet, pandas akan membaca sheet pertama secara default
    df_dosen = pd.read_excel(excel_file, sheet_name=0)

    try:
        df_jadwal = pd.read_excel(excel_file, sheet_name="Jadwal")
    except Exception:
        # Jika sheet Jadwal berada di file excel terpisah (misal: 'jadwal.xlsx')
        path_jadwal = base_dir / "jadwal.xlsx"
        if path_jadwal.exists():
            df_jadwal = pd.read_excel(path_jadwal)
        else:
            df_jadwal = df_dosen.copy()

    # Filter baris yang tidak kosong
    df_dosen = df_dosen[df_dosen["nama"].notnull()].copy()
    list_jadwal = df_jadwal.to_dict(orient="records")

    print(
        f" Ditemukan {len(df_dosen)} dosen & {len(list_jadwal)} data jadwal.\n"
    )

    print("=" * 65)
    print(" 2. MEMPROSES MERGE DAN MERENDER PDF...")
    print("=" * 65)

    for idx, row in enumerate(df_dosen.itertuples(), 1):
        nama = str(getattr(row, "nama", "")).strip()
        mhs = str(getattr(row, "nim", "")) if hasattr(row, "nim") else ""

        # Cari data jadwal pemateri dari list jadwal Excel
        data_jadwal = None
        for item in list_jadwal:
            pemateri_jadwal = str(item.get("pemateri", ""))
            if is_nama_cocok(nama, pemateri_jadwal):
                data_jadwal = item
                break

        # Hitung Nilai JP
        waktu_val = str(data_jadwal.get("waktu", "")) if data_jadwal else ""
        jp_dosen = hitung_jp_dari_waktu(waktu_val)
        str_jp = f"{jp_dosen} JP"

        # Deteksi Fakultas / Template
        fakultas_key = (
            str(data_jadwal.get("fakultas", "default")).lower()
            if data_jadwal
            else "default"
        )
        template_spesifik = folder_templates / f"pkkm_{fakultas_key}.docx"
        template_default = folder_templates / "pkkm.docx"

        if template_spesifik.exists():
            template_path = template_spesifik
        elif template_default.exists():
            template_path = template_default
        else:
            template_path = base_dir / "pkkm.docx"

        # Context Variabel
        context = {
            "nama": nama,
            "mhs": mhs,
            "gugus": data_jadwal.get("gugus", "") if data_jadwal else "",
            "kode": data_jadwal.get("kode", "") if data_jadwal else "",
            "kelas": data_jadwal.get("kelas", "") if data_jadwal else "",
            "tanggal": data_jadwal.get("tanggal", "") if data_jadwal else "",
            "no": "1",  # Selalu dimulai dari angka 1
            "waktu": waktu_val,
            "kegiatan": (
                data_jadwal.get("kegiatan", "") if data_jadwal else ""
            ),
            "ruangan": data_jadwal.get("ruangan", "") if data_jadwal else "",
            "jp": str_jp,
            "total_jp": str_jp,
        }

        # Render Word
        doc = DocxTemplate(template_path)
        doc.render(context)

        # Penamaan File: Sertifikat_[Gugus]_[Nama]
        clean_nama = "".join(
            c for c in nama if c.isalnum() or c in (" ", "_", "-")
        ).strip()
        val_gugus = (
            str(data_jadwal.get("kode", ""))
            if (data_jadwal and data_jadwal.get("kode"))
            else (
                data_jadwal.get("gugus", "Gugus") if data_jadwal else "Gugus"
            )
        )
        clean_gugus = "".join(
            c for c in str(val_gugus) if c.isalnum() or c in (" ", "_", "-")
        ).strip()

        filename = f"Sertifikat_{clean_gugus}_{clean_nama}.docx"
        docx_out = output_folder / filename
        doc.save(docx_out)

        # Konversi Ke PDF
        pdf_out = convert_docx_to_pdf(docx_out, output_folder)

        if pdf_out.exists():
            docx_out.unlink()  # Hapus file docx temporary

        if data_jadwal:
            print(f"[{idx}/{len(df_dosen)}] {pdf_out.name} [PDF OK]")
        else:
            print(
                f"[{idx}/{len(df_dosen)}] ⚠️  {pdf_out.name} --> [JADWAL TIDAK DITEMUKAN DARI EXCEL!]"
            )

    print("\n" + "=" * 65)
    print(" PROSES SELESAI! Seluruh sertifikat PDF berhasil dibuat.")
    print("=" * 65)


if __name__ == "__main__":
    proses_mail_merge_full_excel()