import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate

# =========================================================
# HELPER: HITUNG JP DARI TEKS WAKTU
# =========================================================


def hitung_jp_dari_waktu(waktu_str):
    if not waktu_str or "-" not in str(waktu_str) and "–" not in str(waktu_str):
        return 1

    try:
        # Menangani pemisah dash biasa (-) maupun en-dash (–)
        clean_waktu = str(waktu_str).replace("–", "-")
        match = re.findall(r"(\d{1,2})[\.:](\d{2})", clean_waktu)
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
# HELPER: LIBREOFFICE CONVERTER
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
# EKSEKUSI UTAMA
# =========================================================


def proses_mail_merge_single_sheet():
    base_dir = Path(__file__).parent.resolve()

    excel_path = base_dir / "datapkkm.xlsx"
    folder_templates = base_dir / "templates"
    output_folder = base_dir / "output_pdf"

    output_folder.mkdir(parents=True, exist_ok=True)

    if not excel_path.exists():
        print(f"ERROR: File {excel_path.name} tidak ditemukan!")
        return

    print("=" * 65)
    print(" 1. MEMBACA DATA DARI EXCEL & MENGELOMPOKKAN MATERI DOSEN...")
    print("=" * 65)

    df = pd.read_excel(excel_path)
    df.columns = df.columns.str.strip().str.lower()
    df_data = df[df["nama"].notnull()].copy()

    # Grouping berdasarkan nama dosen
    grouped_dosen = df_data.groupby("nama", sort=False)

    print(f" Ditemukan {len(grouped_dosen)} dosen unik.\n")

    print("=" * 65)
    print(" 2. MEMPROSES RENDER DOCX & KONVERSI PDF...")
    print("=" * 65)

    for idx, (nama_dosen, group_df) in enumerate(grouped_dosen, 1):
        first_row = group_df.iloc[0]

        nama = str(nama_dosen).strip()
        mhs = (
            str(first_row.get("nim", ""))
            if pd.notnull(first_row.get("nim"))
            else ""
        )
        gugus = (
            str(first_row.get("gugus", ""))
            if pd.notnull(first_row.get("gugus"))
            else ""
        )
        kode = (
            str(first_row.get("kode", ""))
            if pd.notnull(first_row.get("kode"))
            else ""
        )
        kelas = (
            str(first_row.get("kelas", ""))
            if pd.notnull(first_row.get("kelas"))
            else ""
        )
        fakultas = (
            str(first_row.get("fakultas", "default")).lower()
            if pd.notnull(first_row.get("fakultas"))
            else "default"
        )

        item_materi = []
        total_jp_angka = 0
        rincian_pelaksanaan = []

        # MEMBACA SETIAP BARIS MATERI SECARA PERSISI
        for sub_idx, (_, row) in enumerate(group_df.iterrows(), 1):
            waktu = (
                str(row.get("waktu", "")).strip()
                if pd.notnull(row.get("waktu"))
                else ""
            )

            kegiatan_clean = re.sub(
                r"\s+",
                " ",
                str(row.get("kegiatan", ""))
                if pd.notnull(row.get("kegiatan"))
                else "",
            ).strip()

            tanggal_item = (
                str(row.get("tanggal", "")).strip()
                if pd.notnull(row.get("tanggal"))
                and str(row.get("tanggal")).strip().lower() != "nan"
                else ""
            )

            ruangan_item = (
                str(row.get("ruangan", "")).strip()
                if pd.notnull(row.get("ruangan"))
                and str(row.get("ruangan")).strip().lower() != "nan"
                else ""
            )

            jp_hitung = hitung_jp_dari_waktu(waktu)
            total_jp_angka += jp_hitung

            item_materi.append({
                "no": str(sub_idx),
                "waktu": waktu,
                "kegiatan": kegiatan_clean,
                "tanggal": tanggal_item,
                "ruangan": ruangan_item,
                "jp": f"{jp_hitung} JP",
            })

            # Susun teks rincian pelaksanaan per materi dari Excel
            if tanggal_item and ruangan_item:
                tek_sesi = f"{tanggal_item} di ruangan {ruangan_item}"
            elif tanggal_item:
                tek_sesi = tanggal_item
            elif ruangan_item:
                tek_sesi = f"ruangan {ruangan_item}"
            else:
                tek_sesi = ""

            if tek_sesi:
                rincian_pelaksanaan.append(tek_sesi)

        # Gabungkan rincian pelaksanaan seluruh materi dengan kata "dan"
        if len(rincian_pelaksanaan) > 1:
            detail_pelaksanaan = " dan ".join([
                ", ".join(rincian_pelaksanaan[:-1]),
                rincian_pelaksanaan[-1],
            ])
        elif len(rincian_pelaksanaan) == 1:
            detail_pelaksanaan = rincian_pelaksanaan[0]
        else:
            detail_pelaksanaan = ""

        str_total_jp = f"{total_jp_angka} JP"

        template_spesifik = folder_templates / f"pkkm_{fakultas}.docx"
        template_default = folder_templates / "pkkm.docx"

        if template_spesifik.exists():
            template_path = template_spesifik
        elif template_default.exists():
            template_path = template_default
        else:
            template_path = base_dir / "pkkm.docx"

        # Context Variabel untuk Word Template
        context = {
            "nama": nama,
            "mhs": mhs,
            "gugus": gugus,
            "kode": kode,
            "kelas": kelas,
            "detail_pelaksanaan": detail_pelaksanaan,
            "item_materi": item_materi,
            "total_jp": str_total_jp,
            "jumlah_materi": len(item_materi),
        }

        doc = DocxTemplate(template_path)
        doc.render(context)

        # Penamaan File Output: Sertifikat_[Gugus]_[Nama]
        clean_nama = re.sub(r'[\\/*?:"<>|]', "", nama).strip()
        val_gugus = gugus if gugus else (kode if kode else "Gugus")
        clean_gugus = re.sub(r'[\\/*?:"<>|]', "", str(val_gugus)).strip()

        filename = f"Sertifikat_{clean_gugus}_{clean_nama}.docx"
        docx_out = output_folder / filename
        doc.save(docx_out)

        # Konversi Ke PDF
        pdf_out = convert_docx_to_pdf(docx_out, output_folder)

        if pdf_out.exists():
            docx_out.unlink()

        print(
            f"[{idx}/{len(grouped_dosen)}] {pdf_out.name} ({len(item_materi)} Materi | {str_total_jp}) [PDF OK]"
        )

    print("\n" + "=" * 65)
    print(" PROSES SELESAI! Seluruh sertifikat PDF berhasil dibuat.")
    print("=" * 65)


if __name__ == "__main__":
    proses_mail_merge_single_sheet()