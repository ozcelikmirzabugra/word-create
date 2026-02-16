#!/usr/bin/env python3
from __future__ import annotations

import datetime as dt
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


TEMPLATE_PATH = Path("/Users/mirzmac/Projects/demo_dodo/tavşanlı müzekkere üstyazısı.DOC")
DOCX_DOCUMENT_XML = "word/document.xml"


def prompt_until_valid(prompt: str, validator):
    while True:
        value = input(prompt).strip()
        if validator(value):
            return value


def validate_nonempty_digits(value: str, field_name: str) -> bool:
    if not value:
        print(f"Hata: {field_name} boş olamaz.")
        return False
    if not value.isdigit():
        print(f"Hata: {field_name} sadece rakamlardan oluşmalıdır.")
        return False
    return True


def validate_date_ddmmyyyy(value: str) -> bool:
    if not re.fullmatch(r"\d{2}/\d{2}/\d{4}", value):
        print("Hata: Tarih formatı GG/AA/YYYY olmalıdır.")
        return False
    try:
        dt.datetime.strptime(value, "%d/%m/%Y")
    except ValueError:
        print("Hata: Geçersiz tarih girdiniz.")
        return False
    return True


def run_command(args: list[str]) -> None:
    try:
        subprocess.run(args, check=True, capture_output=True, text=True)
    except subprocess.CalledProcessError as exc:
        stderr = (exc.stderr or "").strip()
        raise RuntimeError(f"Komut başarısız oldu: {' '.join(args)}\n{stderr}") from exc


def replace_exact_once(text: str, old: str, new: str, label: str) -> str:
    count = text.count(old)
    if count != 1:
        raise ValueError(
            f"Şablonda '{label}' için beklenen metin 1 kez bulunmalıydı, bulunan adet: {count} (aranan: {old!r})"
        )
    return text.replace(old, new, 1)


def replace_preparation_date(text: str, new_date: str) -> str:
    dotted = "20.10.2025"
    slashed = "20/10/2025"
    dotted_count = text.count(dotted)
    slashed_count = text.count(slashed)
    total = dotted_count + slashed_count

    if total != 1:
        raise ValueError(
            "Şablonda evrak hazırlama tarihi için beklenen metin tam 1 kez bulunmalıydı "
            f"(20.10.2025 veya 20/10/2025), bulunan adet: {total}"
        )

    if dotted_count == 1:
        return text.replace(dotted, new_date, 1)
    return text.replace(slashed, new_date, 1)


def tighten_header_date_position(xml: str) -> str:
    marker = '<w:t xml:space="preserve">ANKARA, </w:t>'
    marker_idx = xml.find(marker)
    if marker_idx == -1:
        raise ValueError("Şablonda 'ANKARA, ' başlık alanı bulunamadı.")

    run_matches = list(re.finditer(r"<w:r(?:\s[^>]*)?>", xml[:marker_idx]))
    if len(run_matches) < 2:
        raise ValueError("Şablonda ANKARA satırı run bilgisi bulunamadı.")

    ankara_run_start = run_matches[-1].start()
    spacer_run_start = run_matches[-2].start()

    spacer_run = xml[spacer_run_start:ankara_run_start]
    if "<w:tab/>" not in spacer_run:
        raise ValueError("Şablonda ANKARA satırı için beklenen tab boşluğu bulunamadı.")

    # Sağ üst tarihin alt satıra kaymaması için başlıktaki tab aralığını bir adım daralt.
    new_spacer_run = spacer_run.replace("<w:tab/>", "", 1)
    return xml[:spacer_run_start] + new_spacer_run + xml[ankara_run_start:]


def normalize_justified_paragraphs(xml: str) -> str:
    # textutil -> doc dönüşümünde aşırı kelime aralıklarını azaltmak için justified satırları sola al.
    return xml.replace('<w:jc w:val="both"/>', '<w:jc w:val="left"/>')


def update_docx_document_xml(
    input_docx: Path,
    output_docx: Path,
    ad_soyad: str,
    tc: str,
    esas_no: str,
    mahkeme_tarihi: str,
    evrak_tarihi: str,
) -> None:
    current_year = dt.datetime.now().year
    esas_value = f"{current_year}/{esas_no}"

    with zipfile.ZipFile(input_docx, "r") as zin, zipfile.ZipFile(output_docx, "w") as zout:
        has_document_xml = False

        for item in zin.infolist():
            data = zin.read(item.filename)

            if item.filename == DOCX_DOCUMENT_XML:
                has_document_xml = True
                xml = data.decode("utf-8")
                xml = replace_exact_once(xml, "Doğukan yurt", ad_soyad, "Ad Soyad")
                xml = replace_exact_once(xml, "16291090514", tc, "TC")
                xml = replace_exact_once(xml, "2025/357", esas_value, "Mahkeme Esas Numarası")
                xml = replace_exact_once(xml, "2025/258", esas_value, "İlgi Esas Numarası")
                xml = replace_exact_once(xml, "12/09/2025", mahkeme_tarihi, "Mahkeme Gönderim Tarihi")
                xml = replace_preparation_date(xml, evrak_tarihi)
                xml = tighten_header_date_position(xml)
                xml = normalize_justified_paragraphs(xml)
                data = xml.encode("utf-8")

            zout.writestr(item, data)

    if not has_document_xml:
        raise RuntimeError("DOCX içinde word/document.xml bulunamadı.")


def build_output_path(base_doc_path: Path) -> Path:
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    return base_doc_path.with_name(f"{base_doc_path.stem}_filled_{timestamp}.DOC")


def main() -> int:
    if shutil.which("textutil") is None:
        print("Hata: 'textutil' komutu bulunamadı. Bu script macOS üzerinde çalışmak üzere tasarlandı.")
        return 1

    if not TEMPLATE_PATH.exists():
        print(f"Hata: Şablon dosya bulunamadı: {TEMPLATE_PATH}")
        return 1

    ad_soyad = prompt_until_valid("1) Ad Soyad: ", lambda x: len(x) > 0)
    tc = prompt_until_valid("2) TC Kimlik No: ", lambda x: validate_nonempty_digits(x, "TC Kimlik No"))
    esas_no = prompt_until_valid(
        "3) Mahkeme Esas Numarası (sadece sayı): ",
        lambda x: validate_nonempty_digits(x, "Mahkeme Esas Numarası"),
    )
    mahkeme_tarihi = prompt_until_valid(
        "4) Mahkemenin Gönderdiği Tarih (GG/AA/YYYY): ", validate_date_ddmmyyyy
    )
    evrak_tarihi = prompt_until_valid("5) Evrakın Hazırlandığı Tarih (GG/AA/YYYY): ", validate_date_ddmmyyyy)

    output_doc = build_output_path(TEMPLATE_PATH)

    try:
        with tempfile.TemporaryDirectory(prefix="muzekkere_") as tmp:
            tmpdir = Path(tmp)
            converted_docx = tmpdir / "template.docx"
            updated_docx = tmpdir / "updated.docx"

            run_command(["textutil", "-convert", "docx", "-output", str(converted_docx), str(TEMPLATE_PATH)])
            update_docx_document_xml(
                input_docx=converted_docx,
                output_docx=updated_docx,
                ad_soyad=ad_soyad,
                tc=tc,
                esas_no=esas_no,
                mahkeme_tarihi=mahkeme_tarihi,
                evrak_tarihi=evrak_tarihi,
            )
            run_command(["textutil", "-convert", "doc", "-output", str(output_doc), str(updated_docx)])
    except Exception as exc:
        print(f"Hata: {exc}")
        return 1

    print(f"Başarılı: Yeni dosya oluşturuldu -> {output_doc}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
