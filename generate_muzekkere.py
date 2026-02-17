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
PUANTAJ_TEMPLATE_PATH = Path("/Users/mirzmac/Projects/demo_dodo/tavşanlı puantaj talebine ilişkin örnek üstyazı.DOC")
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


def replace_by_pattern_once(text: str, pattern: str, replacement: str, label: str) -> str:
    matches = re.findall(pattern, text, flags=re.DOTALL)
    if len(matches) != 1:
        raise ValueError(
            f"Şablonda '{label}' için beklenen alan 1 kez bulunmalıydı, bulunan adet: {len(matches)}"
        )
    return re.sub(pattern, replacement, text, count=1, flags=re.DOTALL)


def replace_header_date(xml: str, new_date: str) -> str:
    pattern = (
        r'(<w:t xml:space="preserve">ANKARA, </w:t>\s*</w:r>\s*<w:r(?:\s[^>]*)?>\s*'
        r'<w:rPr>.*?</w:rPr>\s*<w:t xml:space="preserve">)([^<]+)(</w:t>)'
    )
    replacement = r"\g<1>" + new_date + r"\g<3>"
    return replace_by_pattern_once(xml, pattern, replacement, "Evrak Hazırlama Tarihi")


def replace_company_case_number(xml: str, case_value: str) -> str:
    # Tek run formatı: PARK A 2025/357
    single_run_pattern = r'(<w:t xml:space="preserve">\s*PARK\s+[A-ZÇĞİÖŞÜ]+\s+)(\d{4}/\d+)(</w:t>)'
    single_matches = re.findall(single_run_pattern, xml, flags=re.DOTALL)
    if len(single_matches) == 1:
        replacement = r"\g<1>" + case_value + r"\g<3>"
        return re.sub(single_run_pattern, replacement, xml, count=1, flags=re.DOTALL)
    if len(single_matches) > 1:
        raise ValueError("Şablonda şirket esas numarası için birden fazla aday bulundu.")

    # İki run formatı: PARK A | 2025/357
    two_run_pattern = (
        r'(<w:t xml:space="preserve">\s*PARK\s+[A-ZÇĞİÖŞÜ]+\s*</w:t>\s*</w:r>\s*<w:r(?:\s[^>]*)?>\s*'
        r'<w:rPr>.*?</w:rPr>\s*<w:t xml:space="preserve">)(\d{4}/\d+)(</w:t>)'
    )
    two_run_matches = re.findall(two_run_pattern, xml, flags=re.DOTALL)
    if len(two_run_matches) == 1:
        replacement = r"\g<1>" + case_value + r"\g<3>"
        return re.sub(two_run_pattern, replacement, xml, count=1, flags=re.DOTALL)
    if len(two_run_matches) > 1:
        raise ValueError("Şablonda şirket esas numarası için birden fazla aday bulundu.")

    # Parçalı format: PARK H 2026/ | 23
    split_run_pattern = (
        r'(<w:t xml:space="preserve">\s*PARK\s+[A-ZÇĞİÖŞÜ]+\s+)(\d{4}/)(</w:t>\s*</w:r>\s*<w:r(?:\s[^>]*)?>\s*'
        r'<w:rPr>.*?</w:rPr>\s*<w:t xml:space="preserve">)(\d+)(</w:t>)'
    )
    split_matches = re.findall(split_run_pattern, xml, flags=re.DOTALL)
    if len(split_matches) != 1:
        raise ValueError(
            "Şablonda şirket esas numarası için beklenen alan 1 kez bulunmalıydı, "
            f"bulunan adet: {len(split_matches)}"
        )

    year_part, no_part = case_value.split("/", 1)
    replacement = r"\g<1>" + f"{year_part}/" + r"\g<3>" + no_part + r"\g<5>"
    return re.sub(split_run_pattern, replacement, xml, count=1, flags=re.DOTALL)


def replace_ilgi_date(xml: str, new_date: str) -> str:
    pattern = (
        r'(<w:p[^>]*>.*?<w:t xml:space="preserve">İlgi</w:t>.*?)(\d{2}/\d{2}/\d{4})'
        r'(.*?Esas sayılı yazınız.*?</w:p>)'
    )
    replacement = r"\g<1>" + new_date + r"\g<3>"
    return replace_by_pattern_once(xml, pattern, replacement, "Mahkeme Gönderim Tarihi")


def replace_court_case_number(xml: str, case_value: str) -> str:
    pattern = (
        r'(<w:t xml:space="preserve">\s*tarihli ve\s*</w:t>\s*</w:r>\s*<w:r(?:\s[^>]*)?>\s*'
        r'<w:rPr>.*?</w:rPr>\s*<w:t xml:space="preserve">)(\d{4}/\d+)(</w:t>)'
    )
    replacement = r"\g<1>" + case_value + r"\g<3>"
    return replace_by_pattern_once(xml, pattern, replacement, "Mahkeme Esas Numarası")


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


def replace_court_number(xml: str, court_no: str) -> str:
    pattern = (
        r'(<w:t xml:space="preserve">)(\d+\.)(</w:t>\s*</w:r>\s*<w:r(?:\s[^>]*)?>\s*'
        r'<w:rPr>.*?</w:rPr>\s*<w:t xml:space="preserve">\s*İŞ MAHKEMESİ\s*</w:t>)'
    )
    matches = re.findall(pattern, xml, flags=re.DOTALL)
    if len(matches) != 1:
        raise ValueError(
            "Şablonda 'X.İŞ MAHKEMESİ' numarası için beklenen alan 1 kez bulunmalıydı, "
            f"bulunan adet: {len(matches)}"
        )
    replacement = r"\g<1>" + f"{court_no}." + r"\g<3>"
    return re.sub(pattern, replacement, xml, count=1, flags=re.DOTALL)


def move_park_teknik_to_top(xml: str) -> str:
    body_pattern = r"(<w:body>)(.*?)(</w:body>)"
    body_match = re.search(body_pattern, xml, flags=re.DOTALL)
    if body_match is None:
        raise ValueError("Şablonda <w:body> bloğu bulunamadı.")

    body_content = body_match.group(2)
    parts = re.split(r"(<w:p(?:\s[^>]*)?>.*?</w:p>)", body_content, flags=re.DOTALL)
    paragraphs = parts[1::2]
    if not paragraphs:
        raise ValueError("Şablonda paragraf yapısı bulunamadı.")

    park_marker = '<w:t xml:space="preserve">PARK TEKNİK</w:t>'
    park_indices = [i for i, p in enumerate(paragraphs) if park_marker in p]
    if len(park_indices) != 1:
        raise ValueError(
            "Şablonda 'PARK TEKNİK' paragrafı tam 1 adet olmalıydı, "
            f"bulunan adet: {len(park_indices)}"
        )

    sayi_indices = [
        i for i, p in enumerate(paragraphs) if '<w:t xml:space="preserve">Sayı</w:t>' in p or ">Sayı<" in p
    ]
    if not sayi_indices:
        raise ValueError("Şablonda 'Sayı' paragrafı bulunamadı.")

    park_idx = park_indices[0]
    sayi_idx = sayi_indices[0]
    if park_idx == sayi_idx:
        return xml

    park_paragraph = paragraphs.pop(park_idx)
    insert_at = sayi_idx if park_idx > sayi_idx else max(0, sayi_idx - 1)
    paragraphs.insert(insert_at, park_paragraph)

    rebuilt_parts: list[str] = []
    para_iter = iter(paragraphs)
    for i, part in enumerate(parts):
        if i % 2 == 1:
            rebuilt_parts.append(next(para_iter))
        else:
            rebuilt_parts.append(part)

    new_body_content = "".join(rebuilt_parts)
    return xml[: body_match.start(2)] + new_body_content + xml[body_match.end(2) :]


def update_docx_document_xml(
    input_docx: Path,
    output_docx: Path,
    ad_soyad: str,
    tc: str,
    sirket_esas_no: str,
    mahkeme_esas_no: str,
    mahkeme_tarihi: str,
    evrak_tarihi: str,
    court_no: str,
) -> None:
    current_year = dt.datetime.now().year
    sirket_esas_value = f"{current_year}/{sirket_esas_no}"
    mahkeme_esas_value = f"{current_year}/{mahkeme_esas_no}"

    with zipfile.ZipFile(input_docx, "r") as zin, zipfile.ZipFile(output_docx, "w") as zout:
        has_document_xml = False

        for item in zin.infolist():
            data = zin.read(item.filename)

            if item.filename == DOCX_DOCUMENT_XML:
                has_document_xml = True
                xml = data.decode("utf-8")
                xml = replace_exact_once(xml, "Doğukan yurt", ad_soyad, "Ad Soyad")
                xml = replace_exact_once(xml, "16291090514", tc, "TC")
                xml = replace_company_case_number(xml, sirket_esas_value)
                xml = replace_court_case_number(xml, mahkeme_esas_value)
                xml = replace_ilgi_date(xml, mahkeme_tarihi)
                xml = replace_header_date(xml, evrak_tarihi)
                xml = replace_court_number(xml, court_no)
                xml = tighten_header_date_position(xml)
                xml = normalize_justified_paragraphs(xml)
                xml = move_park_teknik_to_top(xml)
                data = xml.encode("utf-8")

            zout.writestr(item, data)

    if not has_document_xml:
        raise RuntimeError("DOCX içinde word/document.xml bulunamadı.")


def validate_yes_no(value: str) -> bool:
    accepted = {"e", "evet", "h", "hayir", "hayır", "yes", "no", "y", "n"}
    if value.strip().lower() in accepted:
        return True
    print("Hata: Lütfen 'evet/e' ya da 'hayır/h' giriniz.")
    return False


def is_yes(value: str) -> bool:
    return value.strip().lower() in {"e", "evet", "yes", "y"}


def process_template(
    template_path: Path,
    output_doc: Path,
    ad_soyad: str,
    tc: str,
    sirket_esas_no: str,
    mahkeme_esas_no: str,
    mahkeme_tarihi: str,
    evrak_tarihi: str,
    court_no: str,
) -> None:
    with tempfile.TemporaryDirectory(prefix="muzekkere_") as tmp:
        tmpdir = Path(tmp)
        converted_docx = tmpdir / "template.docx"
        updated_docx = tmpdir / "updated.docx"

        run_command(["textutil", "-convert", "docx", "-output", str(converted_docx), str(template_path)])
        update_docx_document_xml(
            input_docx=converted_docx,
            output_docx=updated_docx,
            ad_soyad=ad_soyad,
            tc=tc,
            sirket_esas_no=sirket_esas_no,
            mahkeme_esas_no=mahkeme_esas_no,
            mahkeme_tarihi=mahkeme_tarihi,
            evrak_tarihi=evrak_tarihi,
            court_no=court_no,
        )
        run_command(["textutil", "-convert", "doc", "-output", str(output_doc), str(updated_docx)])


def sanitize_filename_component(value: str) -> str:
    cleaned = value.strip().lower()
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = cleaned.replace("/", "-")
    cleaned = re.sub(r'[<>:"\\|?*\x00-\x1F]', "", cleaned)
    return cleaned.strip(" .")


def ensure_unique_output_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    i = 2
    while True:
        candidate = parent / f"{stem} ({i}){suffix}"
        if not candidate.exists():
            return candidate
        i += 1


def build_output_base_name(court_no: str, mahkeme_esas_no: str, ad_soyad: str) -> str:
    year = dt.datetime.now().year
    safe_name = sanitize_filename_component(ad_soyad) or "isimsiz"
    # Dosya adında "/" kullanılamadığı için esas no bölümünde "-" kullanıyoruz.
    return f"{court_no}. İş {year}-{mahkeme_esas_no} {safe_name}"


def build_output_path(base_doc_path: Path, base_name: str, kind_suffix: str = "") -> Path:
    file_name = base_name if not kind_suffix else f"{base_name} {kind_suffix}"
    raw_path = base_doc_path.with_name(f"{file_name}.DOC")
    return ensure_unique_output_path(raw_path)


def main() -> int:
    if shutil.which("textutil") is None:
        print("Hata: 'textutil' komutu bulunamadı. Bu script macOS üzerinde çalışmak üzere tasarlandı.")
        return 1

    if not TEMPLATE_PATH.exists():
        print(f"Hata: Şablon dosya bulunamadı: {TEMPLATE_PATH}")
        return 1
    if not PUANTAJ_TEMPLATE_PATH.exists():
        print(f"Hata: Şablon dosya bulunamadı: {PUANTAJ_TEMPLATE_PATH}")
        return 1

    ad_soyad = prompt_until_valid("1) Ad Soyad: ", lambda x: len(x) > 0)
    tc = prompt_until_valid("2) TC Kimlik No: ", lambda x: validate_nonempty_digits(x, "TC Kimlik No"))
    sirket_esas_no = prompt_until_valid(
        "3) Kendi Şirketimin Esas Numarası (sadece sayı): ",
        lambda x: validate_nonempty_digits(x, "Kendi Şirketimin Esas Numarası"),
    )
    mahkeme_esas_no = prompt_until_valid(
        "4) Mahkeme Esas Numarası (sadece sayı): ",
        lambda x: validate_nonempty_digits(x, "Mahkeme Esas Numarası"),
    )
    mahkeme_tarihi = prompt_until_valid(
        "5) Mahkemenin Gönderdiği Tarih (GG/AA/YYYY): ", validate_date_ddmmyyyy
    )
    evrak_tarihi = prompt_until_valid("6) Evrakın Hazırlandığı Tarih (GG/AA/YYYY): ", validate_date_ddmmyyyy)
    court_no = prompt_until_valid(
        "7) Kaçıncı İş Mahkemesi?: ", lambda x: validate_nonempty_digits(x, "Kaçıncı İş Mahkemesi")
    )
    produce_puantaj = prompt_until_valid(
        "8) Puantaj belgesi de aynı bilgilerle oluşturulsun mu? (evet/hayır): ",
        validate_yes_no,
    )

    base_name = build_output_base_name(court_no, mahkeme_esas_no, ad_soyad)
    output_doc = build_output_path(TEMPLATE_PATH, base_name)

    try:
        process_template(
            template_path=TEMPLATE_PATH,
            output_doc=output_doc,
            ad_soyad=ad_soyad,
            tc=tc,
            sirket_esas_no=sirket_esas_no,
            mahkeme_esas_no=mahkeme_esas_no,
            mahkeme_tarihi=mahkeme_tarihi,
            evrak_tarihi=evrak_tarihi,
            court_no=court_no,
        )

        output_paths = [output_doc]
        if is_yes(produce_puantaj):
            puantaj_output = build_output_path(PUANTAJ_TEMPLATE_PATH, base_name, "puantaj")
            process_template(
                template_path=PUANTAJ_TEMPLATE_PATH,
                output_doc=puantaj_output,
                ad_soyad=ad_soyad,
                tc=tc,
                sirket_esas_no=sirket_esas_no,
                mahkeme_esas_no=mahkeme_esas_no,
                mahkeme_tarihi=mahkeme_tarihi,
                evrak_tarihi=evrak_tarihi,
                court_no=court_no,
            )
            output_paths.append(puantaj_output)
    except Exception as exc:
        print(f"Hata: {exc}")
        return 1

    print("Başarılı: Yeni dosya(lar) oluşturuldu:")
    for path in output_paths:
        print(f"- {path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
