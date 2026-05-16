# -*- coding: utf-8 -*-
"""
Gerador de certificados com interface grafica.

Execute com:
    python3 certificados_gui.py

O app foi pensado para usuarios sem contato com programacao:
- seleciona a planilha Excel/CSV;
- seleciona o modelo do certificado;
- escolhe as colunas de nome, e-mail e matricula/CPF;
- gera uma previa;
- envia os certificados por e-mail usando a senha de app/key informada na tela.
"""

from __future__ import annotations

import csv
import io
import os
import posixpath
import queue
import re
import smtplib
import subprocess
import sys
import tempfile
import threading
import unicodedata
import zipfile
from dataclasses import dataclass
from email.message import EmailMessage
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ModuleNotFoundError:
    tk = None
    ttk = None
    filedialog = None
    messagebox = None

try:
    from PIL import Image, ImageDraw, ImageFont
except ModuleNotFoundError:
    Image = None
    ImageDraw = None
    ImageFont = None

try:
    from pypdf import PdfReader, PdfWriter
except ModuleNotFoundError:
    PdfReader = None
    PdfWriter = None

try:
    from reportlab.lib.colors import HexColor
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfgen import canvas as pdf_canvas
except ModuleNotFoundError:
    HexColor = None
    pdfmetrics = None
    pdf_canvas = None


APP_TITLE = "Gerador de Certificados"
DEFAULT_OUTPUT_FOLDER = "certificados_gerados"
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg"}
PDF_EXTENSIONS = {".pdf"}
SPREADSHEET_EXTENSIONS = {".xlsx", ".csv"}
BLOCKED_STATUS = {
    "PENDENTE",
    "REPROVADO",
    "REPROVADA",
    "CANCELADO",
    "CANCELADA",
    "AUSENTE",
}

DEFAULT_CERT_TEXT = (
    "Certificamos que <nome>, identificado(a) por <matricula ou cpf>, "
    "participou do evento."
)
DEFAULT_EMAIL_SUBJECT = "Seu certificado"
DEFAULT_EMAIL_BODY = (
    "Ola, <nome>!\n\n"
    "Segue em anexo o seu certificado.\n\n"
    "Atenciosamente,\n"
    "Equipe organizadora."
)

COLUMN_ALIASES = {
    "name": [
        "nome",
        "nome completo",
        "participante",
        "aluno",
        "inscrito",
        "cliente",
    ],
    "email": [
        "e-mail",
        "email",
        "mail",
        "correio",
        "email destinatario",
        "e-mail destinatario",
    ],
    "identifier": [
        "documento",
        "cpf",
        "matricula",
        "matrícula",
        "rg",
        "id",
        "identificador",
        "numero de inscricao",
        "número de inscrição",
        "codigo",
        "código",
    ],
    "status": [
        "status",
        "situacao",
        "situação",
        "inscricao",
        "inscrição",
    ],
}

FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation2/LiberationSerif-Regular.ttf",
    "/usr/share/fonts/truetype/msttcorefonts/Arial.ttf",
    "C:/Windows/Fonts/arial.ttf",
    "C:/Windows/Fonts/times.ttf",
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "/Library/Fonts/Arial.ttf",
]


@dataclass
class Participant:
    row_number: int
    name: str
    email: str
    identifier: str
    status: str


@dataclass
class ProcessingStats:
    total_rows: int = 0
    valid_rows: int = 0
    skipped_empty: int = 0
    skipped_pending: int = 0
    skipped_missing_name: int = 0
    skipped_missing_identifier: int = 0
    skipped_missing_email: int = 0
    skipped_invalid_email: int = 0


@dataclass
class ProcessSettings:
    spreadsheet_path: str
    template_path: str
    output_folder: str
    name_column: str
    email_column: str
    identifier_column: str
    status_column: str
    certificate_text: str
    text_x_percent: float
    text_y_percent: float
    text_width_percent: float
    font_size: int
    text_color: str
    text_align: str
    sender_email: str
    sender_key: str
    smtp_host: str
    smtp_port: int
    email_subject: str
    email_body: str
    dry_run: bool
    skip_blocked_status: bool
    preview_only: bool = False


def normalize_text(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = " ".join(text.split())
    return text.upper().strip()


def normalize_key(value: object) -> str:
    return "".join(char for char in normalize_text(value).lower() if char.isalnum())


def safe_filename(value: object) -> str:
    name = "_".join(str(value or "Participante").split())
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    return name[:120] or "Participante"


def email_is_valid(email: object) -> bool:
    email_text = str(email or "").strip()
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email_text))


def unique_path(folder: str | Path, filename: str) -> Path:
    folder_path = Path(folder)
    destination = folder_path / filename
    if not destination.exists():
        return destination

    stem = destination.stem
    suffix = destination.suffix
    counter = 2
    while True:
        candidate = folder_path / f"{stem}_{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def column_value(row: dict[str, str], column: str) -> str:
    if not column or column == "(sem coluna)":
        return ""
    return str(row.get(column, "") or "").strip()


def find_column(headers: list[str], aliases: list[str]) -> str:
    normalized = {normalize_key(header): header for header in headers}
    for alias in aliases:
        found = normalized.get(normalize_key(alias))
        if found:
            return found
    return ""


def parse_float(value: object, default: float) -> float:
    try:
        return float(str(value).replace(",", "."))
    except (TypeError, ValueError):
        return default


def parse_int(value: object, default: int) -> int:
    try:
        return int(float(str(value).replace(",", ".")))
    except (TypeError, ValueError):
        return default


def parse_hex_color(value: str) -> tuple[int, int, int]:
    text = (value or "#111111").strip()
    if not text.startswith("#"):
        text = f"#{text}"
    if not re.match(r"^#[0-9a-fA-F]{6}$", text):
        text = "#111111"
    return tuple(int(text[index : index + 2], 16) for index in (1, 3, 5))


def _column_index(cell_reference: str) -> int:
    letters = "".join(char for char in cell_reference if char.isalpha())
    index = 0
    for letter in letters:
        index = (index * 26) + (ord(letter.upper()) - ord("A") + 1)
    return index - 1


def _discover_sheet_path(spreadsheet_zip: zipfile.ZipFile) -> str:
    namespaces = {
        "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    }

    workbook = ET.fromstring(spreadsheet_zip.read("xl/workbook.xml"))
    first_sheet = workbook.find("a:sheets/a:sheet", namespaces)
    if first_sheet is None:
        return "xl/worksheets/sheet1.xml"

    rel_id = first_sheet.get(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    )
    if not rel_id:
        return "xl/worksheets/sheet1.xml"

    rels = ET.fromstring(spreadsheet_zip.read("xl/_rels/workbook.xml.rels"))
    for rel in rels.findall("pr:Relationship", namespaces):
        if rel.get("Id") != rel_id:
            continue
        target = (rel.get("Target") or "").lstrip("/")
        if not target.startswith("xl/"):
            target = posixpath.join("xl", target)
        return posixpath.normpath(target)

    return "xl/worksheets/sheet1.xml"


def load_xlsx(path: str | Path) -> list[dict[str, str]]:
    namespaces = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

    with zipfile.ZipFile(path) as spreadsheet_zip:
        shared_strings = []
        if "xl/sharedStrings.xml" in spreadsheet_zip.namelist():
            root = ET.fromstring(spreadsheet_zip.read("xl/sharedStrings.xml"))
            for item in root.findall("a:si", namespaces):
                text = "".join(part.text or "" for part in item.findall(".//a:t", namespaces))
                shared_strings.append(text)

        sheet_path = _discover_sheet_path(spreadsheet_zip)
        sheet_root = ET.fromstring(spreadsheet_zip.read(sheet_path))
        rows = []

        for row in sheet_root.findall("a:sheetData/a:row", namespaces):
            columns = {}
            for cell in row.findall("a:c", namespaces):
                reference = cell.get("r", "")
                if not reference:
                    continue

                cell_type = cell.get("t")
                if cell_type == "inlineStr":
                    value = "".join(part.text or "" for part in cell.findall(".//a:t", namespaces))
                else:
                    value_element = cell.find("a:v", namespaces)
                    value = value_element.text if value_element is not None else ""
                    if cell_type == "s" and value:
                        value = shared_strings[int(value)]

                columns[_column_index(reference)] = value

            if columns:
                last_index = max(columns)
                rows.append([columns.get(index, "") for index in range(last_index + 1)])

    if not rows:
        return []

    headers = [str(header).strip() for header in rows[0]]
    records = []
    for values in rows[1:]:
        record = {}
        for index, header in enumerate(headers):
            if not header:
                continue
            record[header] = str(values[index]).strip() if index < len(values) else ""
        if any(value.strip() for value in record.values()):
            records.append(record)
    return records


def load_csv(path: str | Path) -> list[dict[str, str]]:
    raw = Path(path).read_bytes()
    last_error = None

    for encoding in ("utf-8-sig", "utf-8", "latin1"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError as error:
            last_error = error
    else:
        raise ValueError(f"Nao foi possivel ler o CSV: {last_error}")

    sample = text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel

    reader = csv.DictReader(io.StringIO(text), dialect=dialect)
    return [
        {str(key or "").strip(): str(value or "").strip() for key, value in row.items()}
        for row in reader
        if any(str(value or "").strip() for value in row.values())
    ]


def load_table(path: str | Path) -> list[dict[str, str]]:
    suffix = Path(path).suffix.lower()
    if suffix == ".xlsx":
        return load_xlsx(path)
    if suffix == ".csv":
        return load_csv(path)
    raise ValueError("Use uma planilha .xlsx ou .csv.")


def table_headers(records: list[dict[str, str]]) -> list[str]:
    headers = []
    seen = set()
    for record in records:
        for header in record:
            if header not in seen:
                seen.add(header)
                headers.append(header)
    return headers


def replace_placeholders(template: str, participant: Participant) -> str:
    replacements = [
        (["nome", "name", "participante"], participant.name),
        (["email", "e-mail"], participant.email),
        (
            [
                "matricula",
                "matrícula",
                "cpf",
                "documento",
                "matricula ou cpf",
                "matrícula ou cpf",
                "matricula/cpf",
                "matrícula/cpf",
                "id",
            ],
            participant.identifier,
        ),
    ]

    result = template
    for aliases, value in replacements:
        for alias in aliases:
            escaped = re.escape(alias).replace(r"\ ", r"\s+")
            patterns = [
                rf"<\s*{escaped}\s*>",
                rf"\{{\s*{escaped}\s*\}}",
                rf"\[\s*{escaped}\s*\]",
            ]
            for pattern in patterns:
                result = re.sub(pattern, value, result, flags=re.IGNORECASE)
    return result


def contains_name_placeholder(text: str) -> bool:
    return bool(re.search(r"<\s*nome\s*>|\{\s*nome\s*\}|\[\s*nome\s*\]", text, re.I))


def contains_identifier_placeholder(text: str) -> bool:
    aliases = [
        "matricula",
        "matrícula",
        "cpf",
        "documento",
        "matricula ou cpf",
        "matrícula ou cpf",
        "matricula/cpf",
        "matrícula/cpf",
    ]
    for alias in aliases:
        escaped = re.escape(alias).replace(r"\ ", r"\s+")
        patterns = [
            rf"<\s*{escaped}\s*>",
            rf"\{{\s*{escaped}\s*\}}",
            rf"\[\s*{escaped}\s*\]",
        ]
        if any(re.search(pattern, text, re.I) for pattern in patterns):
            return True
    return False


def prepare_participants(
    records: list[dict[str, str]],
    name_column: str,
    email_column: str,
    identifier_column: str,
    status_column: str,
    skip_blocked_status: bool,
) -> tuple[list[Participant], ProcessingStats]:
    stats = ProcessingStats(total_rows=len(records))
    participants = []

    for index, row in enumerate(records, start=2):
        if not any(str(value or "").strip() for value in row.values()):
            stats.skipped_empty += 1
            continue

        name = column_value(row, name_column)
        email = column_value(row, email_column)
        identifier = column_value(row, identifier_column)
        status = column_value(row, status_column)

        if skip_blocked_status and status and normalize_text(status) in BLOCKED_STATUS:
            stats.skipped_pending += 1
            continue
        if not name:
            stats.skipped_missing_name += 1
            continue
        if not identifier:
            stats.skipped_missing_identifier += 1
            continue
        if not email:
            stats.skipped_missing_email += 1
            continue
        if not email_is_valid(email):
            stats.skipped_invalid_email += 1
            continue

        participants.append(
            Participant(
                row_number=index,
                name=" ".join(name.split()),
                email=email,
                identifier=identifier,
                status=status,
            )
        )

    stats.valid_rows = len(participants)
    return participants, stats


def find_font(font_size: int):
    if ImageFont is None:
        return None
    for candidate in FONT_CANDIDATES:
        if os.path.exists(candidate):
            return ImageFont.truetype(candidate, font_size)
    return ImageFont.load_default()


def measure_pil_text(draw, text: str, font) -> int:
    bbox = draw.textbbox((0, 0), text, font=font)
    return int(bbox[2] - bbox[0])


def wrap_text_pil(draw, text: str, font, max_width: int) -> list[str]:
    wrapped_lines = []
    for paragraph in text.splitlines():
        if not paragraph.strip():
            wrapped_lines.append("")
            continue

        current = ""
        for word in paragraph.split():
            candidate = f"{current} {word}".strip()
            if not current or measure_pil_text(draw, candidate, font) <= max_width:
                current = candidate
                continue

            wrapped_lines.append(current)
            current = word

        if current:
            wrapped_lines.append(current)
    return wrapped_lines


def draw_text_on_image(
    image,
    text: str,
    x_percent: float,
    y_percent: float,
    width_percent: float,
    font_size: int,
    color: str,
    align: str,
) -> None:
    draw = ImageDraw.Draw(image)
    font = find_font(font_size)
    image_width, image_height = image.size
    x = int(image_width * x_percent / 100)
    y = int(image_height * y_percent / 100)
    max_width = max(1, int(image_width * width_percent / 100))
    if x + max_width > image_width:
        max_width = max(1, image_width - x)

    lines = wrap_text_pil(draw, text, font, max_width)
    rgb_color = parse_hex_color(color)
    line_height = int(font_size * 1.35)

    for line in lines:
        line_width = measure_pil_text(draw, line, font)
        line_x = x
        if align == "centralizado":
            line_x = x + max(0, (max_width - line_width) // 2)
        elif align == "direita":
            line_x = x + max(0, max_width - line_width)

        draw.text((line_x, y), line, fill=rgb_color, font=font)
        y += line_height


def generate_from_image_template(template_path: str, output_path: str | Path, text: str, settings) -> None:
    if Image is None:
        raise RuntimeError("Para usar modelo em imagem, instale Pillow: pip install pillow")

    with Image.open(template_path) as original:
        image = original.convert("RGB")
        draw_text_on_image(
            image,
            text,
            settings.text_x_percent,
            settings.text_y_percent,
            settings.text_width_percent,
            settings.font_size,
            settings.text_color,
            settings.text_align,
        )
        image.save(output_path, "PDF", resolution=100.0)


def wrap_text_pdf(text: str, font_name: str, font_size: int, max_width: float) -> list[str]:
    lines = []
    for paragraph in text.splitlines():
        if not paragraph.strip():
            lines.append("")
            continue

        current = ""
        for word in paragraph.split():
            candidate = f"{current} {word}".strip()
            width = pdfmetrics.stringWidth(candidate, font_name, font_size)
            if not current or width <= max_width:
                current = candidate
                continue
            lines.append(current)
            current = word

        if current:
            lines.append(current)
    return lines


def generate_from_pdf_template(template_path: str, output_path: str | Path, text: str, settings) -> None:
    if PdfReader is None or PdfWriter is None or pdf_canvas is None or pdfmetrics is None:
        raise RuntimeError(
            "Para usar modelo em PDF, instale as bibliotecas: pip install pypdf reportlab"
        )

    reader = PdfReader(template_path)
    if not reader.pages:
        raise RuntimeError("O PDF modelo nao possui paginas.")

    first_page = reader.pages[0]
    page_width = float(first_page.mediabox.width)
    page_height = float(first_page.mediabox.height)
    x = page_width * settings.text_x_percent / 100
    y_top = page_height - (page_height * settings.text_y_percent / 100)
    max_width = page_width * settings.text_width_percent / 100
    if x + max_width > page_width:
        max_width = max(1, page_width - x)

    packet = io.BytesIO()
    canvas = pdf_canvas.Canvas(packet, pagesize=(page_width, page_height))
    font_name = "Helvetica"
    canvas.setFont(font_name, settings.font_size)
    red, green, blue = parse_hex_color(settings.text_color)
    canvas.setFillColorRGB(red / 255, green / 255, blue / 255)

    line_height = settings.font_size * 1.35
    y = y_top
    for line in wrap_text_pdf(text, font_name, settings.font_size, max_width):
        line_width = pdfmetrics.stringWidth(line, font_name, settings.font_size)
        line_x = x
        if settings.text_align == "centralizado":
            line_x = x + max(0, (max_width - line_width) / 2)
        elif settings.text_align == "direita":
            line_x = x + max(0, max_width - line_width)
        canvas.drawString(line_x, y, line)
        y -= line_height

    canvas.save()
    packet.seek(0)
    overlay = PdfReader(packet).pages[0]
    first_page.merge_page(overlay)

    writer = PdfWriter()
    writer.add_page(first_page)
    for page in reader.pages[1:]:
        writer.add_page(page)

    with open(output_path, "wb") as output_file:
        writer.write(output_file)


def generate_certificate(
    template_path: str,
    output_path: str | Path,
    certificate_text: str,
    settings: ProcessSettings,
) -> None:
    suffix = Path(template_path).suffix.lower()
    if suffix in IMAGE_EXTENSIONS:
        generate_from_image_template(template_path, output_path, certificate_text, settings)
        return
    if suffix in PDF_EXTENSIONS:
        generate_from_pdf_template(template_path, output_path, certificate_text, settings)
        return
    raise ValueError("O modelo deve ser .png, .jpg, .jpeg ou .pdf.")


def build_email_message(
    sender: str,
    recipient: str,
    subject: str,
    body: str,
    attachment_path: str | Path,
) -> EmailMessage:
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = recipient
    msg.set_content(body)

    with open(attachment_path, "rb") as file:
        msg.add_attachment(
            file.read(),
            maintype="application",
            subtype="pdf",
            filename=Path(attachment_path).name,
        )
    return msg


def connect_smtp(settings: ProcessSettings):
    password = settings.sender_key.strip()
    if "gmail" in settings.smtp_host.lower():
        password = password.replace(" ", "")

    smtp = smtplib.SMTP_SSL(settings.smtp_host, settings.smtp_port, timeout=30)
    smtp.login(settings.sender_email.strip(), password)
    return smtp


def open_folder(path: str | Path) -> None:
    folder = str(Path(path).resolve())
    try:
        if sys.platform.startswith("win"):
            os.startfile(folder)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", folder])
        else:
            subprocess.Popen(["xdg-open", folder])
    except Exception:
        pass


def process_batch(settings: ProcessSettings, notify) -> dict[str, int]:
    records = load_table(settings.spreadsheet_path)
    participants, stats = prepare_participants(
        records,
        settings.name_column,
        settings.email_column,
        settings.identifier_column,
        settings.status_column,
        settings.skip_blocked_status,
    )

    notify("log", f"Linhas lidas: {stats.total_rows}")
    notify("log", f"Participantes validos: {stats.valid_rows}")
    if stats.skipped_pending:
        notify("log", f"Ignorados por status: {stats.skipped_pending}")
    if stats.skipped_missing_name:
        notify("log", f"Ignorados sem nome: {stats.skipped_missing_name}")
    if stats.skipped_missing_identifier:
        notify("log", f"Ignorados sem matricula/CPF: {stats.skipped_missing_identifier}")
    if stats.skipped_missing_email:
        notify("log", f"Ignorados sem e-mail: {stats.skipped_missing_email}")
    if stats.skipped_invalid_email:
        notify("log", f"Ignorados com e-mail invalido: {stats.skipped_invalid_email}")

    if not participants:
        raise RuntimeError("Nenhum participante valido foi encontrado.")

    if settings.preview_only:
        participants = participants[:1]
        notify("log", "Modo previa: usando apenas a primeira linha valida.")

    output_folder = Path(settings.output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)

    smtp = None
    if not settings.dry_run and not settings.preview_only:
        notify("log", "Conectando ao servidor de e-mail...")
        smtp = connect_smtp(settings)
        notify("log", "Login realizado com sucesso.")

    summary = {"generated": 0, "sent": 0, "failed": 0}
    total = len(participants)

    try:
        for index, participant in enumerate(participants, start=1):
            notify("progress", (index - 1, total))
            certificate_text = replace_placeholders(settings.certificate_text, participant)
            prefix = "PREVIA_" if settings.preview_only else ""
            filename = f"{prefix}Certificado_{safe_filename(participant.name)}.pdf"
            output_path = unique_path(output_folder, filename)

            try:
                generate_certificate(settings.template_path, output_path, certificate_text, settings)
                summary["generated"] += 1
                notify("log", f"Gerado: {output_path.name}")
            except Exception as error:
                summary["failed"] += 1
                notify("log", f"Falha ao gerar para {participant.name}: {error}")
                continue

            if settings.dry_run or settings.preview_only:
                notify("log", f"Simulacao: envio nao realizado para {participant.email}")
                continue

            try:
                subject = replace_placeholders(settings.email_subject, participant)
                body = replace_placeholders(settings.email_body, participant)
                msg = build_email_message(
                    settings.sender_email.strip(),
                    participant.email,
                    subject,
                    body,
                    output_path,
                )
                smtp.send_message(msg)
                summary["sent"] += 1
                notify("log", f"Enviado: {participant.name} <{participant.email}>")
            except Exception as error:
                summary["failed"] += 1
                notify("log", f"Falha no envio para {participant.name}: {error}")

            notify("progress", (index, total))
    finally:
        if smtp is not None:
            try:
                smtp.quit()
            except Exception:
                pass

    notify("progress", (total, total))
    return summary


class CertificateApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1040x780")
        self.root.minsize(920, 680)

        self.ui_queue = queue.Queue()
        self.loaded_headers: list[str] = []

        self.spreadsheet_var = tk.StringVar()
        self.template_var = tk.StringVar()
        self.output_var = tk.StringVar(value=str(Path.cwd() / DEFAULT_OUTPUT_FOLDER))
        self.name_column_var = tk.StringVar()
        self.email_column_var = tk.StringVar()
        self.identifier_column_var = tk.StringVar()
        self.status_column_var = tk.StringVar(value="(sem coluna)")
        self.rows_label_var = tk.StringVar(value="Nenhuma planilha carregada")

        self.x_var = tk.StringVar(value="10")
        self.y_var = tk.StringVar(value="45")
        self.width_var = tk.StringVar(value="80")
        self.font_size_var = tk.StringVar(value="34")
        self.color_var = tk.StringVar(value="#111111")
        self.align_var = tk.StringVar(value="centralizado")

        self.sender_var = tk.StringVar()
        self.key_var = tk.StringVar()
        self.custom_smtp_var = tk.BooleanVar(value=False)
        self.smtp_host_var = tk.StringVar(value="smtp.gmail.com")
        self.smtp_port_var = tk.StringVar(value="465")
        self.subject_var = tk.StringVar(value=DEFAULT_EMAIL_SUBJECT)
        self.dry_run_var = tk.BooleanVar(value=True)
        self.skip_blocked_status_var = tk.BooleanVar(value=True)

        self._build_ui()
        self.root.after(100, self._poll_queue)

    def _build_ui(self):
        outer = ttk.Frame(self.root, padding=12)
        outer.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(2, weight=1)

        title = ttk.Label(outer, text=APP_TITLE, font=("Arial", 18, "bold"))
        title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        self._build_files_frame(outer).grid(row=1, column=0, sticky="nsew", padx=(0, 8))
        self._build_position_frame(outer).grid(row=1, column=1, sticky="nsew", padx=(8, 0))
        self._build_certificate_text_frame(outer).grid(
            row=2, column=0, sticky="nsew", padx=(0, 8), pady=(10, 0)
        )
        self._build_email_frame(outer).grid(
            row=2, column=1, sticky="nsew", padx=(8, 0), pady=(10, 0)
        )
        self._build_run_frame(outer).grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(10, 0))

    def _build_files_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="1. Arquivos e colunas", padding=10)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Planilha Excel/CSV").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Entry(frame, textvariable=self.spreadsheet_var).grid(row=0, column=1, sticky="ew", pady=3)
        ttk.Button(frame, text="Selecionar", command=self._browse_spreadsheet).grid(
            row=0, column=2, padx=(6, 0), pady=3
        )

        ttk.Label(frame, text="Modelo certificado").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(frame, textvariable=self.template_var).grid(row=1, column=1, sticky="ew", pady=3)
        ttk.Button(frame, text="Selecionar", command=self._browse_template).grid(
            row=1, column=2, padx=(6, 0), pady=3
        )

        ttk.Label(frame, text="Pasta de saida").grid(row=2, column=0, sticky="w", pady=3)
        ttk.Entry(frame, textvariable=self.output_var).grid(row=2, column=1, sticky="ew", pady=3)
        ttk.Button(frame, text="Selecionar", command=self._browse_output).grid(
            row=2, column=2, padx=(6, 0), pady=3
        )

        ttk.Label(frame, textvariable=self.rows_label_var).grid(
            row=3, column=0, columnspan=3, sticky="w", pady=(8, 4)
        )

        self.name_combo = self._add_column_combo(frame, "Coluna do nome", self.name_column_var, 4)
        self.identifier_combo = self._add_column_combo(
            frame, "Coluna matricula/CPF", self.identifier_column_var, 5
        )
        self.email_combo = self._add_column_combo(frame, "Coluna do e-mail", self.email_column_var, 6)
        self.status_combo = self._add_column_combo(
            frame, "Coluna de status", self.status_column_var, 7, allow_empty=True
        )

        ttk.Checkbutton(
            frame,
            text="Pular Pendente/Reprovado/Cancelado",
            variable=self.skip_blocked_status_var,
        ).grid(row=8, column=0, columnspan=3, sticky="w", pady=(6, 0))

        return frame

    def _build_position_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="2. Posicao do texto", padding=10)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="X (%)").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Spinbox(frame, from_=0, to=100, increment=1, textvariable=self.x_var, width=8).grid(
            row=0, column=1, sticky="w", pady=3
        )

        ttk.Label(frame, text="Y (%)").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Spinbox(frame, from_=0, to=100, increment=1, textvariable=self.y_var, width=8).grid(
            row=1, column=1, sticky="w", pady=3
        )

        ttk.Label(frame, text="Largura (%)").grid(row=2, column=0, sticky="w", pady=3)
        ttk.Spinbox(
            frame,
            from_=10,
            to=100,
            increment=1,
            textvariable=self.width_var,
            width=8,
        ).grid(row=2, column=1, sticky="w", pady=3)

        ttk.Label(frame, text="Tamanho da fonte").grid(row=3, column=0, sticky="w", pady=3)
        ttk.Spinbox(
            frame,
            from_=8,
            to=120,
            increment=1,
            textvariable=self.font_size_var,
            width=8,
        ).grid(row=3, column=1, sticky="w", pady=3)

        ttk.Label(frame, text="Cor").grid(row=4, column=0, sticky="w", pady=3)
        ttk.Entry(frame, textvariable=self.color_var, width=12).grid(row=4, column=1, sticky="w", pady=3)

        ttk.Label(frame, text="Alinhamento").grid(row=5, column=0, sticky="w", pady=3)
        ttk.Combobox(
            frame,
            textvariable=self.align_var,
            values=["centralizado", "esquerda", "direita"],
            state="readonly",
            width=16,
        ).grid(row=5, column=1, sticky="w", pady=3)

        ttk.Label(
            frame,
            text="Use Previa para ajustar a posicao antes do envio real.",
        ).grid(row=6, column=0, columnspan=2, sticky="w", pady=(8, 0))

        return frame

    def _build_certificate_text_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="3. Texto no certificado", padding=10)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        ttk.Label(frame, text="Marcadores: <nome> e <matricula ou cpf>").grid(
            row=0, column=0, sticky="w"
        )
        self.certificate_text = tk.Text(frame, height=8, wrap="word")
        self.certificate_text.insert("1.0", DEFAULT_CERT_TEXT)
        self.certificate_text.grid(row=1, column=0, sticky="nsew", pady=(6, 0))

        return frame

    def _build_email_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="4. E-mail", padding=10)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(6, weight=1)

        ttk.Label(frame, text="E-mail remetente").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Entry(frame, textvariable=self.sender_var).grid(row=0, column=1, sticky="ew", pady=3)

        ttk.Label(frame, text="Senha de app/key").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(frame, textvariable=self.key_var, show="*").grid(row=1, column=1, sticky="ew", pady=3)

        ttk.Checkbutton(
            frame,
            text="Usar servidor SMTP personalizado",
            variable=self.custom_smtp_var,
            command=self._toggle_smtp_fields,
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 2))

        self.smtp_frame = ttk.Frame(frame)
        self.smtp_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=3)
        self.smtp_frame.columnconfigure(1, weight=1)
        ttk.Label(self.smtp_frame, text="Servidor").grid(row=0, column=0, sticky="w")
        ttk.Entry(self.smtp_frame, textvariable=self.smtp_host_var).grid(
            row=0, column=1, sticky="ew", padx=(6, 6)
        )
        ttk.Label(self.smtp_frame, text="Porta").grid(row=0, column=2, sticky="w")
        ttk.Entry(self.smtp_frame, textvariable=self.smtp_port_var, width=8).grid(
            row=0, column=3, padx=(6, 0)
        )

        ttk.Label(frame, text="Assunto").grid(row=4, column=0, sticky="w", pady=3)
        ttk.Entry(frame, textvariable=self.subject_var).grid(row=4, column=1, sticky="ew", pady=3)

        ttk.Label(frame, text="Mensagem").grid(row=5, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.email_body_text = tk.Text(frame, height=7, wrap="word")
        self.email_body_text.insert("1.0", DEFAULT_EMAIL_BODY)
        self.email_body_text.grid(row=6, column=0, columnspan=2, sticky="nsew", pady=(6, 0))

        ttk.Checkbutton(
            frame,
            text="Simular envio (gera PDF, nao envia e-mail)",
            variable=self.dry_run_var,
        ).grid(row=7, column=0, columnspan=2, sticky="w", pady=(8, 0))

        self._toggle_smtp_fields()

        return frame

    def _build_run_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="5. Executar", padding=10)
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(2, weight=1)

        buttons = ttk.Frame(frame)
        buttons.grid(row=0, column=0, columnspan=2, sticky="ew")
        buttons.columnconfigure(0, weight=1)
        buttons.columnconfigure(1, weight=1)
        buttons.columnconfigure(2, weight=1)

        self.preview_button = ttk.Button(
            buttons,
            text="Gerar previa",
            command=lambda: self._start_processing(preview_only=True),
        )
        self.preview_button.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        self.run_button = ttk.Button(
            buttons,
            text="Gerar certificados / enviar",
            command=lambda: self._start_processing(preview_only=False),
        )
        self.run_button.grid(row=0, column=1, sticky="ew", padx=6)

        ttk.Button(buttons, text="Abrir pasta de saida", command=self._open_output).grid(
            row=0, column=2, sticky="ew", padx=(6, 0)
        )

        self.progress = ttk.Progressbar(frame, maximum=100)
        self.progress.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 6))

        self.log_text = tk.Text(frame, height=9, wrap="word", state="disabled")
        self.log_text.grid(row=2, column=0, columnspan=2, sticky="nsew")

        return frame

    def _add_column_combo(self, parent, label, variable, row, allow_empty=False):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=3)
        combo = ttk.Combobox(parent, textvariable=variable, state="readonly")
        combo.grid(row=row, column=1, columnspan=2, sticky="ew", pady=3)
        if allow_empty:
            combo["values"] = ["(sem coluna)"]
            combo.set("(sem coluna)")
        return combo

    def _browse_spreadsheet(self):
        path = filedialog.askopenfilename(
            title="Selecionar planilha",
            filetypes=[
                ("Planilhas", "*.xlsx *.csv"),
                ("Excel", "*.xlsx"),
                ("CSV", "*.csv"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if path:
            self.spreadsheet_var.set(path)
            self._load_spreadsheet_preview()

    def _browse_template(self):
        path = filedialog.askopenfilename(
            title="Selecionar modelo",
            filetypes=[
                ("Modelo de certificado", "*.png *.jpg *.jpeg *.pdf"),
                ("Imagem", "*.png *.jpg *.jpeg"),
                ("PDF", "*.pdf"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if path:
            self.template_var.set(path)

    def _browse_output(self):
        path = filedialog.askdirectory(title="Selecionar pasta de saida")
        if path:
            self.output_var.set(path)

    def _open_output(self):
        output = self.output_var.get().strip()
        if output:
            Path(output).mkdir(parents=True, exist_ok=True)
            open_folder(output)

    def _toggle_smtp_fields(self):
        if self.custom_smtp_var.get():
            self.smtp_frame.grid()
        else:
            self.smtp_host_var.set("smtp.gmail.com")
            self.smtp_port_var.set("465")
            self.smtp_frame.grid_remove()

    def _load_spreadsheet_preview(self):
        path = self.spreadsheet_var.get().strip()
        try:
            records = load_table(path)
            headers = table_headers(records)
        except Exception as error:
            self.rows_label_var.set("Falha ao carregar planilha")
            messagebox.showerror("Planilha", str(error))
            return

        self.loaded_headers = headers
        combo_values = headers if headers else []
        self.name_combo["values"] = combo_values
        self.identifier_combo["values"] = combo_values
        self.email_combo["values"] = combo_values
        self.status_combo["values"] = ["(sem coluna)"] + combo_values

        self.name_column_var.set(find_column(headers, COLUMN_ALIASES["name"]))
        self.identifier_column_var.set(find_column(headers, COLUMN_ALIASES["identifier"]))
        self.email_column_var.set(find_column(headers, COLUMN_ALIASES["email"]))
        self.status_column_var.set(find_column(headers, COLUMN_ALIASES["status"]) or "(sem coluna)")

        self.rows_label_var.set(f"Planilha carregada: {len(records)} linhas | {len(headers)} colunas")
        self._log("Planilha carregada e colunas detectadas.")

    def _read_settings(self, preview_only: bool) -> ProcessSettings | None:
        spreadsheet_path = self.spreadsheet_var.get().strip()
        template_path = self.template_var.get().strip()
        output_folder = self.output_var.get().strip()
        certificate_text = self.certificate_text.get("1.0", "end").strip()
        email_body = self.email_body_text.get("1.0", "end").strip()

        try:
            smtp_host = self.smtp_host_var.get().strip() if self.custom_smtp_var.get() else "smtp.gmail.com"
            smtp_port = parse_int(self.smtp_port_var.get(), 465) if self.custom_smtp_var.get() else 465
            settings = ProcessSettings(
                spreadsheet_path=spreadsheet_path,
                template_path=template_path,
                output_folder=output_folder,
                name_column=self.name_column_var.get().strip(),
                email_column=self.email_column_var.get().strip(),
                identifier_column=self.identifier_column_var.get().strip(),
                status_column=self.status_column_var.get().strip(),
                certificate_text=certificate_text,
                text_x_percent=parse_float(self.x_var.get(), 10),
                text_y_percent=parse_float(self.y_var.get(), 45),
                text_width_percent=parse_float(self.width_var.get(), 80),
                font_size=parse_int(self.font_size_var.get(), 34),
                text_color=self.color_var.get().strip() or "#111111",
                text_align=self.align_var.get().strip() or "centralizado",
                sender_email=self.sender_var.get().strip(),
                sender_key=self.key_var.get(),
                smtp_host=smtp_host or "smtp.gmail.com",
                smtp_port=smtp_port,
                email_subject=self.subject_var.get().strip() or DEFAULT_EMAIL_SUBJECT,
                email_body=email_body or DEFAULT_EMAIL_BODY,
                dry_run=bool(self.dry_run_var.get()),
                skip_blocked_status=bool(self.skip_blocked_status_var.get()),
                preview_only=preview_only,
            )
        except Exception as error:
            messagebox.showerror("Configuracao", str(error))
            return None

        error = self._validate_settings(settings)
        if error:
            messagebox.showerror("Configuracao", error)
            return None
        return settings

    def _validate_settings(self, settings: ProcessSettings) -> str:
        spreadsheet = Path(settings.spreadsheet_path)
        template = Path(settings.template_path)

        if not spreadsheet.exists():
            return "Selecione uma planilha valida."
        if spreadsheet.suffix.lower() not in SPREADSHEET_EXTENSIONS:
            return "A planilha deve ser .xlsx ou .csv."
        if not template.exists():
            return "Selecione um modelo de certificado valido."
        if template.suffix.lower() not in IMAGE_EXTENSIONS.union(PDF_EXTENSIONS):
            return "O modelo deve ser .png, .jpg, .jpeg ou .pdf."
        if not settings.output_folder:
            return "Selecione uma pasta de saida."
        if not settings.name_column:
            return "Selecione a coluna do nome."
        if not settings.identifier_column:
            return "Selecione a coluna de matricula/CPF."
        if not settings.email_column:
            return "Selecione a coluna do e-mail."
        if not contains_name_placeholder(settings.certificate_text):
            return "O texto do certificado precisa ter o marcador <nome>."
        if not contains_identifier_placeholder(settings.certificate_text):
            return "O texto do certificado precisa ter <matricula ou cpf>, <matricula> ou <cpf>."
        if settings.font_size < 8:
            return "O tamanho da fonte deve ser maior ou igual a 8."
        if not settings.dry_run and not settings.preview_only:
            if not email_is_valid(settings.sender_email):
                return "Informe um e-mail remetente valido."
            if not settings.sender_key.strip():
                return "Informe a senha de app/key do e-mail."
        return ""

    def _start_processing(self, preview_only: bool):
        settings = self._read_settings(preview_only)
        if settings is None:
            return

        if settings.template_path.lower().endswith(".pdf") and (
            PdfReader is None or PdfWriter is None or pdf_canvas is None
        ):
            messagebox.showerror(
                "Dependencias",
                "Para modelo em PDF, instale: pip install pypdf reportlab\n\n"
                "Alternativa: exporte o modelo como PNG/JPG e selecione a imagem.",
            )
            return

        if Path(settings.template_path).suffix.lower() in IMAGE_EXTENSIONS and Image is None:
            messagebox.showerror(
                "Dependencias",
                "Para modelo em imagem, instale: pip install pillow",
            )
            return

        if not settings.preview_only and not settings.dry_run:
            confirmed = messagebox.askyesno(
                "Confirmar envio",
                "O modo simulacao esta desativado.\n\n"
                "Os certificados serao gerados e enviados por e-mail agora.\n"
                "Deseja continuar?",
            )
            if not confirmed:
                return

        self._set_running(True)
        self.progress["value"] = 0
        self._log("")
        self._log("Iniciando processamento...")

        def notify(kind, payload):
            self.ui_queue.put((kind, payload))

        def worker():
            try:
                summary = process_batch(settings, notify)
                self.ui_queue.put(("done", summary))
            except smtplib.SMTPAuthenticationError:
                self.ui_queue.put(
                    (
                        "error",
                        "Falha de autenticacao. Confira o e-mail e a senha de app/key.",
                    )
                )
            except Exception as error:
                self.ui_queue.put(("error", str(error)))

        threading.Thread(target=worker, daemon=True).start()

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.ui_queue.get_nowait()
                if kind == "log":
                    self._log(payload)
                elif kind == "progress":
                    current, total = payload
                    percent = 0 if total == 0 else int(current * 100 / total)
                    self.progress["value"] = percent
                elif kind == "done":
                    self._set_running(False)
                    generated = payload.get("generated", 0)
                    sent = payload.get("sent", 0)
                    failed = payload.get("failed", 0)
                    self._log(f"Finalizado. Gerados: {generated} | Enviados: {sent} | Falhas: {failed}")
                    messagebox.showinfo(
                        "Finalizado",
                        f"Certificados gerados: {generated}\n"
                        f"E-mails enviados: {sent}\n"
                        f"Falhas: {failed}",
                    )
                elif kind == "error":
                    self._set_running(False)
                    self._log(f"Erro: {payload}")
                    messagebox.showerror("Erro", payload)
        except queue.Empty:
            pass

        self.root.after(100, self._poll_queue)

    def _set_running(self, running: bool):
        state = "disabled" if running else "normal"
        self.preview_button.configure(state=state)
        self.run_button.configure(state=state)

    def _log(self, message: str):
        self.log_text.configure(state="normal")
        if message:
            self.log_text.insert("end", f"{message}\n")
        else:
            self.log_text.delete("1.0", "end")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")


def main() -> int:
    if "--self-test" in sys.argv:
        return self_test()

    if tk is None:
        print("Tkinter nao esta instalado neste Python.")
        print("No Ubuntu/Debian, instale com: sudo apt install python3-tk")
        print("No Windows, use o instalador oficial do Python com a opcao Tcl/Tk marcada.")
        return 1

    root = tk.Tk()
    CertificateApp(root)
    root.mainloop()
    return 0


def self_test() -> int:
    print("Iniciando autoteste do Gerador de Certificados...")
    print(f"Tkinter: {'OK' if tk is not None else 'FALTA'}")
    print(f"Pillow: {'OK' if Image is not None else 'FALTA'}")
    print(f"pypdf: {'OK' if PdfReader is not None else 'FALTA'}")
    print(f"reportlab: {'OK' if pdf_canvas is not None else 'FALTA'}")

    if Image is None:
        print("Falha: Pillow nao esta disponivel.")
        return 1

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        template_path = temp_path / "modelo_teste.png"
        output_path = temp_path / "certificado_teste.pdf"
        Image.new("RGB", (1400, 900), "white").save(template_path)

        settings = ProcessSettings(
            spreadsheet_path="",
            template_path=str(template_path),
            output_folder=str(temp_path),
            name_column="Nome",
            email_column="E-mail",
            identifier_column="Documento",
            status_column="",
            certificate_text="Certificamos que <nome> - <matricula ou cpf>",
            text_x_percent=10,
            text_y_percent=45,
            text_width_percent=80,
            font_size=34,
            text_color="#111111",
            text_align="centralizado",
            sender_email="",
            sender_key="",
            smtp_host="smtp.gmail.com",
            smtp_port=465,
            email_subject="Teste",
            email_body="Ola <nome>",
            dry_run=True,
            skip_blocked_status=True,
        )

        generate_certificate(
            str(template_path),
            output_path,
            "Certificamos que Maria Silva - 123456",
            settings,
        )

        if not output_path.exists() or output_path.stat().st_size == 0:
            print("Falha: PDF de teste nao foi gerado.")
            return 1

        print(f"PDF de teste gerado com sucesso: {output_path.stat().st_size} bytes")

    print("Autoteste concluido com sucesso.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
