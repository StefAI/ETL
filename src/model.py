import os
import shutil
import subprocess
import tempfile

from docling.document_converter import DocumentConverter, PdfFormatOption, WordFormatOption
from docling.exceptions import ConversionError
from docling.datamodel.pipeline_options import LayoutOptions
from docling.datamodel.pipeline_options import PdfPipelineOptions, ConvertPipelineOptions, TableFormerMode
from docling.datamodel.base_models import InputFormat
from docling.datamodel.layout_model_specs import DOCLING_LAYOUT_HERON_101, DOCLING_LAYOUT_EGRET_XLARGE


def _find_libreoffice() -> str | None:
    for cmd in ("libreoffice", "soffice"):
        p = shutil.which(cmd)
        if p:
            return p
    return None


def _rtf_to_docx_or_pdf_libreoffice(rtf_path: str, target: str) -> tuple[str, str]:
    """RTF → DOCX или PDF через LibreOffice. Возвращает (путь к файлу, tempdir для удаления)."""
    lo = _find_libreoffice()
    if not lo:
        raise RuntimeError("LibreOffice не найден в PATH (ожидаются команды libreoffice или soffice).")

    if target not in ("docx", "pdf"):
        target = "docx"

    out_dir = tempfile.mkdtemp(prefix="rtf_libreoffice_")
    base_name = os.path.splitext(os.path.basename(rtf_path))[0]
    expected_ext = ".docx" if target == "docx" else ".pdf"
    expected = os.path.join(out_dir, base_name + expected_ext)

    env = {**os.environ, "HOME": out_dir}
    cmd = [
        lo,
        "--headless",
        "--norestore",
        "--nolockcheck",
        "--nodefault",
        "--convert-to",
        target,
        "--outdir",
        out_dir,
        os.path.abspath(rtf_path),
    ]
    proc = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        timeout=180,
        env=env,
        cwd=out_dir,
    )
    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout or "").strip()
        raise RuntimeError(f"LibreOffice завершился с кодом {proc.returncode}: {err}")

    if os.path.isfile(expected):
        out_path = expected
    else:
        matches = [
            os.path.join(out_dir, f)
            for f in os.listdir(out_dir)
            if f.lower().endswith(expected_ext)
        ]
        if len(matches) != 1:
            raise RuntimeError(
                f"LibreOffice не создал ожидаемый {expected_ext}: {os.listdir(out_dir)}"
            )
        out_path = matches[0]

    return out_path, out_dir


def _rtf_fallback_striprtf(path: str) -> tuple[str, str]:
    """Запасной вариант без LibreOffice: только текст → .txt."""
    from striprtf.striprtf import rtf_to_text

    with open(path, "rb") as f:
        raw = f.read()
    for enc in ("utf-8-sig", "utf-8", "cp1251", "latin-1"):
        try:
            rtf_str = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        rtf_str = raw.decode("utf-8", errors="replace")

    plain = rtf_to_text(rtf_str)
    out_dir = tempfile.mkdtemp(prefix="rtf_striprtf_")
    out_path = os.path.join(out_dir, "extracted.txt")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(plain)
    return out_path, out_dir


def _prepare_path_for_docling(path: str) -> tuple[str, str | None, str | None]:
    """Docling не поддерживает RTF: конвертация LibreOffice → DOCX/PDF (или fallback .txt)."""
    lower = path.lower()
    if not lower.endswith(".rtf"):
        return path, None, None

    target = os.environ.get("RTF_CONVERT_TO", "docx").strip().lower()
    if target not in ("docx", "pdf"):
        target = "docx"

    if _find_libreoffice():
        out_path, tmp_dir = _rtf_to_docx_or_pdf_libreoffice(path, target)
        note = f"RTF сконвертирован LibreOffice в {target.upper()}."
        return out_path, tmp_dir, note

    out_path, tmp_dir = _rtf_fallback_striprtf(path)
    note = (
        "LibreOffice недоступен: RTF разобран как текст (striprtf). "
        "Установите LibreOffice для конвертации в DOCX/PDF."
    )
    return out_path, tmp_dir, note


def _pick_first(d: dict, keys: list[str]):
    for k in keys:
        if k in d and d[k] not in (None, "", [], {}):
            return d[k]
    return None


def _document_summary(document_dump: dict, path: str) -> dict:
    pages = document_dump.get("pages")
    page_count = len(pages) if isinstance(pages, list) else None

    meta = document_dump.get("metadata") if isinstance(document_dump.get("metadata"), dict) else {}

    out = {
        "file": {
            "name": os.path.basename(path),
            "path": path,
            "ext": os.path.splitext(path)[1].lower(),
            "size_bytes": (os.path.getsize(path) if os.path.exists(path) and os.path.isfile(path) else None),
        },
    }

    # Убираем пустые значения, чтобы JSON был "чистым"
    def _compact(v):
        if isinstance(v, dict):
            vv = {k: _compact(x) for k, x in v.items()}
            return {k: x for k, x in vv.items() if x not in (None, "", [], {})}
        if isinstance(v, list):
            vv = [_compact(x) for x in v]
            return [x for x in vv if x not in (None, "", [], {})]
        return v

    return _compact(out)


def _label_to_str(label) -> str:
    return label.value if hasattr(label, "value") else str(label)


def _safe_truncate(s: str, limit: int) -> str:
    s = s.strip()
    if len(s) <= limit:
        return s
    return s[: max(0, limit - 1)].rstrip() + "…"


def _extract_text_structure(doc) -> dict:
    root_items: list[dict] = []
    section_stack: list[dict] = []

    def _get_page_no(x):
        prov = getattr(x, "prov", None)
        if isinstance(prov, list) and len(prov) > 0 and hasattr(prov[0], "page_no"):
            return getattr(prov[0], "page_no")
        return None

    for item, _ in doc.iterate_items(with_groups=False):
        label = getattr(item, "label", None)
        if label is None:
            continue

        label_str = _label_to_str(label)

        text = getattr(item, "text", None)
        page_no = _get_page_no(item)

        if label_str == "section_header":
            if not isinstance(text, str) or not text.strip():
                continue
            level = getattr(item, "level", 1)

            # Закрываем секции с уровнем >= текущего (новый заголовок "на том же или выше" уровне)
            while section_stack and section_stack[-1].get("level", 1) >= level:
                section_stack.pop()

            node = {
                "type": "section_header",
                "text": _safe_truncate(text, 500),
                "level": level,
                "children": [],
            }
            if page_no is not None:
                node["page_no"] = page_no

            if section_stack:
                section_stack[-1]["children"].append(node)
            else:
                root_items.append(node)

            section_stack.append(node)
            continue

        if label_str == "formula":
            formula_placeholder = (
                f"смотри формулу на странице №{page_no}" if page_no is not None else "смотри формулу"
            )
            out = {"type": "formula", "text": formula_placeholder}
        else:
            if not isinstance(text, str) or not text.strip():
                continue
            out = {
                "type": label_str,
                "text": _safe_truncate(text, 500),
            }
        if page_no is not None:
            out["page_no"] = page_no

        # Если после section_header идут text/paragraph — вкладываем в текущую секцию
        if section_stack and label_str in {"text", "paragraph", "list_item", "code", "footnote", "caption", "formula"}:
            section_stack[-1]["children"].append(out)
        else:
            root_items.append(out)

    return {"items": root_items, "items_count": len(root_items)}


class Model:
    def __init__(self):
        super().__init__()

    def _some_method(self):
        pass

    def __call__(self, path: str):
        valid, message = validate_input_file(path)
        if not valid:
            return False, message
        
        pdf_pipeline_options = PdfPipelineOptions()
        pdf_pipeline_options.enable_remote_services = False
        pdf_pipeline_options.do_ocr = True  # Включить OCR
        pdf_pipeline_options.do_table_structure = True  # Включить анализ структуры таблиц
        pdf_pipeline_options.table_structure_options.mode = TableFormerMode.ACCURATE
        pdf_pipeline_options.do_formula_enrichment = True  # Включить распознавание формул (LaTeX)
        pdf_pipeline_options.table_structure_options.do_cell_matching = False  # Использовать предсказанные моделью ячейки
        pdf_pipeline_options.layout_options = LayoutOptions(layout_spec=DOCLING_LAYOUT_HERON_101)

        doc_pipeline_options = ConvertPipelineOptions()
        doc_pipeline_options.enable_remote_services = False

        convert_path, rtf_tmp_dir, rtf_note = _prepare_path_for_docling(path)
        try:
            converter = DocumentConverter(
                format_options={InputFormat.PDF: PdfFormatOption(pipeline_options=pdf_pipeline_options),
                                InputFormat.DOCX: WordFormatOption(pipeline_options=doc_pipeline_options)
                                }
            )
            result = converter.convert(convert_path)
            full = result.document.model_dump(mode="json")
            out = _document_summary(full, path)
            out["text_structure"] = _extract_text_structure(result.document)
            if rtf_note is not None:
                out["rtf_note"] = rtf_note
            return True, out
        except ConversionError as e:
            return False, f"ConversionError: {e}"
        except Exception as e:
            return False, f"Unexpected conversion error: {e}"
        finally:
            if rtf_tmp_dir and os.path.isdir(rtf_tmp_dir):
                try:
                    shutil.rmtree(rtf_tmp_dir, ignore_errors=True)
                except OSError:
                    pass


# Функция проверки файла
def validate_input_file(path: str) -> tuple[bool, str]:
    if not os.path.exists(path):
        return False, "File not found"
    if not os.path.isfile(path):
        return False, "Not a file"
    if not path.endswith(('.pdf', '.docx', '.doc', '.txt', '.rtf')):
        return False, "Unsupported file type"
    try:
        size = os.path.getsize(path)
    except OSError as e:
        return False, f"Не удалось получить размер файла: {e}"
    if size <= 0:
        return False, "Файл пустой."
    return True, "File is valid"
