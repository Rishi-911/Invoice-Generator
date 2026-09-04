import win32com.client
import win32com.client.dynamic
from excel_reader import read_file
import pythoncom
import logging
import os
import shutil
import re
from pathlib import Path
from config import PDF_OUTPUT_FOLDER
from replacer.content_replace import replace_content
from replacer.shape_replacer import replace_shapes
from replacer.header_replacer import replace_header
from replacer.footer_replacer import replace_footer

logger = logging.getLogger("invoice_generator")

WD_FORMAT_PDF = 17

def sanitize_filename(name) -> str:
    if name is None or (isinstance(name, float) and str(name).strip().lower() in ("nan", "none", "")):
        return "unnamed"
    s = str(name).strip()
    s = re.sub(r'[\\/:*?"<>|\r\n]+', '_', s)
    s = s.strip('. ')
    return s if s else "unnamed"

def get_word_app():
    try:
        gen_py_path = getattr(win32com, "__gen_path__", None)
        if gen_py_path and os.path.exists(gen_py_path):
            shutil.rmtree(gen_py_path, ignore_errors=True)
    except Exception:
        pass

    try:
        return win32com.client.DispatchEx("Word.Application")
    except Exception as e:
        logger.warning(f"DispatchEx failed: {e}. Trying dynamic Dispatch...")
        try:
            return win32com.client.dynamic.Dispatch("Word.Application")
        except Exception:
            return win32com.client.Dispatch("Word.Application")

def replace_placeholders(template_path, data_path):
    pythoncom.CoInitialize()
    word_main = None
    pdf_files = []

    try:
        resolved_template = Path(template_path).resolve()
        if not resolved_template.exists():
            raise FileNotFoundError(f"Template file not found: {resolved_template}")

        df = read_file(data_path)
        logger.info(f"Excel read completed. Total rows: {len(df)}")

        word_main = get_word_app()
        word_main.Visible = False
        word_main.DisplayAlerts = 0
        word_main.ScreenUpdating = False

        try:
            word_main.Options.SaveNormalPrompt = False
            word_main.Options.CheckSpellingAsYouType = False
            word_main.Options.CheckGrammarAsYouType = False
        except Exception as e:
            logger.warning(f"Could not set Word options: {e}")

        try:
            word_main.AutomationSecurity = 3
        except Exception:
            pass

        logger.info("Word COM Server initialized/dispatched")

        row_count = len(df)

        try:
            for idx, row in df.iterrows():
                invoice_data = row.to_dict()
                raw_name = row.get('Name') or row.get('name') or f'row_{idx+1}'
                row_name = sanitize_filename(raw_name)

                doc = None
                try:
                    doc = word_main.Documents.Add(Template=str(resolved_template))

                    replace_content(doc, invoice_data)
                    replace_shapes(doc, invoice_data)
                    replace_header(doc, invoice_data)
                    replace_footer(doc, invoice_data)

                    output_pdf = Path(PDF_OUTPUT_FOLDER) / f"{row_name}.pdf"
                    output_pdf.parent.mkdir(parents=True, exist_ok=True)

                    save_path_str = str(output_pdf.resolve())
                    try:
                        doc.SaveAs2(FileName=save_path_str, FileFormat=WD_FORMAT_PDF)
                    except Exception:
                        doc.SaveAs(FileName=save_path_str, FileFormat=WD_FORMAT_PDF)

                    pdf_files.append(output_pdf)
                    logger.info(f"[{idx+1}/{row_count}] Generated PDF for '{row_name}'")

                except Exception as doc_err:
                    logger.error(f"Error processing row {idx+1} ({row_name}): {doc_err}", exc_info=True)
                    raise doc_err
                finally:
                    if doc is not None:
                        try:
                            doc.Close(False)
                        except Exception:
                            pass

        finally:
            if word_main is not None:
                try:
                    word_main.NormalTemplate.Saved = True
                except Exception:
                    pass
                try:
                    word_main.Quit()
                except Exception:
                    pass
                logger.info("Word COM Server Quit completed")
    finally:
        pythoncom.CoUninitialize()

    logger.info(f"replace_placeholders completed. Total files: {len(pdf_files)}")
    return pdf_files