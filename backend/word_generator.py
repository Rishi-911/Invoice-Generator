import win32com.client
from excel_reader import read_file
import pythoncom
import logging
import os
import shutil
from config import TEMPLATES_FOLDER, DOCX_OUTPUT_FOLDER, PDF_OUTPUT_FOLDER
from replacer.content_replace import replace_content
from replacer.shape_replacer import replace_shapes
from replacer.header_replacer import replace_header
from replacer.footer_replacer import replace_footer

logger = logging.getLogger("invoice_generator")

WD_FORMAT_PDF = 17

def get_word_app():
    try:
        return win32com.client.Dispatch("Word.Application")
    except Exception as e:
        logger.warning(f"Initial Word COM Dispatch failed: {e}. Clearing gen_py cache...")
        try:
            gen_py_path = getattr(win32com, "__gen_path__", None)
            if gen_py_path and os.path.exists(gen_py_path):
                shutil.rmtree(gen_py_path, ignore_errors=True)
        except Exception:
            pass
        return win32com.client.Dispatch("Word.Application")

def replace_placeholders(template_path, data_path):
    pythoncom.CoInitialize()
    try:
        df = read_file(data_path)
        logger.info(f"Excel read completed. Total rows: {len(df)}")

        word_main = get_word_app()
        word_main.Visible = False
        
        word_main.ScreenUpdating = False
        word_main.DisplayAlerts = 0
        word_main.Options.CheckSpellingAsYouType = False
        word_main.Options.CheckGrammarAsYouType = False
        logger.info("Word COM Server initialized/dispatched")
        
        try:
            pdf_files = []
            docx_files = []
            
            row_count = len(df)
            
            for idx, row in df.iterrows():
                invoice_data = row.to_dict()
                row_name = row.get('Name', f'row_{idx}')
                
                doc = word_main.Documents.Add(Template=str(template_path.resolve()))

                try:
                    replace_content(doc, invoice_data)
                    replace_shapes(doc, invoice_data)
                    replace_header(doc, invoice_data)
                    replace_footer(doc, invoice_data)
                    
                    output_pdf = PDF_OUTPUT_FOLDER / f"{row_name}.pdf"
                    doc.SaveAs(
                        str(output_pdf.resolve()),
                        FileFormat=WD_FORMAT_PDF
                    )
                    
                    pdf_files.append(output_pdf)
                    logger.info(f"[{idx+1}/{row_count}] Generated PDF for '{row_name}'")

                finally:
                    doc.Close(False)

        finally:
            word_main.Quit()
            logger.info("Word COM Server Quit completed")
    finally:
        pythoncom.CoUninitialize()
    
    logger.info(f"replace_placeholders completed. Total files: {len(pdf_files)}")
    return pdf_files, docx_files