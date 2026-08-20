import win32com
from excel_reader import read_file
import pythoncom
import logging
from config import TEMPLATES_FOLDER, DOCX_OUTPUT_FOLDER, PDF_OUTPUT_FOLDER
from win32com.client import constants
from replacer.content_replace import replace_content
from replacer.shape_replacer import replace_shapes
from replacer.header_replacer import replace_header
from replacer.footer_replacer import replace_footer

logger = logging.getLogger("invoice_generator")

def replace_placeholders(template_path, data_path):
    pythoncom.CoInitialize()
    try:
        df = read_file(data_path)
        logger.info(f"Excel read completed. Total rows: {len(df)}")

        word_main = win32com.client.gencache.EnsureDispatch("Word.Application")
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
                        FileFormat=constants.wdFormatPDF
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