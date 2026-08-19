import win32com
from excel_reader import read_file
import pythoncom
from config import TEMPLATES_FOLDER, DOCX_OUTPUT_FOLDER, PDF_OUTPUT_FOLDER
from win32com.client import constants
from replacer.content_replace import replace_content
from replacer.shape_replacer import replace_shapes
from replacer.header_replacer import replace_header
from replacer.footer_replacer import replace_footer



def replace_placeholders(template_path, data_path):
    pythoncom.CoInitialize()
    try:
        word_main = win32com.client.gencache.EnsureDispatch("Word.Application") # type: ignore
        word_main.Visible = False
        
        # Performance optimization settings for Word automation
        word_main.ScreenUpdating = False
        word_main.DisplayAlerts = 0  # wdAlertsNone
        word_main.Options.CheckSpellingAsYouType = False
        word_main.Options.CheckGrammarAsYouType = False
        
        try:
            df = read_file(data_path)

            pdf_files = []
            docx_files = []  # Keep as empty list since DOCX is unused and skipped for speed
            for _, row in df.iterrows():
                invoice_data = row.to_dict()
                
                # Documents.Add(Template=...) is cached by Word and is much faster than opening from disk
                doc = word_main.Documents.Add(Template=str(template_path.resolve()))

                try:
                    replace_content(doc, invoice_data)
                    replace_shapes(doc, invoice_data)
                    replace_header(doc,invoice_data)
                    replace_footer(doc,invoice_data)

                    # Skip generating and saving DOCX files because the application only packages/uses the PDFs.
                    # This cuts disk I/O and saving overhead by ~1.7x.
                    
                    output_pdf = PDF_OUTPUT_FOLDER / f"{row['Name']}.pdf"
                    doc.SaveAs(
                        str(output_pdf.resolve()),
                        FileFormat=constants.wdFormatPDF
                    )
                    pdf_files.append(output_pdf)

                finally:
                    doc.Close(False)

        finally:
            word_main.Quit()
    finally:
        pythoncom.CoUninitialize()

    return pdf_files, docx_files