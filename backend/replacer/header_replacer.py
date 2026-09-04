from replacer.common import replace_common

def replace_header(doc, invoice_data):
    try:
        for section in doc.Sections:
            for header in section.Headers:
                try:
                    if header.Exists:
                        find = header.Range.Find
                        replace_common(find, invoice_data)
                except Exception:
                    pass
    except Exception:
        pass