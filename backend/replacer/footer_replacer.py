from replacer.common import replace_common

def replace_footer(doc, invoice_data):
    try:
        for section in doc.Sections:
            for footer in section.Footers:
                try:
                    if footer.Exists:
                        find = footer.Range.Find
                        replace_common(find, invoice_data)
                except Exception:
                    pass
    except Exception:
        pass