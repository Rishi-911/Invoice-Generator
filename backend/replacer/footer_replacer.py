from replacer.common import replace_common

def replace_footer(doc,invoice_data):
    for section in doc.Sections:
        for footer in section.Footers:
            if footer.Exists:
                find = footer.Range.Find
                replace_common(find, invoice_data)