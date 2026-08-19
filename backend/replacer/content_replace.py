from replacer.common import replace_common

def replace_content(doc, invoice_data):
    find = doc.Content.Find
    replace_common(find, invoice_data)