from replacer.common import replace_common

def replace_shapes(doc, invoice_data):
    for i in range(1,doc.Shapes.Count + 1):
        shape = doc.Shapes(i);

        if shape.TextFrame.HasText:
            tr = shape.TextFrame.TextRange.Find
            replace_common(tr,invoice_data)            