from replacer.common import replace_common

def replace_shapes(doc, invoice_data):
    for i in range(1, doc.Shapes.Count + 1):
        try:
            shape = doc.Shapes(i)
            if shape.TextFrame.HasText:
                tr = shape.TextFrame.TextRange.Find
                replace_common(tr, invoice_data)
        except Exception:
            pass

    for i in range(1, doc.InlineShapes.Count + 1):
        try:
            ishp = doc.InlineShapes(i)
            if hasattr(ishp, "HasTextFrame") and ishp.HasTextFrame and ishp.TextFrame.HasText:
                tr = ishp.TextFrame.TextRange.Find
                replace_common(tr, invoice_data)
        except Exception:
            pass