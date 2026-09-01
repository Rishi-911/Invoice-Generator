from formatter import format_value

WD_REPLACE_ALL = 2

def replace_common(find, invoice_data):
    try:
        text = find.Parent.Text
    except Exception:
        text = ""

    if not text or "<<" not in text:
        return

    find.ClearFormatting()
    find.Replacement.ClearFormatting()
    find.Forward = True
    find.MatchCase = False
    find.MatchWholeWord = False
    find.MatchWildcards = False

    for key,value in invoice_data.items():
        placeholder = f"<<{key.strip()}>>"
        if placeholder.lower() in text.lower():
            find.Execute(
                FindText = placeholder,
                ReplaceWith = format_value(key,value),
                Replace = WD_REPLACE_ALL
            )

 