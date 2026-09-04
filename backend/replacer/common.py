import re
import logging
from formatter import format_value

logger = logging.getLogger("invoice_generator")

WD_REPLACE_ALL = 2
WD_FIND_CONTINUE = 1
WD_REPLACE_NONE = 0

def replace_common(find, invoice_data):
    try:
        text = find.Parent.Text
    except Exception:
        text = ""

    if not text or "<<" not in text:
        return

    tags = set(re.findall(r"<<\s*[^>]+?\s*>>", text))
    if not tags:
        return

    norm_map = {}
    for k, v in invoice_data.items():
        if k is not None:
            norm_map[str(k).strip().lower()] = (k, v)

    for tag in tags:
        inner = tag[2:-2].strip().lower()
        if inner in norm_map:
            orig_key, val = norm_map[inner]
            rep_text = format_value(orig_key, val)
            if rep_text is None:
                rep_text = ""
            else:
                rep_text = str(rep_text)

            try:
                if len(rep_text) > 250:
                    max_loops = 50
                    loop_cnt = 0
                    find.ClearFormatting()
                    while loop_cnt < max_loops and find.Execute(
                        tag, False, False, False, False, False, True, WD_FIND_CONTINUE, False, "", WD_REPLACE_NONE
                    ):
                        find.Parent.Text = rep_text
                        loop_cnt += 1
                else:
                    find.ClearFormatting()
                    find.Replacement.ClearFormatting()
                    find.Execute(
                        tag,
                        False,
                        False,
                        False,
                        False,
                        False,
                        True,
                        WD_FIND_CONTINUE,
                        False,
                        rep_text,
                        WD_REPLACE_ALL
                    )
            except Exception as e:
                logger.warning(f"Error replacing tag '{tag}': {e}")