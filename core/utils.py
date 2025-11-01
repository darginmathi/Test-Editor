import re

PREFIXES = ["btn", "lnk", "ddl", "cel", "chk", "txt", "lbl", "tgl", "cbk"]

def clean_object_name(name):
    if not isinstance(name, str):
        return name

    pattern = re.compile(f"^({'|'.join(PREFIXES)})", re.IGNORECASE)

    cleaned_name = pattern.sub("", name)

    return cleaned_name
