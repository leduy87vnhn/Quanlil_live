from typing import Dict, Iterable
import xml.etree.ElementTree as ET

def extract_fields_from_state(xml_text: str, field_names: Iterable[str]) -> Dict[str, str]:
    results: Dict[str, str] = {fn: "" for fn in field_names}
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        return results

    for elem in root.iter():
        tag = elem.tag
        if "}" in tag:
            tag = tag.split("}", 1)[1]
        for fn in field_names:
            if not results.get(fn):
                if tag == fn and (elem.text is not None):
                    results[fn] = elem.text.strip()
                elif elem.get("name") == fn and (elem.text is not None):
                    results[fn] = elem.text.strip()
        for child in list(elem):
            name = child.get("name")
            if name in field_names and (child.text is not None) and not results.get(name):
                results[name] = child.text.strip()
    return results
