from typing import Dict, List, Tuple

def map_row_to_commands(row: Dict[str, str], field_map: Dict[str, str]) -> List[Tuple[str, str]]:
    cmds = []
    for sheet_col, vmix_field in field_map.items():
        val = row.get(sheet_col)
        if val is None:
            continue
        s = str(val).strip()
        if s == "":
            continue
        cmds.append((vmix_field, s))
    return cmds
