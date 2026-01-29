import argparse
import os
import yaml
from typing import Dict
from .utils import setup_logging, logger
from .gsheet_client import GSheetClient
from .vmix_client import VmixClient
from .mapper import map_row_to_commands
from .vmix_parser import extract_fields_from_state
import time
from requests.exceptions import RequestException

def load_config(path: str) -> Dict:
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def push_once(cfg: Dict):
    setup_logging()
    logger.info("Start push (Sheets -> vMix)")
    gcfg = cfg["sheets"]
    vcfg = cfg["vmix"]
    gs = GSheetClient(gcfg["spreadsheet_id"], gcfg["credentials_path"])
    rows = gs.read_table(gcfg["read_range"])
    vm = VmixClient(host=vcfg.get("host", "127.0.0.1"), port=vcfg.get("port", 8088))
    field_map = cfg.get("field_map", {})
    max_retries = cfg.get("retries", 3)
    backoff = cfg.get("backoff_seconds", 1)
    for row_idx, r in enumerate(rows, start=1):
        cmds = map_row_to_commands(r, field_map)
        if not cmds:
            logger.debug("Row %d: no mapped commands, skipping", row_idx)
            continue
        for vmix_field, value in cmds:
            attempt = 0
            while attempt < max_retries:
                try:
                    vm.set_text(selected_name=vmix_field, value=value, input_index=vcfg.get("input"))
                    logger.info("Row %d: set %s=%s", row_idx, vmix_field, value)
                    break
                except RequestException as e:
                    attempt += 1
                    logger.warning("vMix request failed (attempt %d/%d): %s", attempt, max_retries, e)
                    if attempt >= max_retries:
                        logger.error("Failed to set %s after %d attempts", vmix_field, max_retries)
                    else:
                        sleep_t = backoff * (2 ** (attempt - 1))
                        time.sleep(sleep_t)
    logger.info("Push complete")

def pull_once(cfg: Dict):
    setup_logging()
    logger.info("Start pull (vMix -> Sheets)")
    gcfg = cfg["sheets"]
    vcfg = cfg["vmix"]
    gs = GSheetClient(gcfg["spreadsheet_id"], gcfg["credentials_path"])
    vm = VmixClient(host=vcfg.get("host", "127.0.0.1"), port=vcfg.get("port", 8088))
    xml = vm.get_state()
    field_map = cfg.get("field_map", {})
    vm_fields = list(field_map.values())
    extracted = extract_fields_from_state(xml, vm_fields)
    sheet_cols = list(field_map.keys())
    header = sheet_cols
    values = [extracted.get(field_map[col], "") for col in sheet_cols]
    # Chỉ update từng cell, không update cả dòng/sheet để tránh ghi đè công thức hoặc dữ liệu khác
    for idx, col in enumerate(sheet_cols):
        cell_range = f"{gcfg.get('write_range').split('!')[0]}!{chr(65+idx)}2"  # luôn ghi dòng 2, cột tương ứng
        gs.write_table(cell_range, [[values[idx]]])
        logger.info(f"Update cell {cell_range} = {values[idx]}")
    logger.info("Pull complete (cell by cell, no overwrite)")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", default="config.yaml", help="Path to config YAML")
    parser.add_argument("--once", action="store_true", help="Run one iteration and exit")
    parser.add_argument("--push", action="store_true", help="Push Sheets -> vMix")
    parser.add_argument("--pull", action="store_true", help="Pull vMix -> Sheets")
    args = parser.parse_args()
    if not os.path.exists(args.config):
        print(f"Config not found: {args.config}")
        return
    cfg = load_config(args.config)
    if args.once and args.push:
        push_once(cfg)
    if args.once and args.pull:
        pull_once(cfg)

if __name__ == "__main__":
    main()
