import re
from collections import Counter
from typing import Dict, Any, Iterator, Tuple, Optional, Union


# ---------------------------------------------------------------------
# 1) ID-only process dictionary (authoritative matching by numeric ID)
# ---------------------------------------------------------------------
# Keep only what you want/need; you can expand this list freely.
MPR_PROCESS_DEFS: Dict[int, str] = {
    # 100: "Workpiece definition (Werkstck)",
    # 101: "Comment",
    102: "V_drill",
    103: "H_drill",
    104: "U_drilling",
    105: "Milling_from_top",
    106: "Edge-banding on contour",
    107: "Flush trimming on contour",
    108: "End trimming / capping on contour",
    109: "Saw_Grooving",
    112: "Pocket_milling",
    113: "Milling_from_below",
    # 114: "Vacuum cup / suction cup (Sauger)",
    # 115: "Console suction cup - type K (SaugerK)",
    # 116: "Console suction cup - type X (SaugerX)",
    # 117: "Console suction cup - type Y (SaugerY)",
    # 118: "Console suction cup - type Z (SaugerZ)",
    # 119: "Workpiece zero point / workpiece origin (Werkstck-Nullpunkt)",
    # 120: "Zero-point movement / origin movement (Nullpunktbewegung)",
    # 121: "Tool definition / tool macro (Werkzeug)",
    # 122: "Sequence / run macro (Ablauf)",
    # 123: "Measurement (Messen)",
    # 124: "Workpiece position probing / measure workpiece location (Lagevermessen)",
    124: "Angle_sawing[45_Handle]",
    # 125: "Block macro (Block)",
    # 126: "Blow-off / air cleaning on contour (Abblasen)",
    # 127: "Laser display / laser indication (Laser-Anzeige)",
    # 128: "Vertical tactile/probed milling (Vert Getastet-Fraesen)",
    # 129: "Free NC code block (NC-Code)",
    # 130: "NC stop macro (NC-Stop)",
    131: "Drilling from below",
    # 132: "Pocketing from below (Unterflur-Tasche)",
    133: "Contour milling",
    # 134: "Scribing / scoring an edge (Ritzen)",
    # 135: "Corner notching (Klink)",
    # 136: "Contour sanding (Schleifen)",
    # 137: "Pressure zone on contour (Drucken)",
    # 138: "T-molding / bridge edge / Stegkante (Stegkante)",
    # 139: "Component macro (Komponente)",
    # 140: "5-axis / vector milling (Vektorfraesen)",
    # 141: "Routing with CF unit (CF-Fraesen)",
    # 142: "Gluing with CF unit (CF-Leimen)",
    # 143: "Flush trimming with CF unit (CF-Buendigfraesen)",
    # 144: "Pressure with CF unit (CF-Druck)",
    # 145: "Cutting / end trimming with CF unit (CF-Kappen)",
    # 146: "Sanding with CF unit (CF-Schleifen)",
    # 147: "Outfeed / discharge transport (AusTransport)",
    # 148: "Workpiece transport with vacuum carpet (OfTransport)",
    # 149: "Area macro / region (Gebiet)",
    # 150: "Technology change / new setup macro (Neustr)",
    # 151: "Laser display / laser indication (Laser-Anzeige)",
    151: "Pocketing_from_below",
    # 152: "FK unit routing (FK-Fraesen)",
    # 153: "Angled/space grooving (Nut_R)",
    # 154: "Laser display of contour (Laser-Anzeige)",
    # 157: "CF flush trimming (CF-Buendigfraesen)",
    # 158: "CF routing (CF-Fraesen)",
    # 159: "CF gluing (CF-Leimen)",
    # 161: "FK routing (FK-Fraesen)",
    # 165: "CF pressure zone (CF-Druck)",
    # 166: "CF end trimming / capping (CF-Kappen)",
    # 167: "CF sanding (CF-Schleifen)",
    # 169: "CF station (CF-Station)",
    # 170: "CF KTR station (CF-KTRStation)",
    # 180: "Discharge / change macro (Neustr)",
    181: "Freeform pocket milling",
    # 182: "Area macro / region (Gebiet)",
    # 184: "Part handling / transport with bridge (Transport)",
}


# ---------------------------------------------------------------------
# 2) Text loading + macro parsing (ID-only header detection)
# ---------------------------------------------------------------------
def _read_text_input(mpr_input: Union[str, bytes]) -> str:
    """
    Accepts:
      - file path (str) OR raw MPR text (str) OR raw bytes (bytes)
    Returns decoded text.
    """
    if isinstance(mpr_input, bytes):
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                return mpr_input.decode(enc)
            except UnicodeDecodeError:
                pass
        return mpr_input.decode("latin-1", errors="replace")

    if isinstance(mpr_input, str):
        # If it's a real path, read it; otherwise treat as content.
        try:
            with open(mpr_input, "rb") as f:
                return _read_text_input(f.read())
        except (FileNotFoundError, OSError):
            return mpr_input

    raise TypeError("mpr_input must be a file path (str), raw text (str), or raw bytes (bytes).")


def _iter_mpr_macro_blocks(text: str) -> Iterator[Tuple[int, str]]:
    """
    Yields (macro_id, macro_block_text).

    IMPORTANT: Matching is based ONLY on macro ID, not macro name.
    Macro header pattern: <102 \...\
    We only rely on: ^\s*<\s*(\d+)\s*\
    """
    header_re = re.compile(r'(?m)^\s*<\s*(\d+)\s*\\')
    matches = list(header_re.finditer(text))
    for i, m in enumerate(matches):
        start = m.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        macro_id = int(m.group(1))
        yield macro_id, text[start:end]


def _get_param(block: str, key: str) -> Optional[str]:
    """
    Extracts parameter values like:
      BM="LSL"  or  DU="5"  or  T_="123"
    """
    pattern = re.compile(rf'(?mi)^\s*{re.escape(key)}\s*=\s*(?:"([^"]*)"|([^\s\\\r\n]+))')
    m = pattern.search(block)
    if not m:
        return None
    return (m.group(1) if m.group(1) is not None else m.group(2)).strip()


def _format_diameter(du_raw: Optional[str], tool_raw: Optional[str], *, tool_prefix: str = "Tool") -> str:
    if du_raw:
        try:
            v = float(du_raw)
            if abs(v - int(v)) < 1e-9:
                return f"{int(v)}D"
            return f"{v}D"
        except ValueError:
            safe = re.sub(r"\s+", "", du_raw)
            return f"{safe}D"

    if tool_raw:
        safe = re.sub(r"\s+", "", tool_raw)
        return f"{tool_prefix}{safe}"

    return "DUNK"


# ---------------------------------------------------------------------
# 3) Special signatures for drilling macros (still ID-driven)
# ---------------------------------------------------------------------
# ID 102: BohrVert signature
BM_STYLE_MAP_VERT = {
    "LS": "SF",
    "SS": "FF",
    "LSL": "SFS",
    "SSS": "FFF",
    "LSU": "SF",
    "LSLU": "SFS",
}
BM_SUFFIX_MAP_VERT = {
    "LS": "ToDepth",
    "SS": "ToDepth",
    "LSL": "Through",
    "SSS": "Through",
    "LSU": "FromBottom",
    "LSLU": "FromBottom",
}

def bohrvert_signature(block: str) -> str:
    bm = (_get_param(block, "BM") or "").upper()
    du = _get_param(block, "DU")
    tno = _get_param(block, "TNO")  # common alternative for vertical
    diam = _format_diameter(du, tno, tool_prefix="Tool")

    style = BM_STYLE_MAP_VERT.get(bm, f"BM{bm or 'UNK'}")
    suffix = BM_SUFFIX_MAP_VERT.get(bm, "UNK")
    return f"VDrill_{diam}_{style}_{suffix}"


# ID 103: BohrHoriz signature
BM_DIR_MAP_HORIZ = {
    "XP": "+X",
    "XM": "-X",
    "YP": "+Y",
    "YM": "-Y",
}

def bohrhoriz_signature(block: str) -> str:
    bm = (_get_param(block, "BM") or "").upper()
    du = _get_param(block, "DU")
    t_ = _get_param(block, "T_")  # common alternative for horizontal
    diam = _format_diameter(du, t_, tool_prefix="Tool")

    if bm in BM_DIR_MAP_HORIZ:
        dir_part = BM_DIR_MAP_HORIZ[bm]
    elif bm == "C":
        wi = _get_param(block, "WI")  # optional free C-angle
        dir_part = f"C{wi}" if wi else "C"
    else:
        dir_part = f"BM{bm or 'UNK'}"

    return f"HDrill_{diam}_{dir_part}"


# ---------------------------------------------------------------------
# 4) Main: map + count (ID-only matching)
# ---------------------------------------------------------------------
def map_and_count_mpr_processes(
    mpr_input: Union[str, bytes],
    process_defs: Dict[int, str] = MPR_PROCESS_DEFS,
    *,
    include_disabled: bool = True,
) -> Dict[str, Any]:
    """
    Counts macros by ID (authoritative), then enriches with:
      - mapped description (if ID exists in process_defs)
      - special signature counts for ID 102 and 103
      - angle groove lengths for ID 124 (XA/YA -> XE/YE)
    """
    text = _read_text_input(mpr_input)

    counts_by_id = Counter()
    unknown_ids = Counter()

    bohrvert_sig_counts = Counter()
    bohrhoriz_sig_counts = Counter()
    angle124_lengths = []
    groove109_lengths = []

    la_100 = br_100 = 0.0
    # Pre-parse macro 100 for LA/BR
    for macro_id, block in _iter_mpr_macro_blocks(text):
        if macro_id == 100:
            la_val = _get_param(block, "LA")
            br_val = _get_param(block, "BR")
            try:
                la_100 = float(la_val) if la_val is not None else 0.0
            except ValueError:
                la_100 = 0.0
            try:
                br_100 = float(br_val) if br_val is not None else 0.0
            except ValueError:
                br_100 = 0.0
            break

    for macro_id, block in _iter_mpr_macro_blocks(text):
        if not include_disabled:
            en = _get_param(block, "EN")
            if en is not None and en.strip() == "0":
                continue

        counts_by_id[macro_id] += 1
        if macro_id not in process_defs:
            unknown_ids[macro_id] += 1

        if macro_id == 102:
            bohrvert_sig_counts[bohrvert_signature(block)] += 1
        elif macro_id == 103:
            bohrhoriz_sig_counts[bohrhoriz_signature(block)] += 1
        elif macro_id in (109, 124):
            # Groove length from XA/YA to XE/YE (one delta should be zero)
            def _safe_float(val: Optional[str]) -> Optional[float]:
                if val is None:
                    return None
                try:
                    return float(val)
                except ValueError:
                    return None

            xa = _safe_float(_get_param(block, "XA"))
            ya = _safe_float(_get_param(block, "YA"))
            xe = _safe_float(_get_param(block, "XE"))
            ye = _safe_float(_get_param(block, "YE"))
            t_val = (_get_param(block, "T_") or "").replace('"', "").replace("!", "").strip()
            if None not in (xa, ya, xe, ye):
                dx = abs(xa - xe)
                dy = abs(ya - ye)
                is_below = macro_id == 109 and t_val.endswith("xxxxx2")
                if macro_id == 124:
                    if dy == 0:
                        length_str = f"{dx}_On_PL<{la_100}>"
                    elif dx == 0:
                        length_str = f"{dy}_On_PW<{br_100}>"
                    else:
                        length_str = f"{max(dx, dy)}"
                    angle124_lengths.append(length_str)
                else:
                    if dy == 0:
                        suffix = "Milling_From_Below" if is_below else "Top_Saw_Grv"
                        length_str = f"{dx}_On_PL<{la_100}_{suffix}>"
                    elif dx == 0:
                        suffix = "Milling_From_Below" if is_below else "Top_Saw_Grv"
                        length_str = f"{dy}_On_PW<{br_100}_{suffix}>"
                    else:
                        length_str = f"{max(dx, dy)}"
                    groove109_lengths.append(length_str)

    # Build a mapped summary keyed by ID
    mapped_counts = {
        mid: {
            "count": cnt,
            "description": process_defs.get(mid, "Unknown/Unmapped macro ID"),
        }
        for mid, cnt in sorted(counts_by_id.items())
    }

    return {
        "process_counts_by_id": dict(counts_by_id),
        "mapped_process_counts": mapped_counts,
        "unknown_macro_counts": dict(unknown_ids),
        "bohrvert_signature_counts": dict(bohrvert_sig_counts),
        "bohrhoriz_signature_counts": dict(bohrhoriz_sig_counts),
        "angle124_lengths": angle124_lengths,
        "groove109_lengths": groove109_lengths,
        "la_100": la_100,
        "br_100": br_100,
    }


# ----------------------------
# Example
# ----------------------------
# result = map_and_count_mpr_processes(r"C:\path\part.mpr")
# print(result["mapped_process_counts"])          # ID-driven macro summary
# print(result["bohrvert_signature_counts"])      # VDrill_...
# print(result["bohrhoriz_signature_counts"])     # HDrill_...