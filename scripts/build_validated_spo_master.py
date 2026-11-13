"""
Build the master tab and summary table from the human-validated SPO sentiment workbook.

Input (internal, NOT committed to Git):
  outputs/analysis_results/no_interview/sentiments/spo_validated_final.xlsx
    One tab per shared node (entities present in both the Option 1 and Option 2
    knowledge graphs). Each row is an SPO triple touching that node, with Role
    (Agent = node is the subject, Patient = node is the object), a manually
    validated Sentiment label, and the team's working notes and source quotes.

Outputs (public, committed; working notes and quotes removed at JX's request):
  outputs/analysis_results/no_interview/sentiments/spo_validated_public.xlsx
      - Master          one row per triple instance (deduplicated across tabs)
      - Summary         sentiment counts and % per node x corpus x role (live formulas on Master)
      - Corpus_Totals   sentiment counts and % per corpus from the deduplicated Master
      - Reconciliation  every merge and correction applied, for audit
      - one tab per shared node (Who, Action, To_Whom, Source, Role, Sentiment only)
  outputs/analysis_results/no_interview/sentiments/spo_validated_master.csv
  outputs/analysis_results/no_interview/sentiments/spo_validated_summary.csv

With --internal, also writes spo_validated_master_internal.xlsx / .csv, which keep the
notes and quotes for the team. These match *_internal.* and are excluded by .gitignore.

Counting rules (agreed with JX, Nov 2026):
  * Repeated triples WITHIN a tab are kept: they come from different passages.
  * A triple linking two shared nodes appears in two tabs (once under its subject node
    as Agent, once under its object node as Patient). In Master it is kept once, with
    both Subject_Node and Object_Node filled in, so corpus totals do not double count.
  * In Summary a node's Agent rows are Master rows with Subject_Node == node, and its
    Patient rows are Master rows with Object_Node == node. A cross-tab triple therefore
    counts once for each node it touches, as it should for per-node comparison.

Either workbook can be used as --input: the generated sheets are skipped on reading.

Usage:
  python scripts/build_validated_spo_master.py            # public outputs
  python scripts/build_validated_spo_master.py --internal # also the internal copy
"""

import argparse
import re
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SENT_DIR = "outputs/analysis_results/no_interview/sentiments"
CORE = ["Who", "Action", "To_Whom", "Source", "Role", "Sentiment"]
NOTE_COLS = ["Notes", "Second_Reviewer_Notes", "Original_Sentence", "Context"]
SENTIMENTS = ["Positive", "Negative", "Neutral"]
CORPORA = ["option1", "option2"]
ROLES = ["Agent", "Patient"]
GENERATED_SHEETS = {"Master", "Summary", "Corpus_Totals", "Reconciliation"}

# Triples worded slightly differently in two tabs but referring to the same edge.
# Each entry maps a variant (Who, Action, To_Whom) to the canonical wording.
EQUIVALENT_TRIPLES = {
    ("Unionist", "(Not) Want", "Rule"): ("Unionist", "(not) Want", "Rule (by Dublin)"),
}

# Label corrections confirmed by JX, applied to every tab. Key: (Who, Action, To_Whom, Source).
SENTIMENT_CORRECTIONS = {
    ("Government", "Introduce", "Internment", "option1"): (
        "Negative",
        "Was Neutral in the Government tab and Negative in the Internment tab; "
        "JX confirmed Negative (Nov 2026).",
    ),
}


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def clean(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def text_col(df: pd.DataFrame, names: list[str]) -> pd.Series:
    """Join whichever of the named columns exist; empty strings if none do."""
    present = [c for c in names if c in df.columns]
    if not present:
        return pd.Series([""] * len(df), index=df.index)
    return df[present].apply(lambda r: " | ".join(clean(v) for v in r if clean(v)), axis=1)


def load_tabs(path) -> tuple[list[str], pd.DataFrame, list]:
    sheets = pd.read_excel(path, sheet_name=None, dtype=str)
    frames = []
    for tab, df in sheets.items():
        if tab in GENERATED_SHEETS or "Who" not in df.columns:
            continue
        df = df.dropna(subset=["Who"]).copy()
        df["Tab"] = tab
        frames.append(df)
    nodes = [f["Tab"].iloc[0] for f in frames]
    d = pd.concat(frames, ignore_index=True)
    for c in CORE:
        d[c] = d[c].map(clean)
    d["Notes"] = text_col(d, ["Notes"])
    d["Second_Reviewer_Notes"] = text_col(d, ["SD Notes", "SadhbhNotes", "Second_Reviewer_Notes"])
    d["Original_Sentence"] = text_col(d, ["OriginalSentence (quote)", "Original_Sentence"])
    d["Context"] = text_col(d, ["Context (page number, book author)", "Context"])

    log = []
    for i, r in d.iterrows():
        key = (r.Who, r.Action, r.To_Whom)
        if key in EQUIVALENT_TRIPLES:
            canon = EQUIVALENT_TRIPLES[key]
            d.loc[i, ["Who", "Action", "To_Whom"]] = canon
            log.append({"Type": "Wording merged", "Triple": f"{' | '.join(key)} [{r.Source}]",
                        "Tabs": r.Tab, "Detail": f"Treated as the same triple as {' | '.join(canon)}"})

    for (who, act, obj, src), (label, why) in SENTIMENT_CORRECTIONS.items():
        hit = (d.Who == who) & (d.Action == act) & (d.To_Whom == obj) & (d.Source == src)
        changed = hit & (d.Sentiment != label)
        if changed.any():
            log.append({"Type": "Sentiment corrected", "Triple": f"{who} | {act} | {obj} [{src}]",
                        "Tabs": "; ".join(dict.fromkeys(d.loc[changed, "Tab"])),
                        "Detail": f"{int(changed.sum())} row(s) set to {label}. {why}"})
        d.loc[hit, "Sentiment"] = label

    d["Key"] = d.Who.map(norm) + "|" + d.Action.map(norm) + "|" + d.To_Whom.map(norm) + "|" + d.Source
    return nodes, d, log


def build_master(d: pd.DataFrame, log: list) -> pd.DataFrame:
    rows = []
    for key, g in d.groupby("Key", sort=False):
        tabs = list(dict.fromkeys(g.Tab))
        if len(tabs) == 1:
            for _, r in g.iterrows():
                rows.append({
                    **{c: r[c] for c in ["Who", "Action", "To_Whom", "Source", "Sentiment"]},
                    "Subject_Node": r.Tab if r.Role == "Agent" else "",
                    "Object_Node": r.Tab if r.Role == "Patient" else "",
                    "Tabs": r.Tab, "Reconciliation": "",
                    **{c: r[c] for c in NOTE_COLS},
                })
            continue

        # Triple appears in two tabs: keep it once per passage.
        parts = {t: g[g.Tab == t].reset_index(drop=True) for t in tabs}
        counts = {t: len(p) for t, p in parts.items()}
        if len(set(counts.values())) > 1:
            raise ValueError(f"Row counts differ across tabs for {key}: {counts}")
        agent_tab = next((t for t in tabs if (parts[t].Role == "Agent").all()), "")
        patient_tab = next((t for t in tabs if (parts[t].Role == "Patient").all()), "")
        if not agent_tab or not patient_tab or len(tabs) != 2:
            raise ValueError(f"Unexpected cross-tab pattern for {key}: {tabs}")
        a, p = parts[agent_tab], parts[patient_tab]
        triple_str = f"{a.Who[0]} | {a.Action[0]} | {a.To_Whom[0]} [{a.Source[0]}]"
        for i in range(len(a)):
            if a.Sentiment[i] != p.Sentiment[i]:
                raise ValueError(f"Unresolved sentiment conflict for {triple_str}: "
                                 f"{agent_tab}={a.Sentiment[i]}, {patient_tab}={p.Sentiment[i]}. "
                                 "Add the agreed label to SENTIMENT_CORRECTIONS.")
            rows.append({
                "Who": a.Who[i], "Action": a.Action[i], "To_Whom": a.To_Whom[i],
                "Source": a.Source[i], "Sentiment": a.Sentiment[i],
                "Subject_Node": agent_tab, "Object_Node": patient_tab,
                "Tabs": f"{agent_tab}; {patient_tab}",
                "Reconciliation": "Appeared in 2 tabs; kept once.",
                "Notes": " | ".join(dict.fromkeys(x for x in [a.Notes[i], p.Notes[i]] if x)),
                **{c: a[c][i] or p[c][i] for c in NOTE_COLS[1:]},
            })
        log.append({"Type": "Cross-tab duplicate merged", "Triple": triple_str,
                    "Tabs": f"{agent_tab}; {patient_tab}",
                    "Detail": f"{len(a)} instance(s) kept once (not {2 * len(a)})"})
    m = pd.DataFrame(rows)
    m.insert(0, "Triple_ID", [f"T{i:03d}" for i in range(1, len(m) + 1)])
    return m


def summary_frame(master: pd.DataFrame, nodes: list[str]) -> pd.DataFrame:
    out = []
    for node in nodes:
        for corpus in CORPORA:
            for role in ROLES:
                col = "Subject_Node" if role == "Agent" else "Object_Node"
                sub = master[(master[col] == node) & (master.Source == corpus)]
                n = len(sub)
                row = {"Node": node, "Corpus": corpus, "Role": role, "N": n}
                for s in SENTIMENTS:
                    row[s] = int((sub.Sentiment == s).sum())
                for s in SENTIMENTS:
                    row[f"% {s}"] = round(100 * row[s] / n, 1) if n else None
                out.append(row)
    return pd.DataFrame(out)


HDR_FONT = Font(name="Arial", bold=True)
BODY_FONT = Font(name="Arial")
HDR_FILL = PatternFill("solid", fgColor="D9E1F2")
WIDTHS = {"Triple_ID": 9, "Who": 22, "Action": 14, "To_Whom": 28, "Source": 9, "Role": 9,
          "Sentiment": 10, "Subject_Node": 18, "Object_Node": 18, "Tabs": 24,
          "Reconciliation": 30, "Notes": 40, "Second_Reviewer_Notes": 40,
          "Original_Sentence": 50, "Context": 26}


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")


def write_table(ws, df: pd.DataFrame):
    cols = list(df.columns)
    ws.append(cols)
    for rec in df.itertuples(index=False):
        ws.append(list(rec))
    style_header(ws, 1, len(cols))
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = BODY_FONT
    for i, c in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(i)].width = WIDTHS.get(c, 14)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def write_workbook(dst: Path, master: pd.DataFrame, tabs: pd.DataFrame, nodes: list[str],
                   log: list, include_notes: bool):
    wb = Workbook()
    wb.remove(wb.active)
    note_cols = NOTE_COLS if include_notes else []

    # ---- Master (values) ----
    ws = wb.create_sheet("Master")
    m = master[[c for c in master.columns if c not in NOTE_COLS] + note_cols]
    write_table(ws, m)
    ws.freeze_panes = "B2"
    cols, last = list(m.columns), ws.max_row
    L = {c: get_column_letter(cols.index(c) + 1) for c in cols}
    rng = lambda c: f"Master!${L[c]}$2:${L[c]}${last}"

    # ---- Summary (formulas on Master) ----
    ss = wb.create_sheet("Summary")
    ss["A1"] = "Sentiment of shared nodes by corpus and role (manually validated labels)"
    ss["A1"].font = Font(name="Arial", bold=True, size=12)
    ss["A2"] = ("Agent = node is the subject of the triple; Patient = node is the object. "
                "Counts are live formulas on the Master tab. A triple linking two shared nodes "
                "counts once for each node, so do not sum N across nodes (see Corpus_Totals).")
    ss["A2"].font = Font(name="Arial", italic=True)
    hdr = ["Node", "Corpus", "Role", "N", "Positive", "Negative", "Neutral",
           "% Positive", "% Negative", "% Neutral"]
    HR = 4
    for i, h in enumerate(hdr, 1):
        ss.cell(row=HR, column=i, value=h)
    style_header(ss, HR, len(hdr))
    r = HR + 1
    for node in nodes:
        for corpus in CORPORA:
            for role in ROLES:
                ss.cell(r, 1, node); ss.cell(r, 2, corpus); ss.cell(r, 3, role)
                for j in range(3):
                    col = get_column_letter(5 + j)
                    ss.cell(r, 5 + j, (
                        f'=IF($C{r}="Agent",'
                        f'COUNTIFS({rng("Subject_Node")},$A{r},{rng("Source")},$B{r},{rng("Sentiment")},{col}${HR}),'
                        f'COUNTIFS({rng("Object_Node")},$A{r},{rng("Source")},$B{r},{rng("Sentiment")},{col}${HR}))'))
                ss.cell(r, 4, f"=SUM(E{r}:G{r})")
                for j in range(3):
                    c = ss.cell(r, 8 + j, f'=IF($D{r}=0,"",{get_column_letter(5 + j)}{r}/$D{r})')
                    c.number_format = "0.0%"
                for c in range(1, 11):
                    ss.cell(r, c).font = BODY_FONT
                r += 1
    for i, w in enumerate([20, 9, 9, 6, 9, 9, 9, 11, 11, 11], 1):
        ss.column_dimensions[get_column_letter(i)].width = w
    ss.freeze_panes = f"D{HR + 1}"
    ss.auto_filter.ref = f"A{HR}:J{r - 1}"

    # ---- Corpus totals (deduplicated) ----
    ct = wb.create_sheet("Corpus_Totals")
    ct["A1"] = "Sentiment by corpus, all shared-node triples (deduplicated: each triple instance counted once)"
    ct["A1"].font = Font(name="Arial", bold=True, size=12)
    hdr2 = ["Corpus", "N", "Positive", "Negative", "Neutral", "% Positive", "% Negative", "% Neutral"]
    for i, h in enumerate(hdr2, 1):
        ct.cell(3, i, h)
    style_header(ct, 3, len(hdr2))
    for k, corpus in enumerate(CORPORA + ["Total"]):
        rr = 4 + k
        ct.cell(rr, 1, corpus)
        for j in range(3):
            colL = get_column_letter(3 + j)
            ct.cell(rr, 3 + j, f"=SUM({colL}4:{colL}5)" if corpus == "Total" else
                    f'=COUNTIFS({rng("Source")},$A{rr},{rng("Sentiment")},{colL}$3)')
        ct.cell(rr, 2, f"=SUM(C{rr}:E{rr})")
        for j in range(3):
            c = ct.cell(rr, 6 + j, f'=IF($B{rr}=0,"",{get_column_letter(3 + j)}{rr}/$B{rr})')
            c.number_format = "0.0%"
        for c in range(1, 9):
            ct.cell(rr, c).font = HDR_FONT if corpus == "Total" else BODY_FONT
    for i, w in enumerate([10, 6, 9, 9, 9, 11, 11, 11], 1):
        ct.column_dimensions[get_column_letter(i)].width = w

    # ---- Reconciliation log ----
    order = {"Sentiment corrected": 0, "Wording merged": 1, "Cross-tab duplicate merged": 2}
    rl = wb.create_sheet("Reconciliation")
    write_table(rl, pd.DataFrame(sorted(log, key=lambda e: order.get(e["Type"], 9)),
                                 columns=["Type", "Triple", "Tabs", "Detail"]))
    for i, w in enumerate([28, 50, 34, 80], 1):
        rl.column_dimensions[get_column_letter(i)].width = w

    # ---- Node tabs (corrected labels) ----
    for node in nodes:
        t = tabs[tabs.Tab == node][CORE + note_cols]
        write_table(wb.create_sheet(node[:31]), t)

    wb.save(dst)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", default=f"{SENT_DIR}/spo_validated_final.xlsx")
    ap.add_argument("--outdir", default=SENT_DIR)
    ap.add_argument("--internal", action="store_true",
                    help="also write *_internal.xlsx/.csv with notes and quotes (git-ignored)")
    args = ap.parse_args()
    src, outdir = Path(args.input), Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    nodes, tabs, log = load_tabs(src)
    master = build_master(tabs, log)
    summary = summary_frame(master, nodes)
    public = master.drop(columns=NOTE_COLS)

    write_workbook(outdir / "spo_validated_public.xlsx", master, tabs, nodes, log, include_notes=False)
    public.to_csv(outdir / "spo_validated_master.csv", index=False, encoding="utf-8-sig")
    summary.to_csv(outdir / "spo_validated_summary.csv", index=False, encoding="utf-8-sig")
    if args.internal:
        write_workbook(outdir / "spo_validated_master_internal.xlsx", master, tabs, nodes, log,
                       include_notes=True)
        master.to_csv(outdir / "spo_validated_master_internal.csv", index=False, encoding="utf-8-sig")

    print(f"Tab rows: {len(tabs)}  ->  Master rows (deduplicated): {len(master)}")
    print(master.groupby(["Source", "Sentiment"]).size().unstack(fill_value=0))
    print(f"Wrote public outputs to {outdir}" + (" (plus *_internal files)" if args.internal else ""))


if __name__ == "__main__":
    main()
