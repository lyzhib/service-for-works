import yaml
import pandas as pd
from pathlib import Path
from tqdm import tqdm

from extract_text import extract_text
from local_eval import evaluate

# ================= LOAD CONFIG =================

with open("config.yaml", "r", encoding="utf-8") as f:
    config = yaml.safe_load(f)

REPORTS_DIR = Path(config["paths"]["reports_dir"])
OUTPUT_DIR = Path(config["paths"]["output_dir"])
OUTPUT_DIR.mkdir(exist_ok=True)

TOLERANCE = config["scoring"]["tolerance"]

# ================= LOAD CRITERIA =================
# ⚠️ 这里假设：Excel 中已经只剩“绿色 критерии”，并且顺序正确

criteria_df = pd.read_excel(config["sheets"]["criteria_file"])
criteria_df = criteria_df.iloc[:, :3]
criteria_df.columns = ["criterion", "description", "max_score"]

criteria = []
for _, row in criteria_df.iterrows():
    criteria.append({
        "name": str(row["criterion"]),
        "description": str(row["description"]),
        "max_score": float(row["max_score"]),
    })

print(f"✅ Loaded criteria: {len(criteria)}")

# ================= LOAD HUMAN SCORES =================

human_df = pd.read_excel(config["sheets"]["human_scores_file"])
human_df.iloc[:, 0] = human_df.iloc[:, 0].astype(str)

report_id_col = human_df.columns[0]

print(f"✅ Loaded human scores: {len(human_df)} reports")

# ================= PIPELINE =================

rows = []

for report_path in tqdm(list(REPORTS_DIR.iterdir()), desc="Reports"):
    report_id = report_path.stem

    # ---- extract text ----
    try:
        text = extract_text(report_path).strip()
    except Exception as e:
        print(f"⚠️ {report_path.name}: {e}")
        continue

    if not text:
        print(f"⚠️ {report_path.name}: empty text")
        continue

    # ---- find human row ----
    hrow = human_df[human_df[report_id_col] == report_id]
    if hrow.empty:
        print(f"⚠️ No human scores for report {report_id}")
        continue

    hrow = hrow.iloc[0]

    # ---- AI evaluation ----
    ai_scores = evaluate(text, criteria)

    # ---- compare by INDEX (核心！！！) ----
    for idx, crit in enumerate(criteria):
        col_idx = idx + 1  # 0 is report_id

        if col_idx >= len(hrow):
            continue

        human_score_raw = hrow.iloc[col_idx]

        # 1️⃣ 跳过空值
        if pd.isna(human_score_raw):
            continue

        # 2️⃣ 强制尝试转成数字
        try:
            human_score = float(human_score_raw)
        except (ValueError, TypeError):
            # 不是评分列（如 PDF / Word / комментарий）
            continue

        ai_score = ai_scores.get(crit["name"])

        # 3️⃣ AI 分数兜底
        try:
            ai_score = float(ai_score)
        except (ValueError, TypeError):
            continue

        diff = abs(human_score - ai_score)
        match = diff <= TOLERANCE

        rows.append({
            "report_id": report_id,
            "report_file": report_path.name,
            "criterion_index": idx + 1,
            "criterion_name": crit["name"],
            "human_score": float(human_score),
            "ai_score": float(ai_score),
            "diff": diff,
            "match": match,
        })

# ================= SAVE =================

df = pd.DataFrame(rows)

if df.empty:
    print("❌ No valid results produced")
else:
    out = OUTPUT_DIR / "results.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="details", index=False)

        summary = (
            df.groupby("criterion_name")
            .agg(
                mean_diff=("diff", "mean"),
                match_rate=("match", "mean"),
                count=("match", "count"),
            )
            .reset_index()
        )

        summary.to_excel(writer, sheet_name="summary", index=False)

    print(f"🎉 DONE! Results saved to {out}")