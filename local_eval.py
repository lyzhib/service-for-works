import re
from typing import List, Dict

def evaluate(report_text: str, criteria: List[Dict]) -> Dict[str, float]:
    """
    永远返回每个 criterion 的分数（0 ~ max_score）
    """
    results = {}

    text_len = len(report_text)

    for crit in criteria:
        max_score = crit["max_score"]

        # 一个非常稳妥的“伪评分逻辑”
        if text_len == 0:
            score = 0
        else:
            score = min(max_score, max(0, text_len // 1000))

        results[crit["name"]] = float(score)

    return results