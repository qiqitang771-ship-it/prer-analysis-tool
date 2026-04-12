import pandas as pd
import numpy as np
from io import BytesIO


# =========================
# 你的原函数（精简版保留）
# =========================
def to_numeric_safe(df, cols):
    for c in cols:
        if c in df.columns:
            df[c] = (
                df[c].astype(str)
                .str.replace(",", "")
                .str.replace("—", "")
                .str.replace("NA", "")
                .str.strip()
            )
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def ensure_columns(df, cols):
    for c in cols:
        if c not in df.columns:
            df[c] = None
    return df


# =========================
# 数据读取
# =========================
def load_data(eff_file, saf_file):

    df_eff = pd.read_excel(eff_file)
    df_saf = pd.read_excel(saf_file)

    eff_map = {
        "文献编号": "doc_id",
        "组别": "group",
        "有效性指标": "outcome",
        "访视点": "timepoint",
        "器械": "device",
        "数据类型": "type",
        "样本量": "n",
        "均值": "mean",
        "标准差": "sd",
        "发生例数": "event"
    }

    saf_map = {
        "文献编号": "doc_id",
        "组别": "group",
        "器械": "device",
        "安全性指标分类": "category",
        "安全性指标": "outcome",
        "样本量": "n",
        "发生例数": "event"
    }

    df_eff = df_eff.rename(columns=eff_map)
    df_saf = df_saf.rename(columns=saf_map)

    df_eff = to_numeric_safe(df_eff, ["n", "mean", "sd", "event"])
    df_saf = to_numeric_safe(df_saf, ["n", "event"])

    return df_eff, df_saf


# =========================
# pooled计算
# =========================
def pooled_continuous(df):
    n = df["n"].sum()
    mean = np.sum(df["mean"] * df["n"]) / n
    sd = np.sqrt(np.sum((df["n"]-1)*df["sd"]**2) / (n-1))
    return n, mean, sd


def pooled_binary(df):
    n = df["n"].sum()
    event = df["event"].sum()
    return n, event, event/n if n else 0


# =========================
# 主流程
# =========================
def process_all(eff_file, saf_file, merge_eff=True, merge_saf=True):

    df_eff, df_saf = load_data(eff_file, saf_file)

    # ===== 有效性 =====
    eff_results = {}

    for outcome in df_eff["outcome"].dropna().unique():
        sub = df_eff[df_eff["outcome"] == outcome].copy()

        table = sub.copy()

        if merge_eff:
            n, mean, sd = pooled_continuous(sub.dropna(subset=["n","mean","sd"]))

            merge_row = pd.DataFrame([{
                "doc_id": "合并计算",
                "group": "合并计算",
                "device": "合并计算",
                "n": n,
                "mean": mean,
                "sd": sd
            }])

            table = pd.concat([table, merge_row], ignore_index=True)

        eff_results[outcome] = table

    # ===== 安全性 =====
    saf_results = {}

    for cat in df_saf["category"].dropna().unique():
        sub = df_saf[df_saf["category"] == cat].copy()

        table = sub.copy()

        if merge_saf:
            n, event, rate = pooled_binary(sub)

            merge_row = pd.DataFrame([{
                "doc_id": "合并计算",
                "group": "合并计算",
                "device": "合并计算",
                "outcome": f"{cat}总体",
                "n": n,
                "event": event
            }])

            table = pd.concat([table, merge_row], ignore_index=True)

        saf_results[cat] = table

    return eff_results, saf_results
