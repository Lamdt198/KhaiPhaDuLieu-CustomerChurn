# -*- coding: utf-8 -*-
"""
src/eda_stats.py
Kiem dinh thong ke (Chi-Square, Cramer's V, Mann-Whitney U, tuong quan Pearson)
va xuat 3 bieu do kham pha du lieu (EDA) do phan giai 300 DPI.
"""

import json
import os
import sys
import numpy as np
import pandas as pd
from scipy import stats
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

from src.data_loader import (
    load_prepared_data,
    NUMERIC_FEATURES,
    CATEGORICAL_FEATURES,
    BINARY_NUMERIC_FEATURES,
)

FIG_DIR = os.path.join(BASE_DIR, "reports", "figures")
MODEL_DIR = os.path.join(BASE_DIR, "models")
os.makedirs(FIG_DIR, exist_ok=True)
os.makedirs(MODEL_DIR, exist_ok=True)

# Cau hinh giao dien bieu do dong nhat he mau HUBT
plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["axes.edgecolor"] = "#CBD5E1"
plt.rcParams["axes.linewidth"] = 0.8
plt.rcParams["grid.color"] = "#E2E8F0"
plt.rcParams["grid.linestyle"] = "--"
plt.rcParams["grid.alpha"] = 0.7

NAVY = "#0F2544"
BLUE = "#1E3A8A"
TEAL = "#0F766E"
AMBER = "#D97706"
CRIMSON = "#B91C1C"
SLATE = "#475569"
LIGHT_BG = "#F8FAFC"


def cramers_v(x: pd.Series, y: pd.Series):
    ct = pd.crosstab(x, y)
    chi2, p, dof, _ = stats.chi2_contingency(ct)
    n = ct.sum().sum()
    min_dim = min(ct.shape) - 1
    v = np.sqrt(chi2 / (n * min_dim)) if min_dim > 0 else 0.0
    return float(chi2), float(p), int(dof), float(v)


def run_eda():
    _, df, X, y, audit = load_prepared_data()
    df_analysis = X.copy()
    df_analysis["Churn"] = y.values

    # 1. Kiem dinh Chi-Square & Cramer's V cho bien phan loai + bien nhi phan
    cat_cols_all = CATEGORICAL_FEATURES + BINARY_NUMERIC_FEATURES
    chi2_results = []
    for col in cat_cols_all:
        chi2_val, p_val, dof, v_val = cramers_v(df_analysis[col], df_analysis["Churn"])
        chi2_results.append({
            "feature": col,
            "chi2": round(chi2_val, 2),
            "dof": dof,
            "p_value": float(f"{p_val:.3e}"),
            "cramers_v": round(v_val, 4),
            "significant": bool(p_val < 0.05),
        })
    chi2_df = pd.DataFrame(chi2_results).sort_values("cramers_v", ascending=False)

    # 2. Kiem dinh Mann-Whitney U cho bien so lien tuc
    num_results = []
    for col in NUMERIC_FEATURES:
        g0 = df_analysis.loc[df_analysis["Churn"] == 0, col]
        g1 = df_analysis.loc[df_analysis["Churn"] == 1, col]
        u_stat, p_val = stats.mannwhitneyu(g0, g1, alternative="two-sided")
        num_results.append({
            "feature": col,
            "median_non_churn": round(float(g0.median()), 2),
            "median_churn": round(float(g1.median()), 2),
            "mean_non_churn": round(float(g0.mean()), 2),
            "mean_churn": round(float(g1.mean()), 2),
            "u_stat": round(float(u_stat), 1),
            "p_value": float(f"{p_val:.3e}"),
            "significant": bool(p_val < 0.05),
        })

    # 3. Ty le Churn theo cac nhom chu chot
    def group_churn_rate(col_name):
        g = df_analysis.groupby(col_name, observed=False)["Churn"].agg(["count", "mean"]).reset_index()
        return [
            {
                "category": str(r[col_name]),
                "count": int(r["count"]),
                "churn_rate_pct": round(float(r["mean"]) * 100, 2),
            }
            for _, r in g.iterrows()
        ]

    breakdown = {
        "Contract": group_churn_rate("Contract"),
        "InternetService": group_churn_rate("InternetService"),
        "PaymentMethod": group_churn_rate("PaymentMethod"),
        "TenureGroup": group_churn_rate("TenureGroup"),
        "HasProtectionBundle": group_churn_rate("HasProtectionBundle"),
        "IsAutoPayment": group_churn_rate("IsAutoPayment"),
    }

    # Ma tran tuong quan Pearson bien so
    corr_mat = df_analysis[NUMERIC_FEATURES + ["Churn"]].corr().round(3)

    eda_summary = {
        "audit": audit,
        "chi2_top12": chi2_df.head(12).to_dict(orient="records"),
        "chi2_all": chi2_df.to_dict(orient="records"),
        "mann_whitney": num_results,
        "group_breakdown": breakdown,
        "pearson_tenure_totalcharges": float(corr_mat.loc["tenure", "TotalCharges"]),
        "pearson_monthly_totalcharges": float(corr_mat.loc["MonthlyCharges", "TotalCharges"]),
    }

    with open(os.path.join(MODEL_DIR, "eda_stats.json"), "w", encoding="utf-8") as f:
        json.dump(eda_summary, f, ensure_ascii=False, indent=2)

    # ==================== FIG 01: PHAN BO NHAN & TY LE CHURN THEO HOP DONG ====================
    fig, axes = plt.subplots(1, 2, figsize=(11.5, 4.6), dpi=300)
    fig.patch.set_facecolor("white")

    # Panel A: Donut chart phan bo nhan Churn
    ax = axes[0]
    counts = [audit["non_churn_count"], audit["churn_count"]]
    labels_pie = [
        f"Ở lại (Non-Churn)\n{counts[0]:,} KH ({counts[0]/audit['n_rows']*100:.1f}%)",
        f"Rời bỏ (Churn)\n{counts[1]:,} KH ({counts[1]/audit['n_rows']*100:.1f}%)",
    ]
    wedges, texts = ax.pie(
        counts,
        labels=labels_pie,
        colors=[BLUE, CRIMSON],
        startangle=90,
        wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2),
        textprops=dict(fontsize=9.5, color="#0F172A", fontweight="bold"),
    )
    ax.text(0, 0, f"N = {audit['n_rows']:,}\nkhách hàng", ha="center", va="center",
            fontsize=10.5, fontweight="bold", color=NAVY)
    ax.set_title("A. Phân bố biến mục tiêu (Churn)", fontsize=11.5, fontweight="bold", color=NAVY, pad=12)

    # Panel B: Ty le Churn theo loai hop dong
    ax = axes[1]
    order_c = ["Month-to-month", "One year", "Two year"]
    labels_c = ["Theo tháng\n(Month-to-month)", "1 năm\n(One year)", "2 năm\n(Two year)"]
    rates_c = [df_analysis.loc[df_analysis["Contract"] == c, "Churn"].mean() * 100 for c in order_c]
    counts_c = [int((df_analysis["Contract"] == c).sum()) for c in order_c]
    bars = ax.bar(labels_c, rates_c, color=[CRIMSON, AMBER, TEAL], width=0.55, edgecolor="none")
    for bar, r, cnt in zip(bars, rates_c, counts_c):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            r + 1.4,
            f"{r:.1f}%\n(n={cnt:,})",
            ha="center", va="bottom", fontsize=9.2, fontweight="bold", color="#0F172A"
        )
    ax.set_ylim(0, 54)
    ax.set_ylabel("Tỷ lệ rời bỏ (%)", fontsize=10, fontweight="bold", color=SLATE)
    ax.set_title("B. Tỷ lệ rời bỏ theo loại hợp đồng (Contract)", fontsize=11.5, fontweight="bold", color=NAVY, pad=12)
    ax.yaxis.grid(True)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    plt.tight_layout(pad=2.0)
    plt.savefig(os.path.join(FIG_DIR, "fig01_churn_distribution.png"), dpi=300, bbox_inches="tight")
    plt.close(fig)

    # ==================== FIG 02: CAC YEU TO ANH HUONG CHINH (2x2) ====================
    fig, axes = plt.subplots(2, 2, figsize=(12.0, 8.2), dpi=300)
    fig.patch.set_facecolor("white")

    # (0,0) Phan phoi Tenure
    ax = axes[0, 0]
    sns.kdeplot(df_analysis.loc[df_analysis["Churn"] == 0, "tenure"], ax=ax, fill=True,
                color=BLUE, alpha=0.28, linewidth=2, label="Ở lại (Non-Churn)")
    sns.kdeplot(df_analysis.loc[df_analysis["Churn"] == 1, "tenure"], ax=ax, fill=True,
                color=CRIMSON, alpha=0.28, linewidth=2, label="Rời bỏ (Churn)")
    ax.set_title("A. Phân phối thời gian gắn bó (tenure - tháng)", fontsize=11, fontweight="bold", color=NAVY)
    ax.set_xlabel("Số tháng sử dụng dịch vụ (tenure)", fontsize=9.5)
    ax.set_ylabel("Mật độ xác suất", fontsize=9.5)
    ax.legend(frameon=True, facecolor="white", fontsize=9)
    ax.grid(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    # (0,1) Phan phoi MonthlyCharges
    ax = axes[0, 1]
    sns.kdeplot(df_analysis.loc[df_analysis["Churn"] == 0, "MonthlyCharges"], ax=ax, fill=True,
                color=BLUE, alpha=0.28, linewidth=2, label="Ở lại (Non-Churn)")
    sns.kdeplot(df_analysis.loc[df_analysis["Churn"] == 1, "MonthlyCharges"], ax=ax, fill=True,
                color=CRIMSON, alpha=0.28, linewidth=2, label="Rời bỏ (Churn)")
    ax.set_title("B. Phân phối cước phí hàng tháng (MonthlyCharges)", fontsize=11, fontweight="bold", color=NAVY)
    ax.set_xlabel("Cước hàng tháng (USD)", fontsize=9.5)
    ax.set_ylabel("Mật độ xác suất", fontsize=9.5)
    ax.legend(frameon=True, facecolor="white", fontsize=9)
    ax.grid(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    # (1,0) Ty le Churn theo dich vu Internet & Bao mat/Ho tro
    ax = axes[1, 0]
    inet_order = ["Fiber optic", "DSL", "No"]
    inet_labels = ["Cáp quang\n(Fiber optic)", "DSL", "Không dùng\nInternet"]
    inet_rates = [df_analysis.loc[df_analysis["InternetService"] == k, "Churn"].mean() * 100 for k in inet_order]
    bars = ax.bar(inet_labels, inet_rates, color=[CRIMSON, AMBER, TEAL], width=0.52)
    for b, r in zip(bars, inet_rates):
        ax.text(b.get_x() + b.get_width()/2, r + 1.2, f"{r:.1f}%", ha="center", va="bottom",
                fontsize=9.5, fontweight="bold", color="#0F172A")
    ax.set_ylim(0, 52)
    ax.set_ylabel("Tỷ lệ rời bỏ (%)", fontsize=9.5)
    ax.set_title("C. Tỷ lệ rời bỏ theo hạ tầng Internet", fontsize=11, fontweight="bold", color=NAVY)
    ax.yaxis.grid(True)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    # (1,1) Ty le Churn theo phuong thuc thanh toan
    ax = axes[1, 1]
    pm_stats = (
        df_analysis.groupby("PaymentMethod")["Churn"].mean().sort_values(ascending=True) * 100
    )
    pm_map = {
        "Electronic check": "Electronic check",
        "Mailed check": "Mailed check",
        "Bank transfer (automatic)": "Bank transfer (auto)",
        "Credit card (automatic)": "Credit card (auto)",
    }
    y_lbls = [pm_map.get(k, k) for k in pm_stats.index]
    colors_pm = [TEAL, TEAL, AMBER, CRIMSON]
    bars = ax.barh(y_lbls, pm_stats.values, color=colors_pm, height=0.55)
    for b, r in zip(bars, pm_stats.values):
        ax.text(r + 0.9, b.get_y() + b.get_height()/2, f"{r:.1f}%", va="center",
                fontsize=9.5, fontweight="bold", color="#0F172A")
    ax.set_xlim(0, 55)
    ax.set_xlabel("Tỷ lệ rời bỏ (%)", fontsize=9.5)
    ax.set_title("D. Tỷ lệ rời bỏ theo phương thức thanh toán", fontsize=11, fontweight="bold", color=NAVY)
    ax.xaxis.grid(True)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    plt.tight_layout(pad=2.0)
    plt.savefig(os.path.join(FIG_DIR, "fig02_eda_drivers.png"), dpi=300, bbox_inches="tight")
    plt.close(fig)

    # ==================== FIG 03: TUONG QUAN PEARSON & CRAMER'S V ====================
    fig, axes = plt.subplots(1, 2, figsize=(12.8, 5.2), dpi=300, gridspec_kw={"width_ratios": [1.05, 1.0]})
    fig.patch.set_facecolor("white")

    # Panel A: Pearson correlation heatmap
    ax = axes[0]
    sns.heatmap(
        corr_mat,
        annot=True,
        fmt=".2f",
        cmap="RdBu_r",
        center=0,
        vmin=-1,
        vmax=1,
        square=True,
        linewidths=0.6,
        cbar_kws={"shrink": 0.8},
        annot_kws={"size": 8.5},
        ax=ax,
    )
    ax.set_title("A. Ma trận tương quan Pearson (Biến số & Churn)", fontsize=11, fontweight="bold", color=NAVY, pad=10)
    ax.tick_params(axis="x", rotation=35, labelsize=8.5)
    ax.tick_params(axis="y", rotation=0, labelsize=8.5)

    # Panel B: Top 10 Cramer's V
    ax = axes[1]
    top_v = chi2_df.head(10).iloc[::-1]
    bar_colors = [CRIMSON if v >= 0.30 else (AMBER if v >= 0.20 else BLUE) for v in top_v["cramers_v"]]
    bars = ax.barh(top_v["feature"], top_v["cramers_v"], color=bar_colors, height=0.62)
    for b, v in zip(bars, top_v["cramers_v"]):
        ax.text(v + 0.008, b.get_y() + b.get_height()/2, f"{v:.3f}", va="center",
                fontsize=8.8, fontweight="bold", color="#0F172A")
    ax.set_xlim(0, max(top_v["cramers_v"]) + 0.08)
    ax.set_xlabel("Hệ số liên kết Cramér's V (Kiểm định Chi-Square)", fontsize=9.5)
    ax.set_title("B. Top 10 biến phân loại liên kết mạnh nhất với Churn", fontsize=11, fontweight="bold", color=NAVY, pad=10)
    ax.xaxis.grid(True)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    plt.tight_layout(pad=2.0)
    plt.savefig(os.path.join(FIG_DIR, "fig03_correlation_cramersv.png"), dpi=300, bbox_inches="tight")
    plt.close(fig)

    print("EDA stats saved to models/eda_stats.json")
    print("Top 5 Cramer's V:")
    print(chi2_df.head(5).to_string(index=False))
    print("\nMann-Whitney U results:")
    print(pd.DataFrame(num_results)[["feature", "median_non_churn", "median_churn", "p_value"]].to_string(index=False))


if __name__ == "__main__":
    run_eda()
