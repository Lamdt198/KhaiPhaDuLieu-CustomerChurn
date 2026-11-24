# -*- coding: utf-8 -*-
"""
src/train_pipeline.py
Pipeline chong ro ri du lieu (Anti-Leakage), huan luyen 5 mo hinh voi
Stratified 5-Fold CV + GridSearchCV, Ablation Study xu ly lech lop,
toi uu nguong nhay chi phi T* tren OOF CV, trich xuat luat CART,
giai thich SHAP Global/Local, phan tich loi chuyen sau va doi chung ROI.
"""

import json
import os
import sys
import time
import warnings
warnings.filterwarnings("ignore")

import joblib
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.compose import ColumnTransformer
from sklearn.impute import SimpleImputer
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (
    accuracy_score,
    average_precision_score,
    brier_score_loss,
    confusion_matrix,
    f1_score,
    precision_recall_curve,
    precision_score,
    recall_score,
    roc_auc_score,
    roc_curve,
)
from sklearn.model_selection import GridSearchCV, StratifiedKFold, cross_val_predict, train_test_split
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.tree import DecisionTreeClassifier, _tree, plot_tree
from sklearn.ensemble import RandomForestClassifier
from xgboost import XGBClassifier
from lightgbm import LGBMClassifier
import shap

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

from src.data_loader import (
    BINARY_NUMERIC_FEATURES,
    CATEGORICAL_FEATURES,
    NUMERIC_FEATURES,
    load_prepared_data,
)

FIG_DIR = os.path.join(BASE_DIR, "reports", "figures")
MODEL_DIR = os.path.join(BASE_DIR, "models")
DATA_DIR = os.path.join(BASE_DIR, "data")
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
BLUE_MID = "#2563EB"
TEAL = "#0F766E"
AMBER = "#D97706"
CRIMSON = "#B91C1C"
SLATE = "#475569"

# Ma tran chi phi bat doi xung (don vi: USD / khach hang)
COST_FP = 100.0   # Chi phi uu dai / cham soc cho ca du doan nham
COST_TP = 100.0   # Chi phi uu dai / cham soc giu chan khach co nguy co roi bo
COST_FN = 500.0   # Ton that gia tri vong doi (CLV) khi bo sot khach roi bo
COST_TN = 0.0     # Khong can thiep, khach van o lai


def make_preprocessor() -> ColumnTransformer:
    num_pipe = Pipeline([
        ("imputer", SimpleImputer(strategy="median")),
        ("scaler", StandardScaler()),
    ])
    bin_pipe = Pipeline([
        ("imputer", SimpleImputer(strategy="most_frequent")),
    ])
    cat_pipe = Pipeline([
        ("imputer", SimpleImputer(strategy="most_frequent")),
        ("ohe", OneHotEncoder(handle_unknown="ignore", sparse_output=False)),
    ])
    return ColumnTransformer(
        transformers=[
            ("num", num_pipe, NUMERIC_FEATURES),
            ("bin", bin_pipe, BINARY_NUMERIC_FEATURES),
            ("cat", cat_pipe, CATEGORICAL_FEATURES),
        ],
        remainder="drop",
    )


def get_clean_feature_names(prep: ColumnTransformer) -> list:
    raw_names = prep.get_feature_names_out()
    clean = []
    for n in raw_names:
        for prefix in ("num__", "bin__", "cat__"):
            if n.startswith(prefix):
                n = n[len(prefix):]
        clean.append(n)
    return clean


def compute_metrics(y_true, y_prob, threshold=0.5):
    y_pred = (y_prob >= threshold).astype(int)
    tn, fp, fn, tp = confusion_matrix(y_true, y_pred).ravel()
    total_cost = fp * COST_FP + tp * COST_TP + fn * COST_FN + tn * COST_TN
    return {
        "threshold": round(float(threshold), 3),
        "accuracy": round(float(accuracy_score(y_true, y_pred)), 4),
        "precision": round(float(precision_score(y_true, y_pred, zero_division=0)), 4),
        "recall": round(float(recall_score(y_true, y_pred, zero_division=0)), 4),
        "f1": round(float(f1_score(y_true, y_pred, zero_division=0)), 4),
        "roc_auc": round(float(roc_auc_score(y_true, y_prob)), 4),
        "pr_auc": round(float(average_precision_score(y_true, y_prob)), 4),
        "brier": round(float(brier_score_loss(y_true, y_prob)), 4),
        "tn": int(tn),
        "fp": int(fp),
        "fn": int(fn),
        "tp": int(tp),
        "total_cost_usd": round(float(total_cost), 2),
        "avg_cost_per_customer_usd": round(float(total_cost / len(y_true)), 2),
    }


def extract_cart_rules(tree_clf, feature_names, total_samples):
    """Trich xuat danh sach luat IF-THEN tu la cua cay quyet dinh CART."""
    tree_ = tree_clf.tree_
    feature_name = [
        feature_names[i] if i != _tree.TREE_UNDEFINED else "undefined!"
        for i in tree_.feature
    ]
    rules = []

    def recurse(node, conditions):
        if tree_.feature[node] != _tree.TREE_UNDEFINED:
            name = feature_name[node]
            threshold = tree_.threshold[node]
            recurse(tree_.children_left[node], conditions + [f"{name} <= {threshold:.2f}"])
            recurse(tree_.children_right[node], conditions + [f"{name} > {threshold:.2f}"])
        else:
            vals = tree_.value[node][0]
            # tree_.value with class_weight stores weighted counts; use n_node_samples for true support
            n_samp = int(tree_.n_node_samples[node])
            # Estimate class probability from weighted or unweighted node distribution
            prob_churn = float(vals[1] / vals.sum()) if vals.sum() > 0 else 0.0
            rules.append({
                "conditions": " AND ".join(conditions),
                "n_samples": n_samp,
                "support_pct": round(n_samp / total_samples * 100, 2),
                "churn_prob_pct": round(prob_churn * 100, 2),
                "predicted_class": "Churn (Rời bỏ)" if prob_churn >= 0.5 else "Non-Churn (Ở lại)",
            })

    recurse(0, [])
    rules.sort(key=lambda r: r["churn_prob_pct"], reverse=True)
    return rules


def run_training():
    t0_all = time.time()
    raw_df, clean_df, X, y, audit = load_prepared_data()

    # 1. Stratified Split 80/20 ngay tu dau (chong ro ri du lieu)
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.20, stratify=y, random_state=42
    )
    print(f"Split complete: Train={len(X_train)} (churn={y_train.mean()*100:.2f}%), "
          f"Test={len(X_test)} (churn={y_test.mean()*100:.2f}%)")

    # Luu file CSV mau (100 KH tu tap Test khong co nhan Churn) de upload thu tren Streamlit
    test_raw_sample = raw_df.loc[X_test.index[:100]].drop(columns=["Churn"]).reset_index(drop=True)
    test_raw_sample.to_csv(os.path.join(DATA_DIR, "sample_test_customers.csv"), index=False)

    cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
    pos_ratio = float((y_train == 0).sum() / (y_train == 1).sum())

    # 2. Khong gian 5 mo hinh doi chung (2 Baselines + 3 Ensembles)
    # Su dung xac suat khong lam meo (unweighted) lam goc de bao toan hieu chuan xac suat (Brier Score)
    # va ket hop toi uu nguong chi phi; dong thoi doi chung voi class_weight='balanced' trong Ablation Study.
    model_specs = {
        "Logistic Regression": (
            LogisticRegression(max_iter=1000, random_state=42),
            {"clf__C": [0.05, 0.2, 1.0, 5.0]},
        ),
        "Decision Tree (CART)": (
            DecisionTreeClassifier(random_state=42),
            {"clf__max_depth": [3, 4, 5, 6, 8], "clf__min_samples_leaf": [20, 40, 60]},
        ),
        "Random Forest": (
            RandomForestClassifier(n_estimators=250, random_state=42, n_jobs=-1),
            {"clf__max_depth": [6, 8, 10], "clf__min_samples_leaf": [5, 10, 20]},
        ),
        "XGBoost": (
            XGBClassifier(
                random_state=42,
                eval_metric="logloss",
                subsample=0.8,
                colsample_bytree=0.8,
                n_jobs=-1,
            ),
            {
                "clf__n_estimators": [120, 200],
                "clf__max_depth": [3, 4],
                "clf__learning_rate": [0.03, 0.06],
            },
        ),
        "LightGBM": (
            LGBMClassifier(
                random_state=42,
                subsample=0.8,
                colsample_bytree=0.8,
                verbose=-1,
                n_jobs=-1,
            ),
            {
                "clf__n_estimators": [120, 200],
                "clf__max_depth": [3, 4],
                "clf__learning_rate": [0.03, 0.06],
                "clf__num_leaves": [15, 31],
            },
        ),
    }

    trained_models = {}
    test_probs = {}
    cv_summary = {}
    test_metrics_05 = {}
    best_params_all = {}
    train_times = {}

    for name, (est, grid) in model_specs.items():
        t_start = time.time()
        pipe = Pipeline([("prep", make_preprocessor()), ("clf", est)])
        gs = GridSearchCV(
            pipe,
            param_grid=grid,
            scoring="roc_auc",
            cv=cv,
            n_jobs=-1,
            refit=True,
        )
        gs.fit(X_train, y_train)
        elapsed = round(time.time() - t_start, 2)

        best_pipe = gs.best_estimator_
        y_prob_test = best_pipe.predict_proba(X_test)[:, 1]
        m = compute_metrics(y_test, y_prob_test, threshold=0.50)

        trained_models[name] = best_pipe
        test_probs[name] = y_prob_test
        cv_summary[name] = round(float(gs.best_score_), 4)
        test_metrics_05[name] = m
        best_params_all[name] = {k.replace("clf__", ""): v for k, v in gs.best_params_.items()}
        train_times[name] = elapsed

        print(f"[{name:20s}] CV AUC={gs.best_score_:.4f} | Test AUC={m['roc_auc']:.4f} | "
              f"PR-AUC={m['pr_auc']:.4f} | F1={m['f1']:.4f} | Recall={m['recall']:.4f} | "
              f"Brier={m['brier']:.4f} ({elapsed:.1f}s)")

    # Chon mo hinh tot nhat theo CV ROC-AUC va mo hinh Tree Ensemble tot nhat (cho SHAP)
    best_model_name = max(cv_summary, key=cv_summary.get)
    tree_candidates = ["Random Forest", "XGBoost", "LightGBM"]
    best_tree_name = max(tree_candidates, key=lambda k: cv_summary[k])
    print(f"\nBest overall model (by 5-Fold CV ROC-AUC): {best_model_name} (CV AUC={cv_summary[best_model_name]:.4f})")
    print(f"Best tree ensemble (for TreeSHAP): {best_tree_name} (CV AUC={cv_summary[best_tree_name]:.4f})")

    # 3. Ablation Study xu ly mat can bang lop & Toi uu nguong chi phi T* tren OOF CV
    # Tinh OOF probabilities tren X_train cho best_model va best_tree de tim T* khong ro ri tap Test
    oof_probs_best = cross_val_predict(
        trained_models[best_model_name], X_train, y_train, cv=cv, method="predict_proba", n_jobs=-1
    )[:, 1]

    thresholds = np.linspace(0.05, 0.85, 81)
    oof_costs = []
    for thr in thresholds:
        oof_m = compute_metrics(y_train, oof_probs_best, threshold=thr)
        oof_costs.append(oof_m["avg_cost_per_customer_usd"])
    best_idx = int(np.argmin(oof_costs))
    optimal_threshold = round(float(thresholds[best_idx]), 2)
    print(f"Optimal Cost-Sensitive Threshold T* (on OOF Train CV): {optimal_threshold:.2f} "
          f"(Theoretical Bayes T* = {COST_FP / (COST_FN - COST_TP + COST_FP):.2f})")

    # Huan luyen phien ban Balanced Weight cua best_model de doi chung Ablation Study
    if best_model_name == "Logistic Regression":
        bal_clf = LogisticRegression(
            max_iter=1000, random_state=42, class_weight="balanced", **best_params_all[best_model_name]
        )
    elif best_model_name == "XGBoost":
        bal_clf = XGBClassifier(
            random_state=42, eval_metric="logloss", subsample=0.8, colsample_bytree=0.8,
            scale_pos_weight=pos_ratio, n_jobs=-1, **best_params_all[best_model_name]
        )
    elif best_model_name == "LightGBM":
        bal_clf = LGBMClassifier(
            random_state=42, subsample=0.8, colsample_bytree=0.8,
            class_weight="balanced", verbose=-1, n_jobs=-1, **best_params_all[best_model_name]
        )
    else:
        bal_clf = RandomForestClassifier(
            n_estimators=250, random_state=42, class_weight="balanced", n_jobs=-1,
            **best_params_all[best_model_name]
        )

    bal_pipe = Pipeline([("prep", make_preprocessor()), ("clf", bal_clf)])
    bal_pipe.fit(X_train, y_train)
    bal_probs_test = bal_pipe.predict_proba(X_test)[:, 1]

    ablation_study = {
        "Unweighted (T=0.50)": compute_metrics(y_test, test_probs[best_model_name], threshold=0.50),
        "Class-Weighted Balanced (T=0.50)": compute_metrics(y_test, bal_probs_test, threshold=0.50),
        f"Cost-Sensitive Optimal (T*={optimal_threshold:.2f})": compute_metrics(
            y_test, test_probs[best_model_name], threshold=optimal_threshold
        ),
    }

    # Danh gia tat ca 5 mo hinh tai nguong toi uu T* (de so sanh chi phi cong bang)
    test_metrics_tstar = {
        name: compute_metrics(y_test, prob, threshold=optimal_threshold)
        for name, prob in test_probs.items()
    }

    # 4. Doi chung hieu qua kinh te - van hanh (4 kich ban tren tap Test N=1409)
    n_test = len(y_test)
    n_test_churn = int(y_test.sum())

    # Kich ban 1: Khong can thiep (Do Nothing)
    cost_do_nothing = n_test_churn * COST_FN

    # Kich ban 2: Can thiep dai tra 100% khach hang (Blanket Campaign)
    cost_blanket = n_test * COST_FP

    # Kich ban 3: Luat thu cong Excel (Hop dong theo thang & Cuoc hang thang > 70 USD)
    manual_pred = (
        (X_test["Contract"] == "Month-to-month") & (X_test["MonthlyCharges"] > 70.0)
    ).astype(int)
    tn_m, fp_m, fn_m, tp_m = confusion_matrix(y_test, manual_pred).ravel()
    cost_manual = fp_m * COST_FP + tp_m * COST_TP + fn_m * COST_FN

    # Kich ban 4: Mo hinh Khai pha du lieu toi uu nguong T*
    opt_m = ablation_study[f"Cost-Sensitive Optimal (T*={optimal_threshold:.2f})"]
    cost_model_opt = opt_m["total_cost_usd"]

    roi_benchmark = [
        {
            "strategy": "Không can thiệp (Do Nothing)",
            "targeted_customers": 0,
            "churn_caught": 0,
            "churn_missed": n_test_churn,
            "total_cost_usd": round(cost_do_nothing, 2),
            "savings_vs_nothing_usd": 0.0,
            "savings_pct": 0.0,
        },
        {
            "strategy": "Ưu đãi đại trà 100% (Blanket)",
            "targeted_customers": n_test,
            "churn_caught": n_test_churn,
            "churn_missed": 0,
            "total_cost_usd": round(cost_blanket, 2),
            "savings_vs_nothing_usd": round(cost_do_nothing - cost_blanket, 2),
            "savings_pct": round((cost_do_nothing - cost_blanket) / cost_do_nothing * 100, 1),
        },
        {
            "strategy": "Luật lọc thủ công (Excel)",
            "targeted_customers": int(tp_m + fp_m),
            "churn_caught": int(tp_m),
            "churn_missed": int(fn_m),
            "total_cost_usd": round(float(cost_manual), 2),
            "savings_vs_nothing_usd": round(float(cost_do_nothing - cost_manual), 2),
            "savings_pct": round(float((cost_do_nothing - cost_manual) / cost_do_nothing * 100), 1),
        },
        {
            "strategy": f"{best_model_name} (T*={optimal_threshold:.2f})",
            "targeted_customers": int(opt_m["tp"] + opt_m["fp"]),
            "churn_caught": int(opt_m["tp"]),
            "churn_missed": int(opt_m["fn"]),
            "total_cost_usd": round(float(cost_model_opt), 2),
            "savings_vs_nothing_usd": round(float(cost_do_nothing - cost_model_opt), 2),
            "savings_pct": round(float((cost_do_nothing - cost_model_opt) / cost_do_nothing * 100), 1),
        },
    ]

    # 5. Trich xuat He so Odds Ratio tu Logistic Regression
    lr_pipe = trained_models["Logistic Regression"]
    lr_feat_names = get_clean_feature_names(lr_pipe.named_steps["prep"])
    lr_coefs = lr_pipe.named_steps["clf"].coef_[0]
    lr_df = pd.DataFrame({
        "feature": lr_feat_names,
        "coef": np.round(lr_coefs, 4),
        "odds_ratio": np.round(np.exp(lr_coefs), 4),
    }).sort_values("coef", ascending=False)
    lr_top_pos = lr_df.head(6).to_dict(orient="records")
    lr_top_neg = lr_df.tail(6).iloc[::-1].to_dict(orient="records")

    # 6. Trich xuat luat tuong minh tu Cay quyet dinh CART cat tia do sau 3 (khong scale bien so de doc nguong that!)
    raw_tree_cols_num = ["tenure", "MonthlyCharges", "TotalAddonServices"]
    raw_tree_cols_cat = ["Contract", "InternetService", "OnlineSecurity", "TechSupport", "PaymentMethod"]
    prep_interpret = ColumnTransformer([
        ("num", "passthrough", raw_tree_cols_num),
        ("cat", OneHotEncoder(handle_unknown="ignore", sparse_output=False), raw_tree_cols_cat),
    ])
    X_tr_int = prep_interpret.fit_transform(X_train)
    int_feat_names = [n.replace("num__", "").replace("cat__", "") for n in prep_interpret.get_feature_names_out()]
    cart_interpret = DecisionTreeClassifier(max_depth=3, min_samples_leaf=50, random_state=42)
    cart_interpret.fit(X_tr_int, y_train)
    extracted_rules = extract_cart_rules(cart_interpret, int_feat_names, len(X_train))

    # 7. Giai thich SHAP Global & Local tren best_tree_name
    tree_pipe = trained_models[best_tree_name]
    prep_tree = tree_pipe.named_steps["prep"]
    clf_tree = tree_pipe.named_steps["clf"]
    tree_feat_names = get_clean_feature_names(prep_tree)
    X_test_trans = prep_tree.transform(X_test)
    X_test_df_trans = pd.DataFrame(X_test_trans, columns=tree_feat_names)

    explainer = shap.TreeExplainer(clf_tree)
    shap_raw = explainer.shap_values(X_test_df_trans)
    if isinstance(shap_raw, list):
        shap_vals = np.array(shap_raw[1])
        base_val = float(explainer.expected_value[1])
    elif shap_raw.ndim == 3:
        shap_vals = shap_raw[:, :, 1]
        base_val = float(explainer.expected_value[1])
    else:
        shap_vals = np.array(shap_raw)
        ev = explainer.expected_value
        base_val = float(ev[1] if isinstance(ev, (list, np.ndarray)) and len(np.atleast_1d(ev)) > 1 else np.atleast_1d(ev)[0])

    mean_abs_shap = np.abs(shap_vals).mean(axis=0)
    shap_rank_df = pd.DataFrame({
        "feature": tree_feat_names,
        "mean_abs_shap": np.round(mean_abs_shap, 4),
    }).sort_values("mean_abs_shap", ascending=False)
    top12_shap = shap_rank_df.head(12).to_dict(orient="records")

    # Chon 2 ca dien hinh tren tap Test: 1 ca Churn rui ro cao (>80%), 1 ca An toan (<5%)
    tree_probs_test = test_probs[best_tree_name]
    y_test_arr = y_test.values
    churn_high_idxs = np.where((y_test_arr == 1) & (tree_probs_test >= 0.75))[0]
    safe_low_idxs = np.where((y_test_arr == 0) & (tree_probs_test <= 0.08))[0]
    idx_high = int(churn_high_idxs[np.argmax(tree_probs_test[churn_high_idxs])])
    idx_low = int(safe_low_idxs[np.argmin(tree_probs_test[safe_low_idxs])])

    def describe_case(idx_pos):
        orig_row = X_test.iloc[idx_pos]
        sv = shap_vals[idx_pos]
        top_idx = np.argsort(np.abs(sv))[::-1][:5]
        return {
            "test_pos": idx_pos,
            "customer_id": str(raw_df.loc[X_test.index[idx_pos], "customerID"]),
            "actual_churn": int(y_test_arr[idx_pos]),
            "pred_prob_pct": round(float(tree_probs_test[idx_pos]) * 100, 2),
            "contract": str(orig_row["Contract"]),
            "tenure": int(orig_row["tenure"]),
            "monthly_charges": round(float(orig_row["MonthlyCharges"]), 2),
            "internet": str(orig_row["InternetService"]),
            "payment": str(orig_row["PaymentMethod"]),
            "top_shap_drivers": [
                {"feature": tree_feat_names[j], "shap_value": round(float(sv[j]), 4)}
                for j in top_idx
            ],
        }

    shap_cases = {
        "high_risk_churn": describe_case(idx_high),
        "safe_non_churn": describe_case(idx_low),
    }

    # 8. Phan tich loi chuyen sau (Error Analysis) tren tap Test cho best_model tai nguong 0.50 & T*
    def analyze_errors(probs_arr, thr):
        preds = (probs_arr >= thr).astype(int)
        err_mask = preds != y_test_arr
        fn_mask = (y_test_arr == 1) & (preds == 0)
        fp_mask = (y_test_arr == 0) & (preds == 1)

        # Ca vung bien sat nguong quyet dinh (|p - thr| <= 0.10)
        boundary_err = err_mask & (np.abs(probs_arr - thr) <= 0.10)
        # FN: khach dai han (tenure >= 24 hoac hop dong 1-2 nam) roi bo dot ngot
        fn_loyal = fn_mask & (
            (X_test["tenure"].values >= 24) | (X_test["Contract"].values != "Month-to-month")
        )
        # FN: khach ngan han / hop dong thang nhung mo hinh du doan thap
        fn_short = fn_mask & (~fn_loyal)
        # FP: khach dung hop dong theo thang hoac Fiber optic nhung thuc te van o lai
        fp_high_profile = fp_mask & (
            (X_test["Contract"].values == "Month-to-month") | (X_test["InternetService"].values == "Fiber optic")
        )
        total_err = int(err_mask.sum())
        return {
            "threshold": thr,
            "total_test": int(len(y_test_arr)),
            "total_errors": total_err,
            "error_rate_pct": round(total_err / len(y_test_arr) * 100, 2),
            "fn_count": int(fn_mask.sum()),
            "fp_count": int(fp_mask.sum()),
            "boundary_errors": int(boundary_err.sum()),
            "boundary_share_pct": round(int(boundary_err.sum()) / max(total_err, 1) * 100, 1),
            "fn_loyal_silent": int(fn_loyal.sum()),
            "fn_short_volatile": int(fn_short.sum()),
            "fp_high_risk_profile": int(fp_high_profile.sum()),
            "fp_other": int(fp_mask.sum() - fp_high_profile.sum()),
        }

    error_analysis = {
        "at_default_050": analyze_errors(test_probs[best_model_name], 0.50),
        "at_optimal_tstar": analyze_errors(test_probs[best_model_name], optimal_threshold),
    }

    # ==================== XUAT 7 BIEU DO THUC NGHIEM 300 DPI ====================

    # --- FIG 04: SO SANH HIEU NANG 5 MO HINH ---
    fig, ax = plt.subplots(figsize=(11.8, 5.2), dpi=300)
    fig.patch.set_facecolor("white")
    m_names = list(model_specs.keys())
    metrics_plot = ["accuracy", "recall", "f1", "roc_auc", "pr_auc"]
    metric_labels = ["Accuracy", "Recall (T=0.5)", "F1-Score", "ROC-AUC", "PR-AUC"]
    colors_bar = [SLATE, AMBER, TEAL, BLUE, CRIMSON]

    x_pos = np.arange(len(m_names))
    bar_w = 0.155
    for i, (mk, ml, col) in enumerate(zip(metrics_plot, metric_labels, colors_bar)):
        vals = [test_metrics_05[m][mk] for m in m_names]
        offset = (i - 2) * bar_w
        bars = ax.bar(x_pos + offset, vals, width=bar_w, label=ml, color=col)
        for b, v in zip(bars, vals):
            ax.text(b.get_x() + b.get_width()/2, v + 0.012, f"{v:.2f}", ha="center", va="bottom",
                    fontsize=7.3, fontweight="bold", color="#0F172A")

    ax.set_xticks(x_pos)
    ax.set_xticklabels(m_names, fontsize=9.5, fontweight="bold", color=NAVY)
    ax.set_ylim(0, 1.02)
    ax.set_ylabel("Giá trị chỉ số đánh giá", fontsize=10, fontweight="bold", color=SLATE)
    ax.set_title("So sánh hiệu năng 5 mô hình phân loại trên tập Test độc lập (N = 1.409)",
                 fontsize=12, fontweight="bold", color=NAVY, pad=12)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.12), ncol=5, frameon=True, fontsize=9)
    ax.yaxis.grid(True)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "fig04_model_comparison.png"), dpi=300, bbox_inches="tight")
    plt.close(fig)

    # --- FIG 05: DUONG CONG ROC VA PRECISION-RECALL ---
    fig, axes = plt.subplots(1, 2, figsize=(12.2, 5.0), dpi=300)
    fig.patch.set_facecolor("white")
    line_colors = {
        "Logistic Regression": BLUE_MID,
        "Decision Tree (CART)": SLATE,
        "Random Forest": TEAL,
        "XGBoost": AMBER,
        "LightGBM": CRIMSON,
    }
    line_styles = {
        "Logistic Regression": "-",
        "Decision Tree (CART)": "--",
        "Random Forest": "-.",
        "XGBoost": "-",
        "LightGBM": ":",
    }

    ax = axes[0]
    for name in m_names:
        fpr, tpr, _ = roc_curve(y_test, test_probs[name])
        auc_v = test_metrics_05[name]["roc_auc"]
        ax.plot(fpr, tpr, lw=2.0, color=line_colors[name], linestyle=line_styles[name],
                label=f"{name} (AUC = {auc_v:.4f})")
    ax.plot([0, 1], [0, 1], "k--", lw=1.0, alpha=0.5, label="Ngẫu nhiên (AUC = 0.50)")
    ax.set_xlabel("False Positive Rate (FPR)", fontsize=9.5, fontweight="bold")
    ax.set_ylabel("True Positive Rate (Recall)", fontsize=9.5, fontweight="bold")
    ax.set_title("A. Đường cong ROC (ROC-AUC)", fontsize=11.5, fontweight="bold", color=NAVY)
    ax.legend(loc="lower right", fontsize=8.3, frameon=True, facecolor="white")
    ax.grid(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    ax = axes[1]
    base_pr = float(y_test.mean())
    for name in m_names:
        prec, rec, _ = precision_recall_curve(y_test, test_probs[name])
        ap_v = test_metrics_05[name]["pr_auc"]
        ax.plot(rec, prec, lw=2.0, color=line_colors[name], linestyle=line_styles[name],
                label=f"{name} (PR-AUC = {ap_v:.4f})")
    ax.axhline(base_pr, color="black", linestyle="--", lw=1.0, alpha=0.5,
               label=f"Baseline tỷ lệ lớp dương ({base_pr:.2f})")
    ax.set_xlabel("Recall (Độ nhạy)", fontsize=9.5, fontweight="bold")
    ax.set_ylabel("Precision (Độ chính xác dương)", fontsize=9.5, fontweight="bold")
    ax.set_title("B. Đường cong Precision-Recall (PR-AUC)", fontsize=11.5, fontweight="bold", color=NAVY)
    ax.legend(loc="upper right", fontsize=8.3, frameon=True, facecolor="white")
    ax.grid(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    plt.tight_layout(pad=2.0)
    plt.savefig(os.path.join(FIG_DIR, "fig05_roc_pr_curves.png"), dpi=300, bbox_inches="tight")
    plt.close(fig)

    # --- FIG 06: MA TRAN NHAM LAN (4 MO HINH TIEU BIEU + TOI UU NGUONG T*) ---
    fig, axes = plt.subplots(1, 4, figsize=(13.6, 3.6), dpi=300)
    fig.patch.set_facecolor("white")
    cm_panels = [
        ("Decision Tree (T=0.50)", test_probs["Decision Tree (CART)"], 0.50),
        ("Random Forest (T=0.50)", test_probs["Random Forest"], 0.50),
        (f"{best_model_name} (T=0.50)", test_probs[best_model_name], 0.50),
        (f"{best_model_name} (T*={optimal_threshold:.2f})", test_probs[best_model_name], optimal_threshold),
    ]
    for ax, (title_cm, probs_cm, thr_cm) in zip(axes, cm_panels):
        cm = confusion_matrix(y_test, (probs_cm >= thr_cm).astype(int))
        sns.heatmap(
            cm, annot=True, fmt="d", cmap="Blues", cbar=False, square=True,
            xticklabels=["Ở lại (0)", "Rời bỏ (1)"],
            yticklabels=["Ở lại (0)", "Rời bỏ (1)"],
            annot_kws={"size": 10.5, "fontweight": "bold"},
            ax=ax,
        )
        rec_v = cm[1, 1] / (cm[1, 0] + cm[1, 1]) * 100
        ax.set_title(f"{title_cm}\nRecall = {rec_v:.1f}%", fontsize=9.3, fontweight="bold", color=NAVY)
        ax.set_xlabel("Dự đoán", fontsize=8.8)
        ax.set_ylabel("Thực tế", fontsize=8.8)
    plt.tight_layout(pad=1.4)
    plt.savefig(os.path.join(FIG_DIR, "fig06_confusion_matrices.png"), dpi=300, bbox_inches="tight")
    plt.close(fig)

    # --- FIG 07: TOI UU NGUONG CHI PHI T* VA DOI CHUNG ROI 4 CHIEN LUOC ---
    fig, axes = plt.subplots(1, 2, figsize=(12.6, 4.8), dpi=300, gridspec_kw={"width_ratios": [1.05, 1.0]})
    fig.patch.set_facecolor("white")

    # Panel A: Duong cong chi phi theo nguong T (OOF Train & Test)
    ax = axes[0]
    test_costs_curve = [
        compute_metrics(y_test, test_probs[best_model_name], threshold=t)["avg_cost_per_customer_usd"]
        for t in thresholds
    ]
    ax.plot(thresholds, oof_costs, color=BLUE, lw=2.2, label="Chi phí trung bình trên OOF Train (5-Fold)")
    ax.plot(thresholds, test_costs_curve, color=CRIMSON, lw=2.0, linestyle="--", label="Chi phí trung bình trên tập Test")
    ax.axvline(optimal_threshold, color=TEAL, linestyle=":", lw=2.0,
               label=f"Ngưỡng tối ưu T* = {optimal_threshold:.2f}")
    ax.axvline(0.50, color=SLATE, linestyle="-.", lw=1.3, label="Ngưỡng mặc định T = 0.50")
    ax.set_xlabel("Ngưỡng phân loại xác suất (Threshold T)", fontsize=9.5, fontweight="bold")
    ax.set_ylabel("Chi phí kỳ vọng trung bình (USD / KH)", fontsize=9.5, fontweight="bold")
    ax.set_title("A. Tối ưu ngưỡng quyết định nhạy chi phí (Cost Curve)", fontsize=11, fontweight="bold", color=NAVY)
    ax.legend(loc="upper left", fontsize=8.3, frameon=True, facecolor="white")
    ax.grid(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    # Panel B: Doi chung tong chi phi 4 kich ban tren tap Test (N=1409)
    ax = axes[1]
    s_labels = [
        "Không\ncan thiệp",
        "Ưu đãi\nđại trà 100%",
        "Luật lọc\nthủ công",
        f"{best_model_name}\n(T*={optimal_threshold:.2f})",
    ]
    s_costs = [r["total_cost_usd"] / 1000.0 for r in roi_benchmark]
    s_cols = [CRIMSON, AMBER, SLATE, TEAL]
    bars = ax.bar(s_labels, s_costs, color=s_cols, width=0.52)
    for b, c_k, r in zip(bars, s_costs, roi_benchmark):
        sub = f"\n(-{r['savings_pct']:.1f}%)" if r["savings_pct"] > 0 else ""
        ax.text(b.get_x() + b.get_width()/2, c_k + 3.5, f"${c_k:.1f}k{sub}",
                ha="center", va="bottom", fontsize=8.8, fontweight="bold", color="#0F172A")
    ax.set_ylim(0, max(s_costs) * 1.22)
    ax.set_ylabel("Tổng chi phí nghiệp vụ (nghìn USD)", fontsize=9.5, fontweight="bold")
    ax.set_title("B. Đối chứng tổng chi phí trên tập Test (N = 1.409)", fontsize=11, fontweight="bold", color=NAVY)
    ax.yaxis.grid(True)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    plt.tight_layout(pad=2.0)
    plt.savefig(os.path.join(FIG_DIR, "fig07_cost_threshold_roi.png"), dpi=300, bbox_inches="tight")
    plt.close(fig)

    # --- FIG 08: CAY QUYET DINH CAT TIA DO SAU 3 (TRICH XUAT LUAT) ---
    fig, ax = plt.subplots(figsize=(14.0, 6.2), dpi=300)
    fig.patch.set_facecolor("white")
    plot_tree(
        cart_interpret,
        feature_names=int_feat_names,
        class_names=["Ở lại", "Rời bỏ"],
        filled=True,
        rounded=True,
        proportion=False,
        precision=2,
        fontsize=8,
        ax=ax,
    )
    ax.set_title("Cây quyết định CART cắt tỉa độ sâu 3 phục vụ rút trích luật IF-THEN tường minh",
                 fontsize=12.5, fontweight="bold", color=NAVY, pad=12)
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "fig08_decision_tree_rules.png"), dpi=300, bbox_inches="tight")
    plt.close(fig)

    # --- FIG 09: SHAP SUMMARY BEESWARM PLOT ---
    plt.figure(figsize=(10.5, 6.0), dpi=300)
    shap.summary_plot(
        shap_vals,
        X_test_df_trans,
        max_display=12,
        show=False,
        plot_size=(10.5, 6.0),
    )
    plt.title(f"Giải thích toàn cục SHAP Summary Plot — Top 12 đặc trưng ({best_tree_name})",
              fontsize=11.5, fontweight="bold", color=NAVY, pad=12)
    plt.xlabel("Giá trị SHAP (tác động lên log-odds rời bỏ dịch vụ)", fontsize=9.5, fontweight="bold")
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, "fig09_shap_summary.png"), dpi=300, bbox_inches="tight")
    plt.close()

    # --- FIG 10: SHAP LOCAL WATERFALL / BAR SO SANH 2 CA DIEN HINH ---
    fig, axes = plt.subplots(1, 2, figsize=(13.0, 5.2), dpi=300)
    fig.patch.set_facecolor("white")
    for ax, idx_case, title_case in [
        (axes[0], idx_high, f"A. Ca rủi ro cao ({shap_cases['high_risk_churn']['customer_id']})\n"
                            f"Xác suất Churn = {shap_cases['high_risk_churn']['pred_prob_pct']:.1f}% | Thực tế: Rời bỏ"),
        (axes[1], idx_low, f"B. Ca trung thành ({shap_cases['safe_non_churn']['customer_id']})\n"
                           f"Xác suất Churn = {shap_cases['safe_non_churn']['pred_prob_pct']:.1f}% | Thực tế: Ở lại"),
    ]:
        sv = shap_vals[idx_case]
        top_k = np.argsort(np.abs(sv))[::-1][:8][::-1]
        f_lbls = [tree_feat_names[j] for j in top_k]
        v_vals = [sv[j] for j in top_k]
        b_cols = [CRIMSON if v > 0 else TEAL for v in v_vals]
        bars = ax.barh(f_lbls, v_vals, color=b_cols, height=0.60)
        ax.axvline(0, color=NAVY, lw=1.0)
        span = max(abs(min(v_vals)), abs(max(v_vals))) + 0.15
        ax.set_xlim(-span * 1.25, span * 1.25)
        for b, v in zip(bars, v_vals):
            ha = "left" if v >= 0 else "right"
            off = span * 0.04 if v >= 0 else -span * 0.04
            ax.text(v + off, b.get_y() + b.get_height()/2, f"{v:+.3f}",
                    va="center", ha=ha, fontsize=8.5, fontweight="bold", color="#0F172A")
        ax.set_title(title_case, fontsize=10.2, fontweight="bold", color=NAVY, pad=10)
        ax.set_xlabel("Đóng góp SHAP (+ tăng nguy cơ, - giảm nguy cơ)", fontsize=9)
        ax.xaxis.grid(True)
        ax.set_axisbelow(True)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
    plt.tight_layout(pad=2.0)
    plt.savefig(os.path.join(FIG_DIR, "fig10_shap_waterfall_cases.png"), dpi=300, bbox_inches="tight")
    plt.close(fig)

    # 9. Luu Pipeline tot nhat va file JSON ket qua thuc nghiem toan dien
    artifact_bundle = {
        "best_model_name": best_model_name,
        "best_model_pipeline": trained_models[best_model_name],
        "best_tree_name": best_tree_name,
        "best_tree_pipeline": trained_models[best_tree_name],
        "optimal_threshold": optimal_threshold,
        "cost_params": {
            "cost_fp": COST_FP,
            "cost_tp": COST_TP,
            "cost_fn": COST_FN,
            "cost_tn": COST_TN,
        },
        "tree_feature_names": tree_feat_names,
    }
    joblib.dump(artifact_bundle, os.path.join(MODEL_DIR, "best_churn_pipeline.joblib"))

    results = {
        "dataset_audit": audit,
        "split": {
            "train_size": int(len(X_train)),
            "test_size": int(len(X_test)),
            "train_churn_rate": round(float(y_train.mean()), 4),
            "test_churn_rate": round(float(y_test.mean()), 4),
        },
        "cv_roc_auc": cv_summary,
        "best_params": best_params_all,
        "train_times_sec": train_times,
        "test_metrics_at_050": test_metrics_05,
        "test_metrics_at_tstar": test_metrics_tstar,
        "best_model_name": best_model_name,
        "best_tree_name": best_tree_name,
        "optimal_threshold": optimal_threshold,
        "bayes_theoretical_threshold": round(COST_FP / (COST_FN - COST_TP + COST_FP), 2),
        "cost_matrix": {
            "C_FP": COST_FP,
            "C_TP": COST_TP,
            "C_FN": COST_FN,
            "C_TN": COST_TN,
        },
        "ablation_study": ablation_study,
        "roi_benchmark": roi_benchmark,
        "logistic_odds_ratios": {
            "top_increase_churn": lr_top_pos,
            "top_decrease_churn": lr_top_neg,
        },
        "extracted_cart_rules": extracted_rules,
        "shap_top12": top12_shap,
        "shap_cases": shap_cases,
        "error_analysis": error_analysis,
        "total_runtime_sec": round(time.time() - t0_all, 2),
    }

    with open(os.path.join(MODEL_DIR, "experiment_results.json"), "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    # Luu bang CSV so sanh 5 mo hinh cho Streamlit & kiem tra nhanh
    rows_csv = []
    for name in m_names:
        m = test_metrics_05[name]
        rows_csv.append({
            "Mô hình": name,
            "CV ROC-AUC": cv_summary[name],
            "Accuracy": m["accuracy"],
            "Precision": m["precision"],
            "Recall (T=0.5)": m["recall"],
            "F1-Score": m["f1"],
            "ROC-AUC": m["roc_auc"],
            "PR-AUC": m["pr_auc"],
            "Brier Score": m["brier"],
            f"Recall (T*={optimal_threshold})": test_metrics_tstar[name]["recall"],
            f"Chi phí T* (USD)": test_metrics_tstar[name]["total_cost_usd"],
            "Thời gian (s)": train_times[name],
        })
    pd.DataFrame(rows_csv).to_csv(os.path.join(MODEL_DIR, "model_comparison.csv"), index=False, encoding="utf-8-sig")

    print("\n=== ROI BENCHMARK ON TEST SET (N=1,409) ===")
    for r in roi_benchmark:
        print(f"  {r['strategy']:34s} | Cost=${r['total_cost_usd']:9,.0f} | Savings=${r['savings_vs_nothing_usd']:8,.0f} ({r['savings_pct']}%)")

    print("\n=== TOP 5 SHAP FEATURES ===")
    for item in top12_shap[:5]:
        print(f"  {item['feature']:35s}: mean|SHAP| = {item['mean_abs_shap']:.4f}")

    print(f"\nAll training, evaluation, SHAP, and 7 charts completed in {results['total_runtime_sec']}s.")


if __name__ == "__main__":
    run_training()
