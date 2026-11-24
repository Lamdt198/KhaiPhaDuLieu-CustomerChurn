# -*- coding: utf-8 -*-
"""
src/data_loader.py
Tai du lieu IBM Telco Customer Churn (luu cuc bo de chay 100% offline),
lam sach du lieu va tao 6 dac trung mien (Domain Feature Engineering).
"""

import os
import urllib.request
import numpy as np
import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
DATA_PATH = os.path.join(DATA_DIR, "WA_Fn-UseC_-Telco-Customer-Churn.csv")
IBM_URL = (
    "https://raw.githubusercontent.com/IBM/telco-customer-churn-on-icp4d/"
    "master/data/Telco-Customer-Churn.csv"
)

ADDON_SERVICE_COLS = [
    "OnlineSecurity",
    "OnlineBackup",
    "DeviceProtection",
    "TechSupport",
    "StreamingTV",
    "StreamingMovies",
]

NUMERIC_FEATURES = [
    "tenure",
    "MonthlyCharges",
    "TotalCharges",
    "AvgMonthlySpend",
    "ChargeDiff",
    "TotalAddonServices",
]

BINARY_NUMERIC_FEATURES = [
    "SeniorCitizen",
    "HasProtectionBundle",
    "IsAutoPayment",
]

CATEGORICAL_FEATURES = [
    "gender",
    "Partner",
    "Dependents",
    "PhoneService",
    "MultipleLines",
    "InternetService",
    "OnlineSecurity",
    "OnlineBackup",
    "DeviceProtection",
    "TechSupport",
    "StreamingTV",
    "StreamingMovies",
    "Contract",
    "PaperlessBilling",
    "PaymentMethod",
    "TenureGroup",
]


def ensure_dataset(data_path: str = DATA_PATH) -> str:
    """Dam bao file CSV goc ton tai cuc bo."""
    os.makedirs(os.path.dirname(data_path), exist_ok=True)
    if not os.path.exists(data_path):
        print(f"Downloading IBM Telco Customer Churn dataset to {data_path}...")
        urllib.request.urlretrieve(IBM_URL, data_path)
    return data_path


def load_raw_data(data_path: str = DATA_PATH) -> pd.DataFrame:
    """Doc du lieu tho nguyen ban tu file CSV."""
    path = ensure_dataset(data_path)
    return pd.read_csv(path)


def add_domain_features(df: pd.DataFrame) -> pd.DataFrame:
    """Lam sach co ban va bo sung 6 dac trung mien."""
    out = df.copy()

    # Loai bo ma dinh danh khong mang gia tri du bao
    if "customerID" in out.columns:
        out = out.drop(columns=["customerID"])

    # Chuyen TotalCharges sang so thuc; 11 dong khuyet deu co tenure == 0 -> dien 0.0
    out["TotalCharges"] = pd.to_numeric(out["TotalCharges"], errors="coerce")
    out["TotalCharges"] = out["TotalCharges"].fillna(0.0)

    out["tenure"] = pd.to_numeric(out["tenure"], errors="coerce").fillna(0).astype(int)
    out["MonthlyCharges"] = pd.to_numeric(out["MonthlyCharges"], errors="coerce").fillna(0.0)
    out["SeniorCitizen"] = pd.to_numeric(out["SeniorCitizen"], errors="coerce").fillna(0).astype(int)

    # 1. Chi tieu trung binh thuc te moi thang
    out["AvgMonthlySpend"] = np.where(
        out["tenure"] > 0,
        out["TotalCharges"] / out["tenure"],
        out["MonthlyCharges"],
    )

    # 2. Chenh lech cuoc hien tai so voi muc binh quan tich luy (phat hien bien dong tang cuoc)
    out["ChargeDiff"] = out["MonthlyCharges"] - out["AvgMonthlySpend"]

    # 3. Tong so dich vu gia tri gia tang dang su dung (0-6)
    out["TotalAddonServices"] = (out[ADDON_SERVICE_COLS] == "Yes").sum(axis=1).astype(int)

    # 4. Co dang ky goi bao mat hoac ho tro ky thuat
    out["HasProtectionBundle"] = (
        (out["OnlineSecurity"] == "Yes") | (out["TechSupport"] == "Yes")
    ).astype(int)

    # 5. Co su dung phuong thuc thanh toan tu dong
    out["IsAutoPayment"] = (
        out["PaymentMethod"].astype(str).str.contains("automatic", case=False, na=False)
    ).astype(int)

    # 6. Phan nhom vong doi khach hang theo so thang gan bo
    bins = [-1, 6, 12, 24, 48, 120]
    labels = ["0-6m", "7-12m", "13-24m", "25-48m", ">48m"]
    out["TenureGroup"] = pd.cut(out["tenure"], bins=bins, labels=labels).astype(str)

    # Chuan hoa bien muc tieu neu co
    if "Churn" in out.columns and not pd.api.types.is_numeric_dtype(out["Churn"]):
        out["Churn"] = (out["Churn"].astype(str).str.strip() == "Yes").astype(int)

    return out


def load_prepared_data(data_path: str = DATA_PATH):
    """Tra ve (df_raw, df_clean, X, y, audit_info)."""
    raw = load_raw_data(data_path)
    missing_tc = int(pd.to_numeric(raw["TotalCharges"], errors="coerce").isna().sum())
    dup_rows = int(raw.duplicated().sum())

    clean = add_domain_features(raw)
    feature_cols = NUMERIC_FEATURES + BINARY_NUMERIC_FEATURES + CATEGORICAL_FEATURES
    X = clean[feature_cols].copy()
    y = clean["Churn"].astype(int).copy()

    audit_info = {
        "n_rows": int(len(raw)),
        "n_raw_cols": int(raw.shape[1]),
        "n_features_engineered": int(len(feature_cols)),
        "missing_total_charges": missing_tc,
        "duplicate_rows": dup_rows,
        "churn_count": int(y.sum()),
        "non_churn_count": int((y == 0).sum()),
        "churn_rate": round(float(y.mean()), 4),
    }
    return raw, clean, X, y, audit_info


if __name__ == "__main__":
    raw, clean, X, y, info = load_prepared_data()
    print("Dataset loaded successfully:", info)
    print("X shape:", X.shape, "| y positive rate:", f"{y.mean()*100:.2f}%")
