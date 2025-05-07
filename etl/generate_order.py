import pandas as pd
from openpyxl import load_workbook
import os
from datetime import datetime, timedelta
import requests
import json
import numpy as np


# …your other imports (requests, numpy, etc.)

load_dotenv()  # reads COVA_USERNAME, COVA_PASSWORD, COVA_CLIENT from .env

def generate_order(output_path: str):
    # --- Step 1: Authenticate ---
    USER = os.getenv("COVA_USERNAME")
    PWD  = os.getenv("COVA_PASSWORD")
    KEY  = os.getenv("COVA_CLIENT")

    token_url = "https://signinbackend.iqmetrix.net/v1/oauth2/token"
    auth_payload = {
        "UsernameOrEmailAddress": USER,
        "Password":              PWD,
        "ClientKey":             KEY
    }
    auth_headers = {"Content-Type": "application/json"}
    resp = requests.post(token_url, json=auth_payload, headers=auth_headers)
    resp.raise_for_status()
    tok = resp.json()
    bearer_headers = {
        "Authorization": f"Bearer {tok['token']}",
        "Content-Type":  "application/json"
    }

    # --- Step 2: Helper for daily IOH fetch ---
    def fetch_ioh_for_date(target_date: datetime) -> pd.DataFrame | None:
        url = (
          "https://covareportservice-prod-westus.azurewebsites.net/"
          "v2/Companies/131096/Reports/"
          "1c3c6f4a-d91b-40fa-880f-3852b68de20e/Execute"
        )
        params = {
            "CompanyId":      131096,
            "Date":           target_date.strftime("%Y-%m-%d"),
            "Entities":       [230791, 167209, 237603],
            "Classifications":[3331],
            "InStockOnly":    False
        }
        payload = {
            "ReportId":  "1c3c6f4a-d91b-40fa-880f-3852b68de20e",
            "TimeZone":  "America/Edmonton",
            "Parameters": json.dumps(params)
        }
        r = requests.post(url, json=payload, headers=bearer_headers)
        r.raise_for_status()
        data = r.json()
        if not data:
            return None
        df = pd.DataFrame(data[0].get("Data", []))
        df["Date"] = target_date.strftime("%Y-%m-%d")
        return df

    # --- Step 3: Pull last 30 days of IOH ---
    comb_df = pd.DataFrame()
    today = datetime.now()
    for i in range(30):
        day = today - timedelta(days=i)
        df_day = fetch_ioh_for_date(day)
        if df_day is not None:
            comb_df = pd.concat([comb_df, df_day], ignore_index=True)

    # --- Step 4: Compute stock metrics ---
    comb_df["Date"]      = pd.to_datetime(comb_df["Date"])
    comb_df.sort_values(by=["SKU", "Location", "Date"], inplace=True)
    comb_df["Was In Stock"] = comb_df["In Stock Qty"] > 0

    # Last In-Stock Date
    last_in_stock = (
      comb_df[comb_df["Was In Stock"]]
      .groupby(["SKU","Location"])["Date"]
      .max()
      .reset_index()
      .rename(columns={"Date":"Last In Stock Date"})
    )

    # Stock Change flag
    comb_df["Stock Change"] = (
      comb_df.groupby(["SKU","Location"])["Was In Stock"]
      .diff().fillna(0).astype(int)
    )

    # Average days-in-stock per cycle
    def calc_cycle_days(df_slice):
        starts = df_slice[df_slice["Stock Change"]==1]["Date"].reset_index(drop=True)
        ends   = df_slice[df_slice["Stock Change"]==-1]["Date"].reset_index(drop=True)
        n = min(len(starts), len(ends))
        if n == 0:
            return 0
        durations = (ends[:n].values - starts[:n].values).astype("timedelta64[D]").astype(int)
        durations = durations[durations > 0]
        return float(np.mean(durations)) if len(durations)>0 else 0

    avg_days = (
      comb_df.groupby(["SKU","Location"], group_keys=False)
      .apply(calc_cycle_days)
      .reset_index()
      .rename(columns={0:"Avg Days In Stock Per Cycle"})
    )

    # Variability & stockout frequency
    variability = (
      comb_df.groupby(["SKU","Location"])["In Stock Qty"]
      .std()
      .reset_index()
      .rename(columns={"In Stock Qty":"Stock Variability"})
    )
    stockout_freq = (
      comb_df[comb_df["Stock Change"]==-1]
      .groupby(["SKU","Location"])
      .size()
      .reset_index(name="Stockout Frequency")
    )

    # Days-in-stock index & sums
    comb_df["Days in Stock Index"] = comb_df["Was In Stock"].astype(int)
    grouped = (
      comb_df.groupby(["SKU","Location"])
      .agg({
        "Days in Stock Index":"sum",
        "In Stock Qty":"sum"
      })
      .reset_index()
      .rename(columns={
        "Days in Stock Index":"Total Days in Stock",
        "In Stock Qty":"Total In Stock Qty"
      })
    )

    # Merge all stock metrics
    grouped = (
      grouped
      .merge(last_in_stock, on=["SKU","Location"], how="left")
      .merge(avg_days, on=["SKU","Location"], how="left")
      .merge(variability, on=["SKU","Location"], how="left")
      .merge(stockout_freq, on=["SKU","Location"], how="left")
    )

    # --- Step 5: Helper for time-range sales ---
    def fetch_sales(report_id: str, dr_type: int) -> pd.DataFrame:
        url = (
          "https://covareportservice-prod-westus.azurewebsites.net/"
          f"v2/Companies/131096/Reports/{report_id}/Execute"
        )
        params = {
          "CompanyId":      131096,
          "DateRange":      {
            "StartDate": today.strftime("%Y-%m-%d"),
            "EndDate":   today.strftime("%Y-%m-%d"),
            "DateRangeType": dr_type
          },
          "Entities":       [230791,167209,237603],
          "Classifications":[3331],
          "SaleType":       0,
          "UseType":        0,
          "DeliveryType":   0
        }
        payload = {
          "ReportId":   report_id,
          "TimeZone":   "America/Edmonton",
          "Parameters": json.dumps(params)
        }
        r = requests.post(url, json=payload, headers=bearer_headers)
        r.raise_for_status()
        js = r.json()
        if not js:
            return pd.DataFrame()
        df = pd.DataFrame(js[0].get("Data", []))
        if dr_type == 15:
            df.rename(columns={
              "Net Sold":"Week Net Sold",
              "Avg Sold At Price":"Week Avg Price",
              "Total Cost":"Week Total Cost"
            }, inplace=True)
        else:
            df.rename(columns={
              "Net Sold":"3 Week Net Sold",
              "Avg Sold At Price":"3 Week Avg Price",
              "Total Cost":"3 Week Total Cost"
            }, inplace=True)
        return df

    week_df      = fetch_sales("c1ec9df0-db1e-4698-8d1c-dd640bdbbc04", 15)
    threeweek_df = fetch_sales("c1ec9df0-db1e-4698-8d1c-dd640bdbbc04", 12)

    # --- Step 6: Merge sales & filter SKUs ---
    merged = (
      comb_df
      .merge(
        week_df[["Location","SKU","Week Net Sold","Week Avg Price","Week Total Cost"]],
        on=["Location","SKU"], how="left"
      )
      .fillna(0)
      .merge(
        threeweek_df[["Location","SKU","3 Week Net Sold","3 Week Avg Price","3 Week Total Cost"]],
        on=["Location","SKU"], how="left"
      )
      .fillna(0)
    )

    # Only CNB-prefixed SKUs in the supplier column
    merged["Supplier SKU"] = merged["Supplier SKU"].apply(
      lambda s: next((sku for sku in str(s).split(",") if sku.startswith("CNB-")), "")
    )

    final_df = merged.merge(grouped, on=["Location","SKU"], how="left")
    final_df["Sales per Day"] = final_df["3 Week Net Sold"] / final_df["Total Days in Stock"].replace(0, np.nan)

    # --- Step 7: Write out ---
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    final_df.to_excel(output_path, index=False)