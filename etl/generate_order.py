import pandas as pd
from openpyxl import load_workbook
import os
from dotenv import load_dotenv
from datetime import datetime, timedelta
import requests, json
import numpy as np

load_dotenv()  # reads COVA_USERNAME, COVA_PASSWORD, COVA_CLIENT from .env

def generate_order(output_path: str, hist_days: int = 30, exclude_today: bool = False):
    """
    Authenticates to Cova, pulls historical IOH, computes metrics,
    fetches 7-day & custom-range sales, merges everything, and writes
    one Excel sheet per location.
    """

    # ── Step 1: Authenticate ────────────────────────────────────────────────
    USER = os.getenv("COVA_USERNAME")
    PWD  = os.getenv("COVA_PASSWORD")
    KEY  = os.getenv("COVA_CLIENT")
    auth = requests.post(
        "https://signinbackend.iqmetrix.net/v1/oauth2/token",
        json={
            "UsernameOrEmailAddress": USER,
            "Password":               PWD,
            "ClientKey":              KEY
        },
        headers={"Content-Type": "application/json"}
    )
    auth.raise_for_status()
    token = auth.json()["token"]
    bearer = {
        "Authorization": f"Bearer {token}",
        "Content-Type":  "application/json"
    }
    now = datetime.now()

    # ── Step 2: Historical IOH ─────────────────────────────────────────────
    def fetch_ioh_for_date(dt: datetime) -> pd.DataFrame:
        resp = requests.post(
            "https://covareportservice-prod-westus.azurewebsites.net/"
            "v2/Companies/131096/Reports/"
            "1c3c6f4a-d91b-40fa-880f-3852b68de20e/Execute",
            json={
                "ReportId":  "1c3c6f4a-d91b-40fa-880f-3852b68de20e",
                "TimeZone":  "America/Edmonton",
                "Parameters": json.dumps({
                    "CompanyId":       131096,
                    "Date":            dt.strftime("%Y-%m-%d"),
                    "Entities":        [230791, 167209, 237603],
                    "Classifications": [3331],
                    "InStockOnly":     False
                })
            },
            headers=bearer
        )
        resp.raise_for_status()
        body = resp.json()
        if not body:
            return pd.DataFrame()
        df = pd.DataFrame(body[0].get("Data", []))
        df["Date"] = pd.to_datetime(dt)
        return df

    # build combined historical IOH
    last_day = now - timedelta(days=1) if exclude_today else now
    ioh_frames = []
    for i in range(hist_days):
        day = last_day - timedelta(days=i)
        df_day = fetch_ioh_for_date(day)
        if not df_day.empty:
            ioh_frames.append(df_day)
    comb_df = pd.concat(ioh_frames, ignore_index=True) if ioh_frames else pd.DataFrame()
    if comb_df.empty:
        raise RuntimeError("No historical IOH data fetched")

    # ── Step 3: Compute IOH metrics ─────────────────────────────────────────
    comb_df.sort_values(["SKU","Location","Date"], inplace=True)
    comb_df["Was In Stock"] = comb_df["In Stock Qty"] > 0
    comb_df["Stock Change"]  = (
        comb_df.groupby(["SKU","Location"])["Was In Stock"]
               .diff().fillna(0).astype(int)
    )

    last_in = (
        comb_df[comb_df["Was In Stock"]]
        .groupby(["SKU","Location"])["Date"]
        .max().reset_index(name="Last In Stock Date")
    )

    def avg_cycle_days(g):
        s = g[g["Stock Change"] == 1]["Date"].reset_index(drop=True)
        e = g[g["Stock Change"] == -1]["Date"].reset_index(drop=True)
        n = min(len(s), len(e))
        if n == 0:
            return 0.0
        durs = (e[:n].values - s[:n].values) \
               .astype("timedelta64[D]").astype(int)
        durs = durs[durs > 0]
        return float(durs.mean()) if len(durs) else 0.0

    avg   = comb_df.groupby(["SKU","Location"]).apply(avg_cycle_days) \
                   .reset_index(name="Avg Days In Stock Per Cycle")
    var   = comb_df.groupby(["SKU","Location"])["In Stock Qty"] \
                   .std().reset_index(name="Stock Variability")
    freq  = comb_df[comb_df["Stock Change"]==-1] \
                   .groupby(["SKU","Location"]).size() \
                   .reset_index(name="Stockout Frequency")

    comb_df["Days in Stock Index"] = comb_df["Was In Stock"].astype(int)
    totals = (
        comb_df.groupby(["SKU","Location"], as_index=False)
               .agg({
                   "Days in Stock Index":"sum",
                   "In Stock Qty":       "sum"
               })
               .rename(columns={
                   "Days in Stock Index":"Total Days in Stock",
                   "In Stock Qty":       "Total In Stock Qty"
               })
    )
    grouped = totals.merge(last_in, on=["SKU","Location"], how="left")
    for dfm in (avg, var, freq):
        grouped = grouped.merge(dfm, on=["SKU","Location"], how="left")

    # ── Step 4: Current IOH ─────────────────────────────────────────────────
    inv = requests.post(
        "https://covareportservice-prod-westus.azurewebsites.net/"
        "v2/Companies/131096/Reports/"
        "a8b03840-2e18-4c11-bdb3-6413b972d391/Execute",
        json={
            "ReportId":  "a8b03840-2e18-4c11-bdb3-6413b972d391",
            "TimeZone":  "America/Edmonton",
            "Parameters": json.dumps({
                "CompanyId":       131096,
                "Entities":        [230791,167209,237603],
                "Classifications": [3331],
                "InStockOnly":     False,
                "IncludeLocation": True
            })
        },
        headers=bearer
    )
    inv.raise_for_status()
    ioh_df = pd.DataFrame(inv.json()[0].get("Data", [])) if inv.json() else pd.DataFrame()
    # fill missing first/last received dates to next Thursday
    to_thu = (3 - now.weekday()) % 7
    fill_date = (now + timedelta(days=to_thu)).date()
    for c in ["First Received Date","Last Received Date"]:
        if c in ioh_df:
            ioh_df[c] = pd.to_datetime(ioh_df[c]).dt.date.fillna(fill_date)

    # ── Step 5: Sales‐fetch helper ───────────────────────────────────────────
    def fetch_sales(report_id: str,
                    rename_map: dict,
                    dr_type: int,
                    start_date: datetime = None,
                    end_date:   datetime = None) -> pd.DataFrame:
        # determine window
        if dr_type == 15:
            sd = ed = datetime.now()
        else:
            sd, ed = start_date, end_date

        params = {
            "CompanyId": 131096,
            "DateRange": {
                "StartDate":     sd.strftime("%Y-%m-%dT00:00:00"),
                "EndDate":       ed.strftime("%Y-%m-%dT23:59:59"),
                "DateRangeType": dr_type
            },
            "Entities":        [167209,237603,230791],
            "Classifications": [3331],
            "SaleType":        0,
            "UseType":         0,
            "DeliveryType":    0
        }
        payload = {
            "ReportId":   report_id,
            "TimeZone":   "America/Edmonton",
            "Parameters": json.dumps(params)
        }
        r = requests.post(
            f"https://covareportservice-prod-westus.azurewebsites.net/"
            f"v2/Companies/131096/Reports/{report_id}/Execute",
            json=payload,
            headers=bearer
        )
        r.raise_for_status()
        body = r.json()
        if not body:
            return pd.DataFrame()
        df = pd.DataFrame(body[0].get("Data", []))
        return df.rename(columns=rename_map)

    # ── Step 6: Fetch 7-day & custom‐range sales ────────────────────────────
    week_map = {
        "Net Sold":          "Week Net Sold",
        "Avg Sold At Price": "Week Avg Price",
        "Total Cost":        "Week Total Cost"
    }
    sel_map = {
        "Net Sold":          f"{hist_days}d Net Sold",
        "Avg Sold At Price": f"{hist_days}d Avg Price",
        "Total Cost":        f"{hist_days}d Total Cost"
    }

    # 7-day (excl today) → dr_type=15, no start/end args
    week_df = fetch_sales(
        "c1ec9df0-db1e-4698-8d1c-dd640bdbbc04",
        week_map,
        15
    )

    # custom range → dr_type=9, must compute start_sel/end_sel first
    sel_end   = now - timedelta(days=1) if exclude_today else now
    sel_start = sel_end - timedelta(days=hist_days - 1)
    sel_df = fetch_sales(
        "c1ec9df0-db1e-4698-8d1c-dd640bdbbc04",
        sel_map,
        9,
        start_date=sel_start,
        end_date=sel_end
    )
    needed = ["Location","SKU"] + list(sel_map.values())
    if sel_df.empty:
        sel_df = pd.DataFrame(columns=needed)
    else:
        for col in needed:
            if col not in sel_df.columns:
                sel_df[col] = 0

    # ── Step 7: Merge & finalize ───────────────────────────────────────────
    merged = (
        ioh_df
        .merge(week_df[["Location","SKU"] + list(week_map.values())],
               on=["Location","SKU"], how="left")
        .merge(sel_df[needed], on=["Location","SKU"], how="left")
        .fillna({**{v:0 for v in week_map.values()},
                 **{v:0 for v in sel_map.values()}})
    )
    merged["Supplier SKU"] = merged["Supplier SKU"].apply(
        lambda s: next((x for x in str(s).split(",") if x.startswith("CNB-")), "")
    )
    final_df = merged.merge(grouped, on=["Location","SKU"], how="left")
    final_df["Sales per Day"] = (
        final_df[f"{hist_days}d Net Sold"]
        / final_df["Total Days in Stock"].replace(0, np.nan)
    )

    # ── Step 8: Write to Excel ─────────────────────────────────────────────
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    print(f"Writing data to {output_path}")
    
    # Add a debug summary sheet
    with pd.ExcelWriter(output_path, engine="openpyxl", datetime_format="yyyy-mm-dd") as writer:
        # First write individual location sheets
        locations = final_df["Location"].unique()
        print(f"Found {len(locations)} unique locations: {locations}")
        
        for loc in locations:
            sheet = str(loc)[:31].replace(":","-").replace("/","-")
            print(f"Writing sheet for location: {loc} (sheet name: {sheet})")
            
            loc_df = final_df[final_df["Location"]==loc]
            print(f"  - {len(loc_df)} rows, columns: {loc_df.columns.tolist()}")
            
            loc_df.to_excel(writer, sheet_name=sheet, index=False)
        
        # Also write a combined sheet with all data
        print("Writing combined 'All_Locations' sheet")
        final_df.to_excel(writer, sheet_name="All_Locations", index=False)
        
        # Write a summary sheet with metadata
        summary_data = {
            "Metric": [
                "Generation Date/Time",
                "Number of Locations",
                "Total Products",
                "Historical Days Used",
                "Excluded Today",
                "Columns Available"
            ],
            "Value": [
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                len(locations),
                len(final_df["SKU"].unique()),
                hist_days,
                "Yes" if exclude_today else "No",
                ", ".join(final_df.columns)
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        
    print(f"Excel file written successfully to {output_path}")
