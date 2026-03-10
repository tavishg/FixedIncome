#!/usr/bin/env python3
"""
Daily NAV Change Tracker for Fixed Income Funds.

Fetches daily NAV data for a list of fixed income funds, calculates
day-over-day changes, and outputs results to an Excel workbook.
"""

import argparse
import os
import sys
from datetime import datetime, timedelta

import pandas as pd
import yfinance as yf
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, numbers
from openpyxl.utils import get_column_letter

from funds_config import FUNDS

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
HISTORY_FILE = os.path.join(DATA_DIR, "nav_history.csv")


def fetch_nav(fund: dict, period: str = "10d") -> tuple[str | None, pd.DataFrame]:
    """Fetch NAV history for a fund, trying each ticker in order.

    Args:
        fund: Fund config dict with 'tickers' list.
        period: yfinance period string (e.g., "10d", "1mo", "3mo").

    Returns (working_ticker, dataframe_of_close_prices) or (None, empty_df).
    """
    for ticker_symbol in fund["tickers"]:
        try:
            ticker = yf.Ticker(ticker_symbol)
            hist = ticker.history(period=period)
            if hist.empty or "Close" not in hist.columns:
                print(f"  {ticker_symbol}: no data returned")
                continue
            closes = hist[["Close"]].dropna()
            if len(closes) < 1:
                print(f"  {ticker_symbol}: no closing prices")
                continue
            return ticker_symbol, closes
        except Exception as e:
            err_msg = str(e)
            if "curl" in err_msg.lower() or "connect" in err_msg.lower():
                print(f"  {ticker_symbol}: network error (check internet connection)")
            else:
                print(f"  {ticker_symbol}: {e}")
            continue
    return None, pd.DataFrame()


def fetch_all_navs(period: str = "10d") -> list[dict]:
    """Fetch NAV data for all configured funds."""
    results = []
    for fund in FUNDS:
        print(f"Fetching: {fund['name']}...")
        ticker_used, closes = fetch_nav(fund, period=period)

        if ticker_used is None:
            print(f"  FAILED - no working ticker found for {fund['name']}")
            print(f"  Tried: {fund['tickers']}")
            results.append({
                "name": fund["name"],
                "ticker": "N/A",
                "current_nav": None,
                "previous_nav": None,
                "change_dollar": None,
                "change_pct": None,
                "nav_date": None,
            })
            continue

        print(f"  OK - using {ticker_used} ({len(closes)} data points)")

        if len(closes) >= 2:
            current_nav = closes["Close"].iloc[-1]
            previous_nav = closes["Close"].iloc[-2]
            change_dollar = current_nav - previous_nav
            change_pct = (change_dollar / previous_nav) * 100
            nav_date = closes.index[-1].strftime("%Y-%m-%d")
        else:
            current_nav = closes["Close"].iloc[-1]
            previous_nav = None
            change_dollar = None
            change_pct = None
            nav_date = closes.index[-1].strftime("%Y-%m-%d")

        results.append({
            "name": fund["name"],
            "ticker": ticker_used,
            "current_nav": round(current_nav, 4),
            "previous_nav": round(previous_nav, 4) if previous_nav is not None else None,
            "change_dollar": round(change_dollar, 4) if change_dollar is not None else None,
            "change_pct": round(change_pct, 4) if change_pct is not None else None,
            "nav_date": nav_date,
            "closes": closes,  # Keep full series for backfill
        })

    return results


def update_history(results: list[dict], backfill: bool = False) -> pd.DataFrame:
    """Append NAV data to the history CSV and return the full history.

    If backfill=True, writes all available historical dates from the fetched data.
    Otherwise, only writes today's NAV.
    """
    if os.path.exists(HISTORY_FILE):
        history = pd.read_csv(HISTORY_FILE)
    else:
        history = pd.DataFrame(columns=["date"])

    if backfill:
        # Build a row per date from the full close series
        all_dates = set()
        for r in results:
            closes = r.get("closes")
            if closes is not None and not closes.empty:
                for dt in closes.index:
                    all_dates.add(dt.strftime("%Y-%m-%d"))

        for date_str in sorted(all_dates):
            row = {"date": date_str}
            for r in results:
                closes = r.get("closes")
                if closes is not None and not closes.empty:
                    matching = closes[closes.index.strftime("%Y-%m-%d") == date_str]
                    if not matching.empty:
                        row[r["name"]] = round(float(matching["Close"].iloc[0]), 4)
            new_row = pd.DataFrame([row])
            # Remove existing row for this date (upsert)
            history = history[history["date"] != date_str]
            history = pd.concat([history, new_row], ignore_index=True)
    else:
        today = datetime.now().strftime("%Y-%m-%d")
        row = {"date": today}
        for r in results:
            if r["current_nav"] is not None:
                row[r["name"]] = r["current_nav"]
        # Don't write a row if no funds returned data
        if len(row) <= 1:
            print("  WARNING: No fund data to save. Skipping history update.")
            return history
        new_row = pd.DataFrame([row])
        history = history[history["date"] != today]
        history = pd.concat([history, new_row], ignore_index=True)

    # Sort by date and save
    history = history.sort_values("date").reset_index(drop=True)
    history.to_csv(HISTORY_FILE, index=False)
    return history


def build_change_history(history: pd.DataFrame) -> pd.DataFrame:
    """Calculate daily % changes from the NAV history."""
    if len(history) < 2:
        return pd.DataFrame()

    history = history.set_index("date").sort_index()
    changes = history.pct_change() * 100
    changes = changes.iloc[1:]  # Drop first row (NaN)
    return changes.round(4)


def write_excel(results: list[dict], history: pd.DataFrame):
    """Generate the Excel workbook with daily changes and history."""
    today = datetime.now().strftime("%Y-%m-%d")
    output_path = os.path.join(OUTPUT_DIR, f"nav_changes_{today}.xlsx")

    wb = Workbook()

    # --- Sheet 1: Daily NAV Changes ---
    ws1 = wb.active
    ws1.title = "Daily NAV Changes"

    # Header styling
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    green_font = Font(color="006100")
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_font = Font(color="9C0006")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    gray_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    headers = ["Fund Name", "Ticker", "NAV Date", "Current NAV", "Previous NAV",
               "Change ($)", "Change (%)"]
    for col, header in enumerate(headers, 1):
        cell = ws1.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    for row_idx, r in enumerate(results, 2):
        ws1.cell(row=row_idx, column=1, value=r["name"])
        ws1.cell(row=row_idx, column=2, value=r["ticker"])
        ws1.cell(row=row_idx, column=3, value=r["nav_date"] or "N/A")

        for col, key in [(4, "current_nav"), (5, "previous_nav"), (6, "change_dollar")]:
            cell = ws1.cell(row=row_idx, column=col,
                            value=r[key] if r[key] is not None else "N/A")
            if isinstance(r[key], (int, float)):
                cell.number_format = '#,##0.0000'

        # Change % with conditional formatting
        pct_val = r["change_pct"]
        cell = ws1.cell(row=row_idx, column=7,
                        value=pct_val if pct_val is not None else "N/A")
        if isinstance(pct_val, (int, float)):
            cell.number_format = '0.00"%"'
            if pct_val > 0:
                cell.font = green_font
                cell.fill = green_fill
            elif pct_val < 0:
                cell.font = red_font
                cell.fill = red_fill

        # Alternate row shading
        if row_idx % 2 == 0:
            for col in range(1, 8):
                c = ws1.cell(row=row_idx, column=col)
                if c.fill == PatternFill():  # Only if not already colored
                    c.fill = gray_fill

    # Column widths
    col_widths = [35, 18, 12, 14, 14, 14, 14]
    for i, w in enumerate(col_widths, 1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    # --- Sheet 2: History of Daily % Changes ---
    ws2 = wb.create_sheet("Change History")
    change_history = build_change_history(history)

    if not change_history.empty:
        # Headers
        ws2.cell(row=1, column=1, value="Date").font = header_font
        ws2.cell(row=1, column=1).fill = header_fill
        for col_idx, fund_name in enumerate(change_history.columns, 2):
            cell = ws2.cell(row=1, column=col_idx, value=fund_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", wrap_text=True)

        # Data rows (most recent first)
        for row_idx, (date, row_data) in enumerate(
                change_history.iloc[::-1].iterrows(), 2):
            ws2.cell(row=row_idx, column=1, value=date)
            for col_idx, val in enumerate(row_data, 2):
                cell = ws2.cell(row=row_idx, column=col_idx, value=val if pd.notna(val) else "")
                if pd.notna(val):
                    cell.number_format = '0.00"%"'
                    if val > 0:
                        cell.font = green_font
                        cell.fill = green_fill
                    elif val < 0:
                        cell.font = red_font
                        cell.fill = red_fill

        # Column widths
        ws2.column_dimensions["A"].width = 12
        for i in range(2, len(change_history.columns) + 2):
            ws2.column_dimensions[get_column_letter(i)].width = 20
    else:
        ws2.cell(row=1, column=1, value="No history available yet. Run again tomorrow.")

    wb.save(output_path)
    return output_path


def main():
    parser = argparse.ArgumentParser(description="Daily NAV Change Tracker")
    parser.add_argument(
        "--backfill", action="store_true",
        help="Fetch and store ~1 month of historical NAV data"
    )
    args = parser.parse_args()

    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    period = "2mo" if args.backfill else "10d"
    mode = "BACKFILL (past month)" if args.backfill else "Daily"

    print("=" * 60)
    print(f"Daily NAV Change Tracker - {mode}")
    print(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 60)
    print()

    # Fetch NAV data for all funds
    results = fetch_all_navs(period=period)

    # Update history
    print()
    print("Updating NAV history...")
    history = update_history(results, backfill=args.backfill)

    # Clean up closes from results before Excel (not needed there)
    for r in results:
        r.pop("closes", None)

    # Generate Excel
    print("Generating Excel report...")
    output_path = write_excel(results, history)

    # Summary
    print()
    print("=" * 60)
    print("SUMMARY")
    print("=" * 60)

    success = sum(1 for r in results if r["current_nav"] is not None)
    failed = sum(1 for r in results if r["current_nav"] is None)

    print(f"Funds fetched successfully: {success}/{len(results)}")
    if failed > 0:
        print(f"Funds FAILED: {failed}")
        for r in results:
            if r["current_nav"] is None:
                print(f"  - {r['name']}")

    print()
    print(f"Excel report: {output_path}")
    print(f"NAV history:  {HISTORY_FILE}")

    # Print quick table
    print()
    print(f"{'Fund':<35} {'NAV':>10} {'Change':>10} {'Chg %':>8}")
    print("-" * 65)
    for r in results:
        nav = f"{r['current_nav']:.4f}" if r["current_nav"] else "N/A"
        chg = f"{r['change_dollar']:+.4f}" if r["change_dollar"] is not None else "N/A"
        pct = f"{r['change_pct']:+.2f}%" if r["change_pct"] is not None else "N/A"
        print(f"{r['name']:<35} {nav:>10} {chg:>10} {pct:>8}")

    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
