"""
Fund configuration for the Daily NAV Change Tracker.

Each fund has:
  - name: Short display name for Excel output
  - tickers: List of Yahoo Finance tickers to try (in order of preference)
  - fund_type: "etf" or "mutual_fund"

Ticker formats:
  - ETFs on TSX use ".TO" suffix (e.g., XCB.TO)
  - Some mutual funds use long-form names with ".TO" suffix (e.g., MANUVIEREVST.TO)
  - Some mutual funds use Morningstar IDs with ".TO" suffix (e.g., 0P0000MOQD.TO)

NOTE: The ".CF" suffix (Globe & Mail / CADFUNDS format) does NOT work with yfinance.
If a fund's primary ticker doesn't work, update it here.
"""

FUNDS = [
    {
        "name": "Dynamic Global FI Fund F",
        "tickers": ["DXBG.TO", "0P0001OE5B.TO"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "Dynamic Credit Abs Return F",
        "tickers": ["DYN27550.TO", "0P0001I6W0.TO"],
        "fund_type": "mutual_fund",
        "morningstar_id": "0P0001I6W0",
    },
    {
        "name": "Dynamic Short Term Credit PLUS F",
        "tickers": ["DXCP.TO", "0P0001OE4X.TO"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "Manulife Strategic Income F",
        "tickers": ["0P0000NFNA.TO"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "Lysander-Canso Corp Value Bond F",
        "tickers": ["0P0000XXNG.TO"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "PIMCO Monthly Income F",
        "tickers": ["PMIF.TO", "0P0000S9O5.TO"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "RBC Global Bond F",
        "tickers": ["0P0000718L.TO"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "Mackenzie Unconstrained FI F",
        "tickers": ["0P0001K2BW.TO"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "XIG US IG Corp Bond (CAD Hdg)",
        "tickers": ["XIG.TO"],
        "fund_type": "etf",
    },
    {
        "name": "XCB Cdn Corporate Bond",
        "tickers": ["XCB.TO"],
        "fund_type": "etf",
    },
    {
        "name": "Pender Corporate Bond F",
        "tickers": ["0P0000MOQD.TO", "PENDERBONDD.TO"],
        "fund_type": "mutual_fund",
        "morningstar_id": "0P0000MOQD",
    },
    {
        "name": "DXDB Dynamic Discount Bond ETF",
        "tickers": ["DXDB.TO"],
        "fund_type": "etf",
    },
    {
        "name": "Lysander-Fulcra Corp Sec F",
        "tickers": ["0P0001CIH7.TO"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "RBC Core Bond F",
        "tickers": ["0P0001F5OB.TO"],
        "fund_type": "mutual_fund",
    },
]
