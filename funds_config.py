"""
Fund configuration for the Daily NAV Change Tracker.

Each fund has:
  - name: Short display name for Excel output
  - tickers: List of Yahoo Finance tickers to try (in order of preference)
  - fund_type: "etf" or "mutual_fund"
  - morningstar_ids: (optional) Morningstar security IDs for chart API fallback
  - morningstar_search: (optional) search term for Morningstar screener API fallback

Ticker formats:
  - ETFs on TSX use ".TO" suffix (e.g., XCB.TO)
  - Some mutual funds use long-form names with ".TO" suffix
  - Some mutual funds use Morningstar IDs with ".TO" suffix (e.g., 0P0000MOQD.TO)
  - US-listed mutual fund tickers (e.g., DBZBX) work without suffix

NOTE: The ".CF" suffix (Globe & Mail / CADFUNDS format) does NOT work with yfinance.
If a fund's primary ticker doesn't work, update it here.

Order: Dynamic funds first, then Lysander, then Pender, then others.
"""

FUNDS = [
    # --- Dynamic Funds ---
    {
        "name": "Dynamic Global FI Fund F",
        "tickers": ["DXBG.TO", "0P0001OE5B.TO"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "Dynamic Credit Abs Return F",
        "tickers": [
            "0P0001ROZ4.TO",   # Series F (correct Morningstar ID)
            "DYN27550.TO",     # Series F (FundSERV code)
            "DBZBX",           # US Nasdaq listing (Series F1 NL)
        ],
        "fund_type": "mutual_fund",
        "morningstar_ids": ["0P0001ROZ4"],
        # Fallback: search Morningstar screener by name to find correct securityID
        "morningstar_search": "Dynamic Credit Absolute Return",
        # Last resort: scrape TMX Money quote page
        "tmx_symbol": "DYN27550",
    },
    {
        "name": "Dynamic Short Term Credit PLUS F",
        "tickers": ["DXCP.TO", "0P0001OE4X.TO"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "DXDB Dynamic Discount Bond ETF",
        "tickers": ["DXDB.TO"],
        "fund_type": "etf",
    },
    # --- Lysander Funds ---
    {
        "name": "Lysander-Canso Corp Value Bond F",
        "tickers": ["0P0000XXNG.TO"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "Lysander-Fulcra Corp Sec F",
        "tickers": ["0P0001CIH7.TO"],
        "fund_type": "mutual_fund",
    },
    # --- Pender ---
    {
        "name": "Pender Corporate Bond F",
        # Series F tickers first, then Series D/A as fallbacks (same portfolio,
        # daily % changes are nearly identical — only MER differs slightly)
        "tickers": [
            "0P0000MOQD.TO",   # Series F (Morningstar ID)
            "PENDERBONDD.TO",  # Series F (long-form)
            "F00000W3SV.TO",   # Series D (same portfolio, lower MER than A)
            "PENDERCORPOR.TO", # Series A
            "0P0000MOQB.TO",   # Series A (Morningstar ID)
        ],
        "fund_type": "mutual_fund",
        "morningstar_ids": ["0P0000MOQD", "0P0000TISC"],
    },
    # --- ETFs ---
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
    # --- Other Mutual Funds ---
    {
        "name": "PIMCO Monthly Income F",
        "tickers": ["PMIF.TO", "0P0000S9O5.TO"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "Manulife Strategic Income F",
        "tickers": ["0P0000NFNA.TO"],
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
        "name": "RBC Core Bond F",
        "tickers": ["0P0001F5OB.TO"],
        "fund_type": "mutual_fund",
    },
]
