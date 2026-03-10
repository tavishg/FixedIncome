"""
Fund configuration for the Daily NAV Change Tracker.

Each fund has:
  - name: Short display name for Excel output
  - tickers: List of Yahoo Finance tickers to try (in order of preference)
  - fund_type: "etf" or "mutual_fund"

Ticker formats:
  - ETFs on TSX use ".TO" suffix (e.g., XCB.TO)
  - Canadian mutual funds use Fundserv codes with ".CF" suffix (e.g., MMF659.CF)
  - Some mutual funds use Morningstar IDs with ".TO" suffix (e.g., 0P0001OE5B.TO)

If a fund's primary ticker doesn't work, update it here.
"""

FUNDS = [
    {
        "name": "Dynamic Global FI Fund F",
        "tickers": ["DYN1560F.CF", "0P0001OE5B.TO"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "Dynamic Credit Abs Return F",
        "tickers": ["DYN1753F.CF", "DYN27550.CF"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "Dynamic Short Term Credit PLUS F",
        "tickers": ["DYN3330F.CF", "0P0001OE4X.TO"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "Manulife Strategic Income F",
        "tickers": ["MMF659.CF", "0P0000NFNA.TO"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "Lysander-Canso Corp Value Bond F",
        "tickers": ["LYZ801F.CF", "0P0001QOXS.TO"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "PIMCO Monthly Income F",
        "tickers": ["PMO205.CF", "0P0000S9O5.TO"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "RBC Global Bond F",
        "tickers": ["RBF603.CF", "0P0000718L.TO"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "Mackenzie Unconstrained FI F",
        "tickers": ["MFC4765.CF", "0P0001K2BW.TO"],
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
        "tickers": ["PGF510.CF"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "DXDB Dynamic Discount Bond ETF",
        "tickers": ["DXDB.TO"],
        "fund_type": "etf",
    },
    {
        "name": "Lysander-Fulcra Corp Sec F",
        "tickers": ["LYZ935F.CF"],
        "fund_type": "mutual_fund",
    },
    {
        "name": "RBC Core Bond F",
        "tickers": ["RBF2603.CF", "0P0001F5OB.TO"],
        "fund_type": "mutual_fund",
    },
]
