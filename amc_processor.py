import io
import csv
import os
from datetime import datetime, date
import requests
import pyxirr

NAV_URL = "https://www.amfiindia.com/spages/NAVAll.txt"
NAV_FILENAME = "NAVAll.txt"
HISTNAV_URL = "https://api.mfapi.in/mf/"

def ensure_nav_file():
    """
    Ensures that NAVAll.txt is present in the current folder.
    If not, fetches it from NAV_URL and saves it.
    """
    if not os.path.exists(NAV_FILENAME):
        try:
            print("NAVAll.txt not found locally. Fetching from AMFI...")
            response = requests.get(NAV_URL)
            response.raise_for_status()
            with open(NAV_FILENAME, "w", newline='') as f:
                f.write(response.text)
            print("NAVAll.txt has been downloaded and saved.")
        except Exception as e:
            print(f"Error fetching NAVAll.txt: {e}")

def fetch_nav(ISIN):
    """
    Opens the local NAVAll.txt file (ensuring it exists) and searches for a line
    where the given ISIN appears in either the "ISIN Div Payout/ ISIN Growth" or
    "ISIN Div Reinvestment" field.
    
    Expected file schema (semicolon-separated):
      Scheme Code;ISIN Div Payout/ ISIN Growth;ISIN Div Reinvestment;Scheme Name;Net Asset Value;Date

    Returns:
      (nav, nav_date) if a match is found, where nav is a float and nav_date is a datetime object.
      Returns None if no match is found.
    """
    ensure_nav_file()
    nav_file_path = os.path.join(".", NAV_FILENAME)
    try:
        with open(nav_file_path, "r") as f:
            lines = f.read().splitlines()
    except Exception as e:
        print(f"Error opening NAVAll.txt: {e}")
        return None

    # Skip header; assume first line is header.
    for line in lines[1:]:
        parts = line.split(";")
        if len(parts) < 6:
            continue
        isin1 = parts[1].strip()
        isin2 = parts[2].strip()
        if ISIN == isin1 or ISIN == isin2:
            try:
                nav = float(parts[4].strip())
            except Exception as e:
                print(f"Error converting NAV for ISIN {ISIN}: {e}")
                return None
            try:
                nav_date = datetime.strptime(parts[5].strip(), '%d-%b-%Y')
            except Exception as e:
                print(f"Error parsing NAV date for ISIN {ISIN}: {e}")
                return None
            return (nav, nav_date)
    return None

def fetch_ebal(ISIN, folder):
    """
    Fetches the effective balance for the given ISIN.
    """
    filepath = os.path.join(folder, 'AMClist.csv')
    if not os.path.exists(filepath):
        return ""
    with open(filepath, 'r') as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row.get('isin') == ISIN:
                return (row.get('Ending Value', ''), row.get('cagr', ''))
            
def is_valid_balance(val_str):
    """
    Determines if val_str represents a valid numeric balance.
    Invalid values: blank, "na", "nan", "nill", "null", etc.
    """
    if not val_str:
        return False
    val_str = val_str.strip().lower()
    if val_str in ["", "na", "nan", "nill", "null"]:
        return False
    try:
        float(val_str)
        return True
    except:
        return False

def historical_nav(amfi):
    """
    Fetches historical NAV data for the given AMFI code.
    """
    url = f"{HISTNAV_URL}{amfi}"
    # print(f"Fetching historical NAV for AMFI {amfi} from {url}")
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        # print(f"Data received: {data.get('data', [])}")
        return data.get('data', [])
    except Exception as e:
        print(f"Error fetching historical NAV for AMFI {amfi}: {e}")
        return None
    

def process_amc_files(csv_data, folder):
    """
    Processes CSV data to create separate AMC CSV files and an updated AMClist.csv.

    Final CSV header (in amc_list_file):
      amc, filename, schemes, isin, amfi, valid, NAV, irr, Beginning Value, Ending Value, effective_balance, cagr
    
    For each distinct scheme within each AMC:
      - Finds the last valid numeric balance (effective_balance), its date, and ISIN from that transaction.
      - If effective_balance is nonzero, scheme is valid = "yes", else "no".
      - If valid, fetch NAV and compute:
          * beginning_value = (sum(investment transactions) - sum(other transactions))
          * ending_value = effective_balance * NAV
          * cagr = (ending_value / beginning_value)^(1/periods) - 1
          * irr = xirr from negative (investment) and positive (others) cashflows
      - Accumulates portfolio data for all valid schemes:
          portfolio_invested += beginning_value
          portfolio_current += ending_value
          etc.
    Lastly, appends a "Portfolio" row with the aggregated data.
    """
    csv_file = io.StringIO(csv_data)
    reader = csv.DictReader(csv_file)
    rows = list(reader)
    header = reader.fieldnames

    # Group rows by AMC; rows with empty AMC go to error_rows.
    grouped_by_amc = {}
    error_rows = []
    for row in rows:
        amc = row.get('amc')
        if not amc or amc.strip() == "":
            error_rows.append(row)
        else:
            amc = amc.strip()
            grouped_by_amc.setdefault(amc, []).append(row)

    investment_types = {"PURCHASE", "PURCHASE_SIP", "DIVIDEND_PAYOUT", "DIVIDEND_REINVESTMENT", "SWITCH_IN", "SWITCH_IN_MERGER"}
    amc_list = []  # List for AMClist.csv

    # Initialize portfolio accumulators.
    portfolio_invested = 0.0
    portfolio_current = 0.0
    portfolio_inv_dates = []
    portfolio_xirr_dates = []
    portfolio_xirr_amounts = []

    for amc, amc_rows in grouped_by_amc.items():
        safe_amc = amc.replace(" ", "_")
        amc_filename = f"{safe_amc}.csv"
        full_amc_path = os.path.join(folder, amc_filename)
        with open(full_amc_path, 'w', newline='') as amc_file:
            writer = csv.DictWriter(amc_file, fieldnames=header)
            writer.writeheader()
            writer.writerows(amc_rows)
        
        # Group rows by scheme within the current AMC.
        grouped_by_scheme = {}
        for row in amc_rows:
            scheme = row.get('scheme', '').strip()
            if scheme:
                grouped_by_scheme.setdefault(scheme, []).append(row)
        
        for scheme, scheme_rows in grouped_by_scheme.items():
            amfi_set = set()
            effective_balance = None
            effective_date = None
            effective_isin = None

            # For scheme-level calculations:
            inv_cash = 0.0
            other_cash = 0.0
            inv_dates = []
            xirr_dates = []
            xirr_amounts = []

            # 1) Identify last valid balance, accumulate flows
            for r in scheme_rows:
                # If balance is valid, update effective_balance, date, and ISIN
                bal_str = r.get('balance', '')
                if is_valid_balance(bal_str):
                    try:
                        effective_balance = float(bal_str)
                        date_str = r.get('date', '').strip()
                        try:
                            effective_date = datetime.strptime(date_str, '%Y-%m-%d')
                        except:
                            effective_date = None
                        isin_candidate = r.get('isin', '').strip()
                        if isin_candidate:
                            effective_isin = isin_candidate
                    except:
                        pass
                
                # Aggregate AMFI
                if r.get('amfi') and r.get('amfi').strip():
                    amfi_set.add(r.get('amfi').strip())
                
                # 2) Accumulate flows for CAGR & XIRR
                amt_str = r.get('amount', '').strip()
                try:
                    amt = float(amt_str)
                except:
                    amt = 0.0
                date_str = r.get('date', '').strip()
                # print(f"Date_str: {date_str}")
                try:
                    trans_date = datetime.strptime(date_str, '%Y-%m-%d')
                except:
                    trans_date = None
                if r.get('type', '').strip().upper() in investment_types:
                    inv_cash += amt
                    if trans_date:
                        inv_dates.append(trans_date)
                        # print(f"trans_date: {trans_date}")
                        xirr_dates.append(trans_date)
                        xirr_amounts.append(-amt)
                else:
                    other_cash += amt
                    if trans_date:
                        xirr_dates.append(trans_date)
                        xirr_amounts.append(amt)
            
            valid = "yes" if (effective_balance is not None and effective_balance != 0) else "no"

            # 3) If valid, fetch NAV
            nav_value = ""
            nav_date = None
            if valid == "yes" and effective_isin:
                nav_record = fetch_nav(effective_isin)
                if nav_record:
                    nav_value, nav_date = nav_record

            # 4) Calculate metrics if valid & have nav
            beginning_value = 0.0
            ending_value = 0.0
            cagr = ""
            irr = ""
            if valid == "yes" and nav_value:
                beginning_value = inv_cash - other_cash
                ending_value = (effective_balance or 0.0) * nav_value

                if inv_dates:
                    first_date = min(inv_dates)
                    # print(f"First date: {first_date}")
                    last_date = max(inv_dates)
                    # print(f"Last date: {last_date}")
                    if last_date == first_date: 
                        last_date = datetime.strptime(date.strftime(date.today(), '%Y-%m-%d'), '%Y-%m-%d')
                        # print(f"Last date: {last_date}")

                    periods = (last_date - first_date).days / 365.25
                    # print(f"Scheme: {scheme}, Periods: {periods}, last_date: {last_date}, first_date: {first_date}")
                    if beginning_value > 0 and periods > 0:
                        try:
                            cagr_calc = (ending_value / beginning_value) ** (1 / periods) - 1
                            cagr = round(cagr_calc, 4)
                        except Exception as e:
                            print(f"Error calculating CAGR for {scheme}: {e}")
                            cagr = ""
                    # If no positive flow, append final positive at nav_date
                    if not any(a > 0 for a in xirr_amounts) and nav_date:
                        xirr_dates.append(nav_date)
                        xirr_amounts.append(ending_value)
                    try:
                        xirr_result = pyxirr.xirr(xirr_dates, xirr_amounts)
                        if xirr_result is not None:
                            irr = round(xirr_result, 4)
                    except Exception as e:
                        print(f"Error calculating XIRR for {scheme}: {e}")
                        irr = ""
                
                # 5) Update portfolio accumulators
                portfolio_invested += beginning_value
                portfolio_current += ending_value
                portfolio_inv_dates.extend(inv_dates)
                if nav_date:
                    portfolio_xirr_dates.append(nav_date)
                    portfolio_xirr_amounts.append(ending_value)

            amfi_field = ",".join(sorted(amfi_set)) if amfi_set else ""
            
            # 6) Append scheme row to amc_list with new columns
            amc_list.append({
                'amc': amc,
                'filename': f"{safe_amc}.csv",
                'schemes': scheme,
                'isin': effective_isin or "",
                'amfi': amfi_field,
                'valid': valid,
                'NAV': f"{nav_value}" if nav_value else "",
                'irr': f"{irr}" if irr else "",
                'Beginning Value': f"{round(beginning_value, 4)}" if beginning_value else "0.0",
                'Ending Value': f"{round(ending_value, 4)}" if ending_value else "0.0",
                'effective_balance': f"{effective_balance}" if effective_balance else "",
                'cagr': f"{cagr}" if cagr else ""
            })
    
    # 7) Portfolio-level calculations
    portfolio_cagr = ""
    portfolio_irr = ""
    if portfolio_inv_dates:
        port_first_date = min(portfolio_inv_dates)
        port_last_date = max(portfolio_inv_dates)
        port_periods = (port_last_date - port_first_date).days / 365.25
        if portfolio_invested > 0 and port_periods > 0:
            try:
                pcagr = (portfolio_current / portfolio_invested) ** (1 / port_periods) - 1
                portfolio_cagr = round(pcagr, 4)
            except Exception as e:
                print(f"Error calculating Portfolio CAGR: {e}")
                portfolio_cagr = ""
    # If no positive flow, append final positive at nav_date if available
    # This is optional; if you want to ensure a redemption-like flow for XIRR
    # you'd need the final nav_date from the last scheme. 
    # For now, we skip unless you store nav_date from the last valid scheme somewhere.
    try:
        pxirr_result = pyxirr.xirr(portfolio_xirr_dates, portfolio_xirr_amounts)
        if pxirr_result is not None:
            portfolio_irr = round(pxirr_result, 4)
    except Exception as e:
        print(f"Error calculating Portfolio XIRR: {e}")
        portfolio_irr = ""
    
    # 8) Append portfolio row
    portfolio_row = {
        'amc': "",
        'filename': "",
        'schemes': "Portfolio",
        'isin': "",
        'amfi': f"{round(portfolio_invested, 4)}",   # store invested in amfi column
        'valid': "yes",
        'NAV': f"{round(portfolio_current, 4)}",
        'irr': f"{portfolio_irr}" if portfolio_irr else "",
        'Beginning Value': f"{round(portfolio_invested, 4)}",
        'Ending Value': f"{round(portfolio_current, 4)}",
        'effective_balance': "",
        'cagr': f"{portfolio_cagr}" if portfolio_cagr else ""
    }
    amc_list.append(portfolio_row)

    # 9) Write out AMClist.csv with the new header
    amc_list_file = os.path.join(folder, "AMClist.csv")
    with open(amc_list_file, 'w', newline='') as list_file:
        fieldnames = [
            'amc',
            'filename',
            'schemes',
            'isin',
            'amfi',
            'valid',
            'NAV',
            'irr',
            'Beginning Value',
            'Ending Value',
            'effective_balance',
            'cagr'
        ]
        writer = csv.DictWriter(list_file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(amc_list)

    # 10) Write error rows if any
    if error_rows:
        error_file = os.path.join(folder, "errorlist.csv")
        with open(error_file, 'w', newline='') as err_file:
            writer = csv.DictWriter(err_file, fieldnames=header)
            writer.writeheader()
            writer.writerows(error_rows)
