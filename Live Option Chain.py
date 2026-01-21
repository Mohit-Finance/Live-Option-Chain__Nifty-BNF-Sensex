import os
import sys
import time
import json
import threading
import pickle
from datetime import datetime, time as tm

import requests
import pandas as pd
import numpy as np
import xlwings as xw
import pyotp
import upstox_client

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import undetected_chromedriver as uc


tdate = datetime.now().date()
code = None
access = None

os.makedirs(f'Credentials/Data/{tdate}', exist_ok=True)

def time_fun():
    ttime = datetime.now().time().replace(microsecond=0)
    ttime = ttime.strftime("%H:%M:%S")
    return ttime

def show_totp(secret):
    totp = pyotp.TOTP(secret)
    otp = totp.now()
    return otp

if not os.path.exists('Credentials/login_details.json'):
    print("User Details not found. First Create a User Base & Retry. Exiting program.")
    sys.exit()

with open('Credentials/login_details.json', 'r') as file_read:
    users_data = json.load(file_read)

allowed_namess = users_data.keys()
allowed_names = [name.lower() for name in allowed_namess]

name_dict = {}

for i in range(len(allowed_names)):
    name_dict[f'{allowed_names[i]}'] = f'{tdate}_access_code_{allowed_names[i]}.json'

name_list = name_dict.values()

file_list = os.listdir(f'Credentials/Data/{tdate}')

for name in name_list:
    if name in file_list:
        with open(f'Credentials/Data/{tdate}/{name}', 'r') as file_read:
            access = json.load(file_read)
            acc_name = name[23:][:-5]

if not access:

    while True:
        acc_name = input(f'\nEnter Name of Account Holder to Login From {list(allowed_namess)} : ').lower()
        if acc_name in allowed_names:
            break
        else:
            print(f"\nInvalid User. Please Enter Registered User Name {list(allowed_namess)}'.")

    try:
        with open(f'Credentials/Data/{tdate}/{tdate}_access_code_{acc_name}.json', 'r') as file_read:
            access = json.load(file_read)

    except:

        with open('Credentials/login_details.json', 'r') as file_read:
            login_details = json.load(file_read)

        api_key = login_details[f'{acc_name.capitalize()}']['api_key']
        api_secret = login_details[f'{acc_name.capitalize()}']['api_secret']
        api_auth = login_details[f'{acc_name.capitalize()}']['api_auth']
        api_pin = login_details[f'{acc_name.capitalize()}']['pin']
        mobile_no = login_details[f'{acc_name.capitalize()}']['Mob No.']
        hold_name = login_details[f'{acc_name.capitalize()}']['full_name']

        print(f'\nTrying to Login from Account Holder: {hold_name}')

        uri = 'https://www.google.com/'
        url1 = f'https://api.upstox.com/v2/login/authorization/dialog?response_type=code&client_id={api_key}&redirect_uri={uri}\n'

        options = uc.ChromeOptions()
        options.headless = True
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        driver = uc.Chrome(options=options)

        # driver = uc.Chrome() # Use this line instead to run Chrome in normal (visible) mode, (In that case, comment out the 5 lines above that set headless options)

        driver.get(url1)
        wait = WebDriverWait(driver, 20)
        phone_input = wait.until(EC.presence_of_element_located((By.ID, "mobileNum")))
        phone_input.send_keys(mobile_no)
        otp_button = wait.until(EC.element_to_be_clickable((By.ID, "getOtp")))
        otp_button.click()
        # print("✅ Phone number entered, now captcha should appear normally")

        totp_value = show_totp(api_auth)
        totp_input = wait.until(EC.presence_of_element_located((By.ID, "otpNum")))
        totp_input.send_keys(totp_value)
        proceed_button = wait.until(EC.element_to_be_clickable((By.ID, "continueBtn")))
        proceed_button.click()
        # print("✅ TOTP entered and Continue clicked!")

        pin_input = wait.until(EC.presence_of_element_located((By.ID, "pinCode")))
        pin_input.send_keys(api_pin)
        proceed_button = wait.until(EC.element_to_be_clickable((By.ID, "pinContinueBtn")))
        proceed_button.click()

        # print("✅ PIN entered and proceed button clicked!")
        time.sleep(3)
        code_url = driver.current_url

        driver.quit()

        start = code_url.find('code=')
        if start != -1:
            start =start + 5  # move past 'code='
            code = code_url[start:start+6]
        else:
            print("No code found in the URL")

        url = 'https://api.upstox.com/v2/login/authorization/token'
        headers = {
            'accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded',
        }

        data = {
            'code': code,
            'client_id': api_key,
            'client_secret': api_secret,
            'redirect_uri': uri,
            'grant_type': 'authorization_code',
        }

        response = requests.post(url, headers=headers, data=data)
        access = response.json()['access_token']
        print(f'\nLogin Successful, Status Code : {response.status_code}')
        print(f"User Name : {response.json()['user_name']}\nEmail ID : {response.json()['email']}")

        with open(f'Credentials/Data/{tdate}/{tdate}_access_code_{acc_name}.json', 'w') as file_write:
            json.dump(access, file_write)

print(f'\nLogin Successful from Account : {acc_name.capitalize()}')


#############################################################################

def instrument():
    inst_url = 'https://assets.upstox.com/market-quote/instruments/exchange/complete.csv.gz'
    instrument = pd.read_csv(inst_url)
    instrument.to_csv('Credentials/instrument.csv')

if os.path.exists('Credentials/instrument.csv'):
    modified_time = datetime.fromtimestamp(os.path.getmtime('Credentials/instrument.csv'))
    today_9am = datetime.combine(datetime.today(), tm(9, 0, 0))
    yn = '1' if modified_time < today_9am else '0'
else:
    yn = '1'

if yn=='1':
    instrument()
    print("Instrument Data Updated Successfully")
try:
    inst_data = pd.read_csv('Credentials/instrument.csv', index_col=0)
    inst_data['expiry'] = pd.to_datetime(inst_data['expiry'], format='%Y-%m-%d', errors='coerce')
except:
    instrument()
    print("Can't find 'Instrument.csv' file, Latest Instrument Data Downloaded Successfully")
    inst_data = pd.read_csv('Credentials/instrument.csv', index_col=0)
    inst_data['expiry'] = pd.to_datetime(inst_data['expiry'], format='%Y-%m-%d', errors='coerce')

t_time = datetime.now().time().replace(microsecond=0)
start_time = tm(9,15,5,0)
end_time = tm(15,30,0,0)

while t_time < start_time:
    t_time = datetime.now().time().replace(microsecond=0)
    print(f'\rCurrent Time : {t_time} | Market Will Start at {start_time}', end='', flush=True)
    time.sleep(1)

configuration = upstox_client.Configuration()
configuration.access_token = access
data_base = {}
streamer = None
lock = threading.Lock()

def on_open():
    print("✅ WebSocket Connection Established")

def on_message(message):
    global sub_list_ce, sub_list_pe, inst_strike_pair
    if 'feeds' not in message:
        return

    data = message['feeds']
    for key, value in data.items():
        try:
            if "marketFF" in value.get("fullFeed", {}):
                ff = value['fullFeed']['marketFF']
                with lock:
                    data_base[key] = {
                        'ltp': ff.get('ltpc', {}).get('ltp') or None,
                        'type': 'CE' if key in sub_list_ce else 'PE',
                        'strike': inst_strike_pair.loc[key, 'strike'] if key in inst_strike_pair.index else None,
                        'prev_close': ff.get('ltpc', {}).get('cp') or None,
                        'delta': ff.get('optionGreeks', {}).get('delta') or None,
                        'theta': ff.get('optionGreeks', {}).get('theta') or None,
                        'gamma': ff.get('optionGreeks', {}).get('gamma') or None,
                        'vega': ff.get('optionGreeks', {}).get('vega') or None,
                        'volume': (ff.get('marketOHLC', {}).get('ohlc', [{}])[0].get('vol')) or None,
                        'atp': ff.get('atp') or None,
                        'tbq': ff.get('tbq') or None,
                        'tsq': ff.get('tsq') or None,
                        'oi': ff.get('oi') or None,
                        'iv': ff.get('iv') or None}

            elif "indexFF" in value.get("fullFeed", {}):
                ff = value['fullFeed']['indexFF']
                with lock:
                    data_base[key] = {'ltp': ff.get('ltpc', {}).get('ltp') or None}

        except Exception as e:
            print(f'Missing Key in {key} : {e}')
            continue

def start_stream():
    global sub_list, configuration, streamer, index_key

    streamer = upstox_client.MarketDataStreamerV3(upstox_client.ApiClient(configuration), sub_list+index_key, "full")
    streamer.on("open", on_open)
    streamer.on("message", on_message)
    streamer.connect()

def main():
    thread = threading.Thread(target=start_stream)
    thread.start()

apiInstance = upstox_client.MarketQuoteV3Api(upstox_client.ApiClient(configuration))

nifty_expiry = None
bnf_expiry = None
sensex_expiry = None
index_key = ["NSE_INDEX|Nifty 50", "NSE_INDEX|Nifty Bank", "BSE_INDEX|SENSEX", "NSE_INDEX|India VIX"]

def synth_atm_index():

    global nifty_expiry, bnf_expiry, sensex_expiry

    nifty_ce_ltp = nifty_pe_ltp = None
    bnf_ce_ltp = bnf_pe_ltp = None
    sensex_ce_ltp = sensex_pe_ltp = None

    index_key = "NSE_INDEX|Nifty 50,NSE_INDEX|Nifty Bank,BSE_INDEX|SENSEX"
    index_spot = apiInstance.get_ltp(instrument_key=index_key).to_dict()

    nifty_spot = index_spot['data']['NSE_INDEX:Nifty 50']['last_price']
    bnf_spot = index_spot['data']['NSE_INDEX:Nifty Bank']['last_price']
    sensex_spot = index_spot['data']['BSE_INDEX:SENSEX']['last_price']

    nifty_contract = contract_keys(index_spot = nifty_spot, step_size = 50, exchange = 'NSE_FO', name = 'NIFTY')
    bnf_contract =  contract_keys(index_spot = bnf_spot, step_size = 100, exchange = 'NSE_FO', name = 'BANKNIFTY')
    sensex_contract =  contract_keys(index_spot = sensex_spot, step_size = 100, exchange = 'BSE_FO', name = 'SENSEX')

    all_contract_keys = f"{nifty_contract[0]},{nifty_contract[1]},{bnf_contract[0]},{bnf_contract[1]},{sensex_contract[0]},{sensex_contract[1]}"
    
    nifty_atm = nifty_contract[2]
    bnf_atm = bnf_contract[2]
    sensex_atm = sensex_contract[2]

    nifty_expiry = nifty_contract[3]
    bnf_expiry = bnf_contract[3]
    sensex_expiry = sensex_contract[3]

    contract_quote = apiInstance.get_ltp(instrument_key=all_contract_keys).to_dict()

    for key, value in contract_quote['data'].items():
        # ✅ NIFTY
        if key.endswith("CE") and "NSE_FO:NIFTY" in key:
            nifty_ce_ltp = value['last_price']
        elif key.endswith("PE") and "NSE_FO:NIFTY" in key:
            nifty_pe_ltp = value['last_price']

        # ✅ BANKNIFTY
        elif key.endswith("CE") and "NSE_FO:BANKNIFTY" in key:
            bnf_ce_ltp = value['last_price']
        elif key.endswith("PE") and "NSE_FO:BANKNIFTY" in key:
            bnf_pe_ltp = value['last_price']

        # ✅ SENSEX
        elif key.endswith("CE") and "BSE_FO:SENSEX" in key:
            sensex_ce_ltp = value['last_price']
        elif key.endswith("PE") and "BSE_FO:SENSEX" in key:
            sensex_pe_ltp = value['last_price']


    if None in [nifty_ce_ltp, nifty_pe_ltp, bnf_ce_ltp, bnf_pe_ltp, sensex_ce_ltp, sensex_pe_ltp]:
        raise ValueError("Some CE/PE values missing, check response or instrument_key mapping")

    nifty_synthetic_spot = nifty_ce_ltp - nifty_pe_ltp + nifty_atm
    nifty_synthetic_atm = round(nifty_synthetic_spot/50) * 50

    bnf_synthetic_spot = bnf_ce_ltp - bnf_pe_ltp + bnf_atm
    bnf_synthetic_atm = round(bnf_synthetic_spot/100) * 100

    sensex_synthetic_spot = sensex_ce_ltp - sensex_pe_ltp + sensex_atm
    sensex_synthetic_atm = round(sensex_synthetic_spot/100) * 100

    index_synthetic_atm = {'nifty':nifty_synthetic_atm, 'bnf':bnf_synthetic_atm, 'sensex':sensex_synthetic_atm}

    return index_synthetic_atm


def contract_keys(index_spot, step_size, exchange, name):

    index_atm = round(index_spot/step_size) * step_size
    index_dataframe = inst_data[(inst_data['exchange'] == exchange) & (inst_data['instrument_type'] == "OPTIDX") & (inst_data['name'] == name)]
    index_expiry = sorted(index_dataframe['expiry'].unique().tolist())

    index_atm_strikes_df = inst_data[(inst_data['exchange'] == exchange) & (inst_data['instrument_type'] == "OPTIDX") & (inst_data['name'] == name) & (inst_data['expiry'] == index_expiry[0]) & (inst_data['strike'] == index_atm)]
    index_atm_ce_instkey = index_atm_strikes_df[index_atm_strikes_df['option_type'] == 'CE']['instrument_key'].iloc[0]
    index_atm_pe_instkey = index_atm_strikes_df[index_atm_strikes_df['option_type'] == 'PE']['instrument_key'].iloc[0]

    return [index_atm_ce_instkey, index_atm_pe_instkey, index_atm, index_expiry]

synth_atm = synth_atm_index()

####### Upto Here we have brought Synthetic ATM Strike for each index Nifty, BankNifty, Sensex #######

old_synth_atm = synth_atm


def valid_strikes(synth_atm):
    global inst_data

    nifty_upper_strikes = []
    nifty_lower_strikes = []
    bnf_upper_strikes = []
    bnf_lower_strikes = []
    sensex_upper_strikes = []
    sensex_lower_strikes = []

    total_strikes = 15

    for i in range (1, total_strikes+1):
        nifty_upper_strikes.append(synth_atm['nifty'] + (i)*50)
        bnf_upper_strikes.append(synth_atm['bnf'] + (i)*100)
        sensex_upper_strikes.append(synth_atm['sensex'] + (i)*100)

        nifty_lower_strikes.append(synth_atm['nifty'] - (i)*50)
        bnf_lower_strikes.append(synth_atm['bnf'] - (i)*100)
        sensex_lower_strikes.append(synth_atm['sensex'] - (i)*100)

    nifty_option_strikes = nifty_lower_strikes + [synth_atm['nifty']] + nifty_upper_strikes
    bnf_option_strikes = bnf_lower_strikes + [synth_atm['bnf']] + bnf_upper_strikes
    sensex_option_strikes = sensex_lower_strikes + [synth_atm['sensex']] + sensex_upper_strikes
    ####### Got all the Strikes required to prepare Option Chain of Nifty, BNF and Sensex as per Synthetic Strikes of each Index #######

    nifty_option_inst_ce = inst_data[(inst_data['exchange'] == 'NSE_FO') & (inst_data['instrument_type'] == 'OPTIDX') & (inst_data['name'] == 'NIFTY') & (inst_data['expiry'] == nifty_expiry[0]) & (inst_data['option_type'] == 'CE') & (inst_data['strike'].isin(nifty_option_strikes))]['instrument_key'].tolist()
    nifty_option_inst_pe = inst_data[(inst_data['exchange'] == 'NSE_FO') & (inst_data['instrument_type'] == 'OPTIDX') & (inst_data['name'] == 'NIFTY') & (inst_data['expiry'] == nifty_expiry[0]) & (inst_data['option_type'] == 'PE') & (inst_data['strike'].isin(nifty_option_strikes))]['instrument_key'].tolist()
    bnf_option_inst_ce = inst_data[(inst_data['exchange'] == 'NSE_FO') & (inst_data['instrument_type'] == 'OPTIDX') & (inst_data['name'] == 'BANKNIFTY') & (inst_data['expiry'] == bnf_expiry[0]) & (inst_data['option_type'] == 'CE') &(inst_data['strike'].isin(bnf_option_strikes))]['instrument_key'].tolist()
    bnf_option_inst_pe = inst_data[(inst_data['exchange'] == 'NSE_FO') & (inst_data['instrument_type'] == 'OPTIDX') & (inst_data['name'] == 'BANKNIFTY') & (inst_data['expiry'] == bnf_expiry[0]) & (inst_data['option_type'] == 'PE') &(inst_data['strike'].isin(bnf_option_strikes))]['instrument_key'].tolist()
    sensex_option_inst_ce = inst_data[(inst_data['exchange'] == 'BSE_FO') & (inst_data['instrument_type'] == 'OPTIDX') & (inst_data['name'] == 'SENSEX') & (inst_data['expiry'] == sensex_expiry[0]) & (inst_data['option_type'] == 'CE') & (inst_data['strike'].isin(sensex_option_strikes))]['instrument_key'].tolist()
    sensex_option_inst_pe = inst_data[(inst_data['exchange'] == 'BSE_FO') & (inst_data['instrument_type'] == 'OPTIDX') & (inst_data['name'] == 'SENSEX') & (inst_data['expiry'] == sensex_expiry[0]) & (inst_data['option_type'] == 'PE') & (inst_data['strike'].isin(sensex_option_strikes))]['instrument_key'].tolist()
    
    index_ce_pe_list = [nifty_option_inst_ce, nifty_option_inst_pe, bnf_option_inst_ce, bnf_option_inst_pe, sensex_option_inst_ce, sensex_option_inst_pe]

    sub_list_ce = nifty_option_inst_ce + bnf_option_inst_ce + sensex_option_inst_ce
    sub_list_pe = nifty_option_inst_pe + bnf_option_inst_pe + sensex_option_inst_pe
    sub_list = sub_list_ce + sub_list_pe

    inst_strike_pair = inst_data[inst_data['instrument_key'].isin(sub_list)][['instrument_key', 'strike']].set_index('instrument_key')

    return sub_list_ce, sub_list_pe, sub_list, inst_strike_pair, index_ce_pe_list

sub_list_ce, sub_list_pe, sub_list, inst_strike_pair, index_ce_pe_list = valid_strikes(old_synth_atm)

if __name__ == "__main__":
    main()

while not data_base:
    print("⏳ Waiting for WebSocket data...")
    time.sleep(0.5)

print("✅ Local Database Updated!")

index_key = ["NSE_INDEX|Nifty 50", "NSE_INDEX|Nifty Bank", "BSE_INDEX|SENSEX", "NSE_INDEX|India VIX"]

initial_oi_data = {f'nifty_oi_ce_initial':None, f'nifty_oi_pe_initial':None, 'bnf_oi_ce_initial':None, f'bnf_oi_pe_initial':None, 'sensex_oi_ce_initial':None, f'sensex_oi_pe_initial':None,}
def option_chain(ce, pe, df, expiry, index):
    global initial_oi_data
    df_index = df[df.index.isin(ce + pe)]
    df_ce = df_index[df_index['type'] == 'CE']
    df_pe = df_index[df_index['type'] == 'PE']
    df_option = pd.merge(df_ce, df_pe, on='strike', suffixes=('_CE', '_PE')).sort_values(by='strike', ascending=True)
    df_option['expiry'] = expiry
    df_option['iv_CE'] = df_option['iv_CE']*100
    df_option['iv_PE'] = df_option['iv_PE']*100

    df_option[f'{index}_spot'] = df.loc[{'nifty': 'NSE_INDEX|Nifty 50', 'bnf': 'NSE_INDEX|Nifty Bank', 'sensex': 'BSE_INDEX|SENSEX'}[index], 'ltp']
    india_vix = df.loc['NSE_INDEX|India VIX', 'ltp']

    df_option['diff'] = abs(df_option['ltp_CE'] - df_option['ltp_PE'])
    df_option['Δ_CE'] = round(((df_option['ltp_CE'] - df_option['prev_close_CE']) / df_option['prev_close_CE']),2)
    df_option['Δ_PE'] = round(((df_option['ltp_PE'] - df_option['prev_close_PE']) / df_option['prev_close_PE']),2)

    df_option['buy_per_CE'] = (df_option['tbq_CE'])/(df_option['tbq_CE'] + df_option['tsq_CE'])*100
    df_option['sell_per_CE'] = (df_option['tsq_CE'])/(df_option['tbq_CE'] + df_option['tsq_CE'])*100
    
    df_option['buy_per_PE'] = (df_option['tbq_PE'])/(df_option['tbq_PE'] + df_option['tsq_PE'])*100
    df_option['sell_per_PE'] = (df_option['tsq_PE'])/(df_option['tbq_PE'] + df_option['tsq_PE'])*100

    df_option['CE_depth'] = (df_option['buy_per_CE'].round(2).astype(str) + '%' + ' / ' + df_option['sell_per_CE'].round(2).astype(str) + '%')
    df_option['PE_depth'] = (df_option['buy_per_PE'].round(2).astype(str) + '%' + ' / ' + df_option['sell_per_PE'].round(2).astype(str) + '%')
    df_option['India_Vix'] = india_vix


    if initial_oi_data[f'{index}_oi_ce_initial'] is None:
        try:
            with open(f'Credentials/Data/{tdate}/init_oi_{index}_{tdate}.pkl', 'rb') as fileread:
                initial_oi_data = pickle.load(fileread)
                df_option['oi_CE_initial'] = initial_oi_data[f'{index}_oi_ce_initial']
                df_option['oi_PE_initial'] = initial_oi_data[f'{index}_oi_pe_initial']
        except:
            initial_oi_data[f'{index}_oi_ce_initial'] = df_option['oi_CE']
            initial_oi_data[f'{index}_oi_pe_initial'] = df_option['oi_PE']
            with open(f'Credentials/Data/{tdate}/init_oi_{index}_{tdate}.pkl', 'wb') as filewrite:
                pickle.dump(initial_oi_data, filewrite)
                df_option['oi_CE_initial'] = initial_oi_data[f'{index}_oi_ce_initial']
                df_option['oi_PE_initial'] = initial_oi_data[f'{index}_oi_pe_initial']
    
    df_option['oi_CE_initial'] = initial_oi_data[f'{index}_oi_ce_initial']
    df_option['oi_PE_initial'] = initial_oi_data[f'{index}_oi_pe_initial'] 
    # print(df_option, index)

    df_option['Δ_OI_CE'] = (((df_option['oi_CE'] - df_option['oi_CE_initial']) / df_option['oi_CE_initial'])).round(2)
    df_option['Δ_OI_PE'] = (((df_option['oi_PE'] - df_option['oi_PE_initial']) / df_option['oi_PE_initial'])).round(2)

    # df_option['B/S CE'] = round((df_option['tbq_CE'] / df_option['tsq_CE']),2)
    # df_option['B/S PE'] = round((df_option['tbq_PE'] / df_option['tsq_PE']),2)
    df_option = df_option[['expiry', 'prev_close_CE', 'delta_CE', 'theta_CE', 'gamma_CE', 'vega_CE', 'atp_CE', 'tbq_CE', 'tsq_CE', 'CE_depth', 'oi_CE', 'iv_CE', 'volume_CE', 'Δ_OI_CE','Δ_CE', 'ltp_CE', 'strike', 'ltp_PE', 'Δ_PE', 'Δ_OI_PE', 'volume_PE', 'iv_PE', 'oi_PE', 'PE_depth', 'tsq_PE', 'tbq_PE', 'atp_PE', 'vega_PE', 'gamma_PE', 'theta_PE', 'delta_PE', 'prev_close_PE', 'diff', f'{index}_spot', 'India_Vix']]
    
    option_data = df_option.copy()
    option_data["timestamp"] = pd.Timestamp.now()
    write_header = not os.path.exists(f'Credentials/Data/{tdate}/{index}_option_chain_{tdate}.csv')
    option_data.to_csv(f'Credentials/Data/{tdate}/{index}_option_chain_{tdate}.csv', mode='a', header=write_header, index=False)

    return df_option

wb = xw.Book('Analysis.xlsx')
sht_nifty = wb.sheets['nifty']
sht_bnf = wb.sheets['bnf']
sht_sensex = wb.sheets['sensex']
sht_summary = wb.sheets['summary']

while True:
    interval = datetime.now().time()
    minute = interval.minute
    sec = interval.second

    with lock:
        df = pd.DataFrame(data_base).T

    for col in df.columns:
        try:
            df[col] = pd.to_numeric(df[col])
        except (ValueError, TypeError):
            pass

    df = df[df.index.isin(sub_list+index_key)]

    df_nifty_option = option_chain(ce=index_ce_pe_list[0], pe=index_ce_pe_list[1], df=df, expiry=nifty_expiry[0], index='nifty')
    df_bnf_option = option_chain(ce=index_ce_pe_list[2], pe=index_ce_pe_list[3], df=df, expiry=bnf_expiry[0], index = 'bnf')
    df_sensex_option = option_chain(ce=index_ce_pe_list[4], pe=index_ce_pe_list[5], df=df, expiry=sensex_expiry[0], index = 'sensex')

    if (minute % 5 == 0) and (sec == 1):
        print('I entered here - 5min zone')
        new_synth_atm = synth_atm_index()
        if new_synth_atm != old_synth_atm:

            old_sub = sub_list
            sub_list_ce, sub_list_pe, sub_list, inst_strike_pair, index_ce_pe_list = valid_strikes(new_synth_atm)
            new_sub = sub_list

            unsubscribe_list = list(set(old_sub) - set(new_sub))
            subscribe_list = list(set(new_sub) - set(old_sub))


            if unsubscribe_list:
                streamer.unsubscribe(unsubscribe_list)
                print("Unsubscribed:", unsubscribe_list)
            
            if subscribe_list:
                streamer.subscribe(subscribe_list, "full")
                print("Subscribed:", subscribe_list)

            old_synth_atm = new_synth_atm.copy()
            print('I updated the subscription list')
        time.sleep(1)


    exit = sht_summary.range('C14').value
    if (exit == 'e') or (exit == 'E'):
        sht_summary.range('C14').value = None
        streamer.disconnect()
        wb.save()
        wb.close()
        break

    sht_nifty.range('A1').value = df_nifty_option
    sht_bnf.range('A1').value = df_bnf_option
    sht_sensex.range('A1').value = df_sensex_option

    time.sleep(0.5)
