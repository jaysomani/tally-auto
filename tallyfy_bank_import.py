
import pandas as pd
import requests
import logging


logging.basicConfig(filename='tally_import.log', level=logging.INFO, format='%(asctime)s - %(message)s')

TALLY_URL = "http://localhost:9000"


def check_ledger_exists(ledger_name):
    request_xml = """<ENVELOPE>
        <HEADER>
            <TALLYREQUEST>Export</TALLYREQUEST>
        </HEADER>
        <BODY>
            <EXPORTDATA>
                <REQUESTDESC>
                    <REPORTNAME>List of Ledgers</REPORTNAME>
                </REQUESTDESC>
            </EXPORTDATA>
        </BODY>
    </ENVELOPE>"""
    try:
        response = requests.post(TALLY_URL, data=request_xml, headers={'Content-Type': 'text/xml'})
        return ledger_name in response.text
    except Exception as e:
        logging.error(f"Error checking ledger '{ledger_name}': {e}")
        return False


def create_ledger(ledger_name):
    request_xml = f"""<ENVELOPE>
        <HEADER>
            <TALLYREQUEST>Import Data</TALLYREQUEST>
        </HEADER>
        <BODY>
            <IMPORTDATA>
                <REQUESTDESC>
                    <REPORTNAME>All Masters</REPORTNAME>
                </REQUESTDESC>
                <REQUESTDATA>
                    <TALLYMESSAGE>
                        <LEDGER NAME="{ledger_name}" ACTION="Create">
                            <NAME>{ledger_name}</NAME>
                            <PARENT>Sundry Creditors</PARENT>
                        </LEDGER>
                    </TALLYMESSAGE>
                </REQUESTDATA>
            </IMPORTDATA>
        </BODY>
    </ENVELOPE>"""
    try:
        response = requests.post(TALLY_URL, data=request_xml, headers={'Content-Type': 'text/xml'})
        return "<CREATED>1</CREATED>" in response.text
    except Exception as e:
        logging.error(f"Error creating ledger '{ledger_name}': {e}")
        return False


def post_transaction(date, narration, amount, txn_type, ledger):
    date_formatted = pd.to_datetime(date).strftime('%Y%m%d')
    request_xml = f"""<ENVELOPE>
        <HEADER>
            <TALLYREQUEST>Import Data</TALLYREQUEST>
        </HEADER>
        <BODY>
            <IMPORTDATA>
                <REQUESTDESC>
                    <REPORTNAME>Vouchers</REPORTNAME>
                </REQUESTDESC>
                <REQUESTDATA>
                    <TALLYMESSAGE>
                        <VOUCHER VCHTYPE="{txn_type}" ACTION="Create">
                            <DATE>{date_formatted}</DATE>
                            <NARRATION>{narration}</NARRATION>
                            <VOUCHERTYPENAME>{txn_type}</VOUCHERTYPENAME>
                            <ALLLEDGERENTRIES.LIST>
                                <LEDGERNAME>Cash</LEDGERNAME>
                                <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
                                <AMOUNT>-{amount}</AMOUNT>
                            </ALLLEDGERENTRIES.LIST>
                            <ALLLEDGERENTRIES.LIST>
                                <LEDGERNAME>{ledger}</LEDGERNAME>
                                <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
                                <AMOUNT>{amount}</AMOUNT>
                            </ALLLEDGERENTRIES.LIST>
                        </VOUCHER>
                    </TALLYMESSAGE>
                </REQUESTDATA>
            </IMPORTDATA>
        </BODY>
    </ENVELOPE>"""
    try:
        response = requests.post(TALLY_URL, data=request_xml, headers={'Content-Type': 'text/xml'})
        return "<CREATED>1</CREATED>" in response.text
    except Exception as e:
        logging.error(f"Error posting transaction: {e}")
        return False


def process_excel(file_path):
    df = pd.read_excel(file_path, engine='openpyxl')
    for _, row in df.iterrows():
        date = row['Date']
        narration = row['Narration']
        amount = row['Amount']
        txn_type = row['Receipt or Payment']
        ledger = row['Ledger']

        if not check_ledger_exists(ledger):
            print(f"Ledger '{ledger}' not found. Creating...")
            if create_ledger(ledger):
                print(f"Ledger '{ledger}' created successfully.")
            else:
                logging.error(f"Failed to create ledger: {ledger}")
                continue

        if post_transaction(date, narration, amount, txn_type, ledger):
            print(f"Transaction posted successfully: {narration} - {amount}")
        else:
            logging.error(f"Failed to post transaction: {narration}")

if __name__ == "__main__":
    process_excel("bank_statements.xlsx")
