# # # import pandas as pd
# # # import xmlrpc.client
# # # import logging

# # # # Configure logging
# # # logging.basicConfig(filename='tally_import.log', level=logging.INFO, format='%(asctime)s - %(message)s')

# # # # Tally XML-RPC Server Connection
# # # def connect_to_tally():
# # #     try:
# # #         return xmlrpc.client.ServerProxy("http://localhost:9000")
# # #     except Exception as e:
# # #         logging.error(f"Error connecting to Tally: {e}")
# # #         return None

# # # # Check if Ledger Exists in Tally
# # # def check_ledger_exists(tally, ledger_name):
# # #     try:
# # #         response = tally.GetLedger(ledger_name)
# # #         return response is not None
# # #     except Exception as e:
# # #         logging.error(f"Error checking ledger '{ledger_name}': {e}")
# # #         return False

# # # # Create Missing Ledger
# # # def create_ledger(tally, ledger_name):
# # #     try:
# # #         ledger_xml = f"""
# # #         <ENVELOPE>
# # #             <HEADER>
# # #                 <TALLYREQUEST>Import Data</TALLYREQUEST>
# # #             </HEADER>
# # #             <BODY>
# # #                 <IMPORTDATA>
# # #                     <REQUESTDESC>
# # #                         <REPORTNAME>All Masters</REPORTNAME>
# # #                     </REQUESTDESC>
# # #                     <REQUESTDATA>
# # #                         <TALLYMESSAGE xmlns:UDF="TallyUDF">
# # #                             <LEDGER NAME="{ledger_name}" ACTION="Create">
# # #                                 <NAME>{ledger_name}</NAME>
# # #                                 <PARENT>Sundry Creditors</PARENT> 
# # #                                 <ISBILLWISEON>Yes</ISBILLWISEON>
# # #                                 <AFFECTSSTOCK>No</AFFECTSSTOCK>
# # #                                 <GSTAPPLICABLE>&#4; Applicable</GSTAPPLICABLE>
# # #                                 <ISGSTAPPLICABLE>Yes</ISGSTAPPLICABLE>
# # #                             </LEDGER>
# # #                         </TALLYMESSAGE>
# # #                     </REQUESTDATA>
# # #                 </IMPORTDATA>
# # #             </BODY>
# # #         </ENVELOPE>
# # #         """

# # #         # Send XML Request to Tally
# # #         response = tally.ExecuteXML(ledger_xml)
# # #         if "<CREATED>1</CREATED>" in response or "<ALTERED>1</ALTERED>" in response:
# # #             logging.info(f"Ledger '{ledger_name}' created successfully.")
# # #             return True
# # #         else:
# # #             logging.error(f"Failed to create ledger '{ledger_name}'. Response: {response}")
# # #             return False
# # #     except Exception as e:
# # #         logging.error(f"Error creating ledger '{ledger_name}': {e}")
# # #         return False


# # # # Post Transaction to Tally
# # # def post_transaction(tally, date, narration, amount, txn_type, ledger):
# # #     try:
# # #         entry = {
# # #             "Date": date,
# # #             "Narration": narration,
# # #             "Amount": amount,
# # #             "Type": txn_type,
# # #             "Ledger": ledger,
# # #         }
# # #         response = tally.CreateVoucher(entry)
# # #         if response:
# # #             logging.info(f"Transaction posted successfully: {entry}")
# # #             return True
# # #     except Exception as e:
# # #         logging.error(f"Error posting transaction: {e}")
# # #     return False

# # # # Process Excel File
# # # def process_excel(file_path):
# # #     df = pd.read_excel(file_path, engine='openpyxl')
# # #     tally = connect_to_tally()
# # #     if not tally:
# # #         print("Could not connect to Tally. Exiting.")
# # #         return

# # #     for _, row in df.iterrows():
# # #         date = row['Date']
# # #         narration = row['Narration']
# # #         amount = row['Amount']
# # #         txn_type = row['Receipt or Payment']
# # #         ledger = row['Ledger']

# # #         if check_ledger_exists(tally, ledger):
# # #             post_transaction(tally, date, narration, amount, txn_type, ledger)
# # #         else:
# # #             print(f"Ledger '{ledger}' not found. Creating...")
# # #             if create_ledger(tally, ledger):
# # #                 print(f"Ledger '{ledger}' created successfully. Posting transaction...")
# # #                 post_transaction(tally, date, narration, amount, txn_type, ledger)
# # #             else:
# # #                 logging.error(f"Failed to create ledger: {ledger}")

# # # if __name__ == "__main__":
# # #     process_excel("bank_statement.xlsx")
# # import requests

# # TALLY_URL = "http://localhost:9000"

# # xml_request = """<ENVELOPE>
# #     <HEADER>
# #         <VERSION>1</VERSION>
# #         <TALLYREQUEST>Export</TALLYREQUEST>
# #         <TYPE>Collection</TYPE>
# #         <ID>List of Companies</ID>
# #     </HEADER>
# #     <BODY>
# #         <DESC>
# #             <STATICVARIABLES>
# #                 <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
# #             </STATICVARIABLES>
# #         </DESC>
# #     </BODY>
# # </ENVELOPE>"""

# # headers = {'Content-Type': 'text/xml'}

# # try:
# #     response = requests.post(TALLY_URL, data=xml_request, headers=headers)
# #     if response.status_code == 200:
# #         print("✅ Connected to Tally! Response:")
# #         print(response.text)
# #     else:
# #         print(f"⚠️ Failed to connect, HTTP Status Code: {response.status_code}")
# # except Exception as e:
# #     print("❌ Failed to connect to Tally:", e)

# import requests

# TALLY_URL = "http://localhost:9000"

# ledger_xml = """<ENVELOPE>
#     <HEADER>
#         <TALLYREQUEST>Import Data</TALLYREQUEST>
#     </HEADER>
#     <BODY>
#         <IMPORTDATA>
#             <REQUESTDESC>
#                 <REPORTNAME>All Masters</REPORTNAME>
#             </REQUESTDESC>
#             <REQUESTDATA>
#                 <TALLYMESSAGE xmlns:UDF="TallyUDF">
#                     <LEDGER NAME="Test Ledger" ACTION="Create">
#                         <NAME>Test Ledger</NAME>
#                         <PARENT>Sundry Creditors</PARENT>
#                         <OPENINGBALANCE>0.00</OPENINGBALANCE>
#                         <ISBILLWISEON>Yes</ISBILLWISEON>
#                         <AFFECTSSTOCK>No</AFFECTSSTOCK>
#                     </LEDGER>
#                 </TALLYMESSAGE>
#             </REQUESTDATA>
#         </IMPORTDATA>
#     </BODY>
# </ENVELOPE>"""

# headers = {'Content-Type': 'text/xml'}

# try:
#     response = requests.post(TALLY_URL, data=ledger_xml, headers=headers)
#     print("Tally Response:", response.text)
# except Exception as e:
#     print("Error:", e)
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
