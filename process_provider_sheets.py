import pandas as pd 
import os

# Ensures columns in output excel sheets are not truncated
float_format = '{:.2f}'.format
pd.set_option('display.max_colwidth', None)

# Define the current directory
current_directory = os.getcwd()
excel_files = [file for file in os.listdir(current_directory) if file.endswith('.xlsx') and not file.startswith('~$')]

# Define and create the output directory if it doesn't exist
output_directory = 'Output'
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

#Deletes trailing/leading whitespace for BCBS claims
def clean_column_name(column_name):
    column_name = column_name.strip()
    column_name = column_name.replace('\n', '')
    return column_name

#Maps old excel output to desired convention based on Provider
def process_provider_data(file_name):
    if 'DOE' in file_name:
        provider_name = 'PRACTICE'
        sheet_name = "MEDICAL"
        header = 0
    elif 'HEALTHGRAM' in file_name:
        provider_name = 'HEALTHGRAM'
        sheet_name = "SF_Claims Filing (Excel)"
        header = 1
    elif 'BCBS' in file_name:
        provider_name = 'BCBS'
        sheet_name = 'Sheet1'
        header = 9
    else:
        print(f"Ignoring invalid Excel file: {file_name}")
        return None, None, None


    df = pd.read_excel(file_name, sheet_name, header = header)
    if provider_name == 'PRACTICE':
        df ['EmployeeCode'] = 'E'
        df = df.rename(columns = {
            'MEMBER ID': 'SSN', 
            'FIRST NAME': 'ClaimantFirstName', 
            'LAST NAME': 'ClaimantLastName', 
            'CLAIM NUMBER': 'ClaimNo',
            'Dx Code': 'ICD9', 
            'Procedure': 'CPTCode', 
            'CHECK NUMBER': 'CheckNo', 
            'BEGINNING SERVICE': 'ServiceDate', 
            'END SERVICE': 'ClaimReceiptDate', 
            'DATE PAID': 'PaidDate', 
            'Provider Name': 'Provider', 
            'BILLED': 'BilledAmount', 
            'PAID': 'PaidAmount',
            'PROVIDER TIN': 'ProviderNo'
        })
    elif provider_name == "HEALTHGRAM":
        df['EmployeeCode'] = 'e'
        df = df.rename(columns = {
            'membno': 'SSN', 
            'mfstnam': 'ClaimantFirstName', 
            'mlstnam': 'ClaimantLastName', 
            'claimno': 'ClaimNo',
            'diagn': 'ICD9', 
            'svccod': 'CPTCode', 
            'chknum': 'CheckNo', 
            'svcdat': 'ServiceDate', 
            'enddat': 'ClaimReceiptDate', 
            'pidate': 'PaidDate', 
            'plstnam': 'Provider', 
            'claamt': 'BilledAmount', 
            'to pay': 'PaidAmount',
            'provno': 'ProviderNo'
        })
    elif provider_name == 'BCBS':
        df.columns = [clean_column_name(col) for col in df.columns]
        first_name_df = pd.read_excel(file_name, sheet_name, header = None, skiprows = 3, nrows = 1)
        first_name_series = first_name_df.iloc[0]
        full_name = first_name_series.values[0]
        first_name, last_name = full_name.split()
        df['ClaimantFirstName'] = first_name
        df['ClaimantLastName'] = last_name
        df['ProviderNo'] = '22-2222222'
        df['EmployeeCode'] = 'E'
        df['CopyOfFirstDOS'] = df['FIRSTDOS']
        df = df.rename(columns = {
            'SUBSCRIBER#': 'SSN', 
            'CLAIM#': 'ClaimNo',
            'DIAGCODE': 'ICD9', 
            'PROCEDURECODE': 'CPTCode', 
            'CHECK#': 'CheckNo', 
            'CopyOfFirstDOS': 'ServiceDate', 
            'FIRSTDOS': 'ClaimReceiptDate', 
            'PAIDDATE': 'PaidDate', 
            'PROVIDER NAME / TYPE': 'Provider', 
            'TOTALCHARGES': 'BilledAmount', 
            'TOTALPAYMENT': 'PaidAmount',
        })
    df = df[['SSN', 'ClaimantFirstName', 'ClaimantLastName', 'EmployeeCode', 
            'ProviderNo', 'ClaimNo', 'Provider', 'ServiceDate', 'ClaimReceiptDate', 'PaidDate',
            'ICD9', 'BilledAmount', 'PaidAmount', 'CPTCode', 'CheckNo']]
    return df, provider_name, sheet_name

# Call process_provider_data for each file in the current directory
def output_excel_sheet(): 
    for file_name in excel_files:
        print(file_name)
        provider_data, provider_name, sheet_name = process_provider_data(file_name)
        if (sheet_name == None):
            continue
        
        # Set numeric format for integer columns
        numeric_format = '${:,.2f}'
        numeric_columns = {
            'ClaimNo',
            'CheckNo',
            'ProviderNo',
        }

        # Get provider data from excel sheet
        final_df = pd.DataFrame(provider_data)
        last_name = final_df['ClaimantLastName'].iloc[0]
        first_name = final_df['ClaimantFirstName'].iloc[0]
        final_df.dropna(subset=['ClaimNo'], inplace=True)

        # Set output file name & path
        output_file_name = f"{provider_name}_{first_name}_{last_name}.xlsx"
        output_file_path = os.path.join(output_directory, output_file_name)

        writer = pd.ExcelWriter(output_file_path)
        final_df.to_excel(writer, sheet_name = sheet_name, index = False, na_rep = 'NAN')

        for column in final_df.columns:
            column_length = max(final_df[column].astype(str).map(len).max(), len(column)) + 3
            col_index = final_df.columns.get_loc(column)
            writer.sheets[sheet_name].set_column(col_index, col_index, column_length)
        writer.close()

#Run output_excel_sheet to process all the provider data in the current directory
output_excel_sheet()
print("Process completed. Output files saved in the 'Output' directory.")
