import xml.etree.ElementTree as ET
import openpyxl
import datetime
# Parse the XML file
tree = ET.parse('Input.xml')
root = tree.getroot()

# Create a new spreadsheet
workbook = openpyxl.Workbook()
sheet = workbook.active

# Add a header row to the spreadsheet
sheet.append(['Date', 'Transaction Type', 'Vch No.', 'Ref No', 'Ref Date','Ref type', 'Debtor', 'Ref Amount', 'Amount', 'Particulars', 'Vch Type', 'Amount Verified'])

# Iterate through the transactions in the XML file
for voucher in root.findall('.//VOUCHER'):
    # Extract the data for the transaction
    try:
        date = voucher.find('DATE').text
        formatted_date = datetime.datetime.strptime(date, '%Y%m%d').strftime('%Y-%m-%d')
    except AttributeError:
        formatted_date = None
    try:
        voucher_type = voucher.find('VOUCHERTYPENAME').text
    except AttributeError:
        voucher_type = None
    try:
        voucher_number = voucher.find('VOUCHERNUMBER').text
    except AttributeError:
        voucher_number = None

    try:
        reference_element = voucher.find('VOUCHERNUMBER')
        if reference_element is not None:
            reference_no = reference_element.text
        else:
            reference_no = None
    except AttributeError:
        reference_no = None

    try:
        reference_date = voucher.find('REFERENCEDATE').text
        Formatted_reference_date = datetime.datetime.strptime(reference_date, '%Y%m%d').strftime('%Y-%m-%d')
    except AttributeError:
        Formatted_reference_date = None
    try:
        debtor_element = voucher.find('PARTYLEDGERNAME')
        if debtor_element is not None:
            debtor_name = debtor_element.text
        else:
            debtor_name = None
    except AttributeError:
        debtor_name = None

    try:
        reference_amount_element = voucher.find('REFERENCEAMOUNT')
        if reference_amount_element is not None:
            reference_amount = reference_amount_element.text
        else:
            reference_amount = None
    except AttributeError:
        reference_amount = None

    try:
        reference_TYPE = voucher.find('BILLTYPE')
        if reference_TYPE is not None:
            ref_TYPE = reference_TYPE.text
        else:
            ref_TYPE = None
    except AttributeError:
        ref_TYPE = None


    narration = voucher.find('PARTYLEDGERNAME')
    if narration is not None:
        particulars = narration.text
    else:
        particulars = None

    amount = voucher.find('AMOUNT')
    if amount is not None:
        amount_text = amount.text
    else:
        amount_text = None

    # Determine the transaction type and amount verified based on the voucher type
    if voucher_type == 'Receipt':
        transaction_type = 'Parent'
        amount_verified = 'Yes'
    elif voucher_type == 'Bank/GST/etc':
        transaction_type = 'Other'
        amount_verified = 'NA'
    else:
        transaction_type = 'Child'
        amount_verified = 'NA'
        # Only append rows with Vch Type equal to Receipt

    sheet.append(
        [formatted_date, transaction_type, voucher_number, reference_no, Formatted_reference_date, ref_TYPE, debtor_name,
         reference_amount, amount_text, particulars, voucher_type,
         amount_verified])

    # Save the spreadsheet to a file
    workbook.save('response1.xlsx')