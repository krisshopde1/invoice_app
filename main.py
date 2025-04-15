import os
import pdfplumber
from pypdf import PdfReader
import re
import pandas as pd
import streamlit as st
from datetime import datetime
import fitz  # PyMuPDF
import pytesseract
from pdf2image import convert_from_path
import io

# Find the installed Tesseract path in your Conda environment
#tesseract_path = r"C:\Users\yuxun_tee\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
# Set tesseract path
#pytesseract.pytesseract.tesseract_cmd = tesseract_path
# Set poppler path
# poppler_path = r"C:\Users\yuxun_tee\Downloads\Release-24.08.0-0\poppler-24.08.0\Library\bin"

folder_list = [
    "Apple", "BAN LEONG TECHNOLOGIES LTD", "CONVERGENT SYSTEMS",
    "CRYSTAL WINES PTE LTD", "DFASS (SINGAPORE) PTE LTD",
    "DHL EXPRESS (SINGAPORE) PTE LTD", "DIGITAL HUB PTE LTD",
    "iFactory Asia Pte Ltd", "KRIS+ PTE. LTD", "PIVENE PTE LTD",
    "SETELCO COMMUNICATIONS","output"
]
def create_folder_structure(folder_names):
    """
    Creates a folder structure with given folder names inside the directory of the script.
    
    Parameters:
    folder_names (list): A list of folder names to create.
    """
    # Get the directory where the script is located
    root_path = os.path.dirname(os.path.realpath(__file__))
    
    if not os.path.exists(root_path):
        os.makedirs(root_path)  # Create root path if it doesn't exist
    
    for folder in folder_names:
        folder_path = os.path.join(root_path, folder)
        os.makedirs(folder_path, exist_ok=True)

# Function to retrieve string values based on index
def get_value(data, search_text, offset):
    for i, line in enumerate(data):
        if search_text == line:
            return data[i + offset].strip()

def construct_folder_path(*subfolders):
    BASE_FOLDER_PATH = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(BASE_FOLDER_PATH, *subfolders)

# The following functions extracts and processes invoices for each vendor. 
# Each vendor has 2 functions (extract then process)
# Add functions to process new vendors under this section
def extract_invoice_data_crystalwines(file_path):
    
    pages = convert_from_path(file_path,450) # convert to image
    text = pytesseract.image_to_string(pages[0]) # print first page
    
    lines = text.splitlines()
    lines = [x for x in lines if x.strip()]

    if len(re.findall(r'TAX INVOICE', text))>0:
                
        title = 'TAX INVOICE'
                
        try:
            vendor = get_value(lines, "Crystal Wines Pte Ltd", 0)
        except:
            vendor = 'Not Found'

        try:
            vendor_address = get_value(lines, "Crystal Wines Pte Ltd", 1)
        except:
            vendor_address = 'Not Found'

        try:
            invoice_number =re.findall(r'Invoice Number\s+(CW/\d+)', text)[0].strip()
        except:
            invoice_number = 'Not Found'
        
        try:
            invoice_date = get_value(lines, 'Invoice Date', 1)
        except:
            invoice_date = 'Not Found'
        
        try:
            gst_reg_no = re.findall(r'(?<=\nGST Registration No. : ).+', text)[0].strip()
        except:
            gst_reg_no = 'Not Found'
        
        try:
            sold_to = get_value(lines, 'Bill To:', 1)
        except:
            sold_to = 'Not Found'

        try:
            sold_to_address = get_value(lines, 'Bill To:', 2) + ', ' + get_value(lines, 'Bill To:', 3)
        except:
            sold_to_address = 'Not Found'
        
        try:
            currency = re.findall(r'(?<=1. All payment shall be in )\w+\s+\w+', text)[0]
        except:
            currency = 'Not Found'
        
        try:
            gst_amount = re.findall(r'(?<=GST \d% ).+', text)[0].strip()
        except:
            gst_amount = 'Not Found'

        try:
            total_amount = re.findall(r'(?<=Singapore Dollars unless otherwise indicated. SubTotal ).+', text)[0].strip()
        except:
            total_amount = 'Not Found'

        extracted_data = {
                    "Title": title,
                    "Vendor": vendor,
                    "Vendor Address": vendor_address,
                    "GST Reg. No.": gst_reg_no,
                    "Sold To": sold_to,
                    "Sold To Address": sold_to_address,
                    "Invoice Number": invoice_number,
                    "Invoice Date": invoice_date,
                    "Currency": currency,
                    "GST Amount": gst_amount,
                    "Total Amount": total_amount,
                    #"PDF File": file_path,
                }
        return extracted_data
    
    else:

        extracted_data = {
                    "Title": 'Not Found',
                    "Vendor": 'Not Found',
                    "Vendor Address": 'Not Found',
                    "GST Reg. No.": 'Not Found',
                    "Sold To": 'Not Found',
                    "Sold To Address": 'Not Found',
                    "Invoice Number": 'Not Found',
                    "Invoice Date": 'Not Found',
                    "Currency": 'Not Found',
                    "GST Amount": 'Not Found',
                    "Total Amount": 'Not Found',
                    #"PDF File": file_path,
                }
        return extracted_data

def process_crystalwines_and_export():
    folder_path = './CRYSTALWINES'
    out_folder = './output'

    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    if not pdf_files:
        # st.error("No PDF files found in the current directory.")
        return

    data = []
    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        invoice_data = extract_invoice_data_crystalwines(pdf_path)
        if invoice_data:
            invoice_data["PDF File"] = pdf_file
            data.append(invoice_data)

    df = pd.DataFrame(data)
    if df.empty:
        # st.error("No data extracted from the PDFs.")
        pass
    else:
        output_csv = os.getcwd()+r'\output\crystalwines_invoice_data.xlsx'
        df.to_csv(output_csv, index=False)
        with pd.ExcelWriter(output_csv) as writer:
            df.to_excel(writer, index=False)

def extract_invoice_data_setelco(pdf_path):
    try:
        
        pages = convert_from_path(pdf_path,450) # convert to image
        #tessdata_dir_config = r'--tessdata-dir ".\tessdata-main"'
        text = pytesseract.image_to_string(pages[0]) # print first page
        
        lines = text.splitlines()
        lines = [x for x in lines if x.strip()]

        if len(re.findall(r'TAX INVOICE', text))>0:
            
            title = 'TAX INVOICE'

            vendor = "Not Found"
            for line in lines:
                if line.startswith("to "):
                    vendor = line[3:]
                    break

            vendor_address = "Not Found"
            vendor_address_1 = get_value(lines, 'SETELCO COMMUNICATIONS PTE LTD', 1)
            vendor_address_2 = get_value(lines, 'SETELCO COMMUNICATIONS PTE LTD', 2)
            vendor_address = vendor_address_1 + ", "+ vendor_address_2 

            invoice_number = "Not Found"
            invoice_number = re.findall(r'\d+', get_value(lines, 'TAX INVOICE', 1))[0]

            invoice_date = "Not Found"
            for line in lines:
                if line.startswith("Invoice Date : "):
                    invoice_date = line[len("Invoice Date : "):].strip()
                    break

            gst_reg_no = "Not Found"
            for line in lines:
                if line.startswith("GST Regn: "):
                    gst_reg_no = line[len("GST Regn: "):].strip()
                    break

            sold_to = "Not Found"
            sold_to = get_value(lines, "Bill To:", 1)

            sold_to_address = "Not Found"
            sold_to_address_1 = get_value(lines, "Bill To:", 2)
            sold_to_address_2 = get_value(lines, "Bill To:", 3)
            sold_to_address = sold_to_address_1+', '+sold_to_address_2            
            
            currency = "Not Found"
            for line in lines:
                if line.startswith("Total Amount "):
                    currency = line[len("Total Amount "):-2]

            gst_amount = "Not Found"
            for line in lines:
                if line.startswith("GST @9.00% : "):
                    gst_amount = line[len("GST @9.00% : "):].strip()
    
            total_amount = "Not Found"
            for line in lines:
                if line.startswith("Total Amount : "):
                    total_amount = line[len("Total Amount : "):].strip()

            extracted_data = {
                "Title": title,
                "Vendor": vendor,
                "Vendor Address": vendor_address,
                "GST Reg. No.": gst_reg_no,
                "Sold To": sold_to,
                "Sold To Address": sold_to_address,
                "Invoice Number": invoice_number,
                "Invoice Date": invoice_date,
                "Currency": currency,
                "GST Amount": gst_amount,
                "Total Amount": total_amount,
                #"PDF File": pdf_path,
            }
    
        else:
            not_tax_invoice = 'Not Tax Invoice'
            extracted_data = {
                "Title": not_tax_invoice,
                "Vendor": not_tax_invoice,
                "Vendor Address": not_tax_invoice,
                "GST Reg. No.": not_tax_invoice,
                "Sold To": not_tax_invoice,
                "Sold To Address": not_tax_invoice,
                "Invoice Number": not_tax_invoice,
                "Invoice Date": not_tax_invoice,
                "Currency": not_tax_invoice,
                "GST Amount": not_tax_invoice,
                "Total Amount": not_tax_invoice,
                #"PDF File": pdf_path,
            }

        return extracted_data
            
    except Exception as e:
        return {"Error": str(e)}
            
def process_setelco_and_export():
    folder_path = './SETELCO COMMUNICATIONS'
    out_folder = './output'

    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    if not pdf_files:
        # st.error("No PDF files found in the current directory.")
        return

    data = []
    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        invoice_data = extract_invoice_data_setelco(pdf_path)
        if invoice_data:
            invoice_data["PDF File"] = pdf_file
            data.append(invoice_data)

    df = pd.DataFrame(data)
    if df.empty:
        # st.error("No data extracted from the PDFs.")
        pass
    else:
        output_csv = os.getcwd()+r'\output\setelco_invoice_data.xlsx'
        df.to_csv(output_csv, index=False)
        with pd.ExcelWriter(output_csv) as writer:
            df.to_excel(writer, index=False)

def extract_invoice_data_dhl(pdf_path):
    try:
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)
            page = reader.pages[0]
            text = page.extract_text()

        lines = text.splitlines()

        title = "Not found"
        for i, line in enumerate(lines):
            if "Type of Service" in line:
                title = lines[i - 1].strip()
                break

        vendor = "Not found"
        for line in lines:
            if "PLEASE SEND YOUR REMITTANCES TO" in line:
                vendor = line.split("PLEASE SEND YOUR REMITTANCES TO")[-1].strip()
                break

        vendor_address = "Not found"
        for i, line in enumerate(lines):
            if "PLEASE SEND YOUR REMITTANCES TO" in line and i + 1 < len(lines):
                vendor_address = lines[i + 1].strip()
                address_words = vendor_address.split(', ')
                if len(address_words) > 3:
                    address_words = address_words[:-1]
                vendor_address = ', '.join(address_words)

        gst_reg_no = "Not found"
        for line in lines:
            if "GST REG NO.:" in line:
                gst_reg_no = line.split('GST REG NO.:')[-1].strip() if gst_reg_no else "Not found"
                break

        sold_to = "Not found"
        for i, line in enumerate(lines):
            if "Billing Chat" in line and i + 2 < len(lines):
                sold_to = lines[i + 2].strip()  # Second line after "Billing Chat"
                break

        sold_to_address = "Not found"
        for i, line in enumerate(lines):
            if sold_to in line and i + 2 < len(lines):
                sold_to_address = lines[i + 2].strip()  # The line immediately after the sold-to line
                break

        invoice_number = "Not found"
        for i, line in enumerate(lines):
            if "Invoice Number:" in line and i + 1 < len(lines):
                invoice_number = lines[i + 1].strip()  # Line after "Invoice No:"
                break

        invoice_date = "Not found"
        for i, line in enumerate(lines):
            if "Invoice Date:" in line and i + 1 < len(lines):
                invoice_date = lines[i + 1].strip()  # Line after "Invoice No:"
                break

        gst_amount = "Not found"
        for i, line in enumerate(lines):
            if "Total Amount" in line and i + 2 < len(lines):
                gst_amount = lines[i + 2].strip()  # Fourth line after "GST only"
                break

        total_amount = "Not found"
        for i, line in enumerate(lines):
            if "Total Amount" in line and i + 3 < len(lines):
                total_amount = lines[i + 3].strip()  # Fourth line after "GST only"
                break

        currency = "Not found"
        for line in lines:
            if "Total Amount" in line:
                currency = line.split('Total Amount')[-1].strip() if gst_reg_no else "Not found"
                break
        
        return {
            "Title": title,
            "Vendor": vendor,
            "Vendor Address": vendor_address,
            "GST Reg. No.": gst_reg_no,
            "Sold To": sold_to,
            "Sold To Address": sold_to_address,
            "Invoice Number": invoice_number,
            "Invoice Date": invoice_date,
            "Currency": currency,
            "GST Amount": gst_amount,
            "Total Amount": total_amount,
        }

    except Exception as e:
        return {"Error": str(e)}

def process_dhl_and_export():
    # folder_path = construct_folder_path("DHL EXPRESS (SINGAPORE) PTE LTD")
    folder_path = './DHL EXPRESS (SINGAPORE) PTE LTD'
    # out_folder = construct_folder_path("output")
    # os.makedirs(out_folder, exist_ok=True)

    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    if not pdf_files:
        st.error("No PDF files found in the current directory.")
        return

    data = []
    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        invoice_data = extract_invoice_data_dhl(pdf_path)
        if invoice_data:
            invoice_data["PDF File"] = pdf_file
            data.append(invoice_data)

    df = pd.DataFrame(data)
    if df.empty:
        st.error("No data extracted from the PDFs.")
    else:
        output_csv = os.getcwd()+r'\output\dhl_invoice_data.xlsx'
        # df.to_csv(output_csv, index=False)
        with pd.ExcelWriter(output_csv) as writer:
            df.to_excel(writer, index=False)
        st.success(f"DHL data successfully extracted! CSV file saved at {output_csv}")

def extract_invoice_data_dfass(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            text = page.extract_text()

        is_credit_note = "credit note" in text.lower()
        if is_credit_note:
            title = "Credit Note"
            invoice_number = re.search(r'Credit note (\S+)', text)
            invoice_number = invoice_number.group(1) if invoice_number else "Not found"
            date_time = re.search(r'Date and time (\d{1,2}/\d{1,2}/\d{4} \d{1,2}:\d{2} (AM|PM))', text)
            date_time = date_time.group(1) if date_time else "Not found"
            sales_tax = re.search(r'Sales tax\s?\(([\d,\.]+)\)\s?SGD', text)
            sales_tax = sales_tax.group(1) if sales_tax else "Not found"
            total = re.search(r'Total\s?\(([\d,\.]+)\)\s?SGD', text)
            total = total.group(1) if total else "Not found"
            vendor = text.splitlines()[0].strip() if text.splitlines() else "Not found"
            vendor_address = text.splitlines()[1].strip() if len(text.splitlines()) > 1 else "Not found"
            gst_reg_no = re.search(r'GST registration number (\S+)', text)
            gst_reg_no = gst_reg_no.group(1) if gst_reg_no else "Not found"
            
            bill_to_index = text.find("Bill to:")
            currency = re.search(r'\(\d+[\d,\.]*\)\s+([A-Za-z]{3})', text)
            currency = currency.group(1) if currency else "Not found"
            if bill_to_index != -1:
                lines_after_bill_to = text[bill_to_index:].splitlines()
                sold_to = lines_after_bill_to[1].strip() if len(lines_after_bill_to) > 1 else "Not found"

            if bill_to_index != -1:
                lines_after_bill_to = text[bill_to_index:].splitlines()
                sold_to_address = lines_after_bill_to[2].strip() if len(lines_after_bill_to) > 2 else "Not found"



        else:
            title="Tax Invoice"
            invoice_number = re.search(r'Tax Invoice (\S+)', text)
            invoice_number = invoice_number.group(1) if invoice_number else "Not found"
            date_time = re.search(r'Date and time (\d{1,2}/\d{1,2}/\d{4} \d{1,2}:\d{2} (AM|PM))', text)
            date_time = date_time.group(1) if date_time else "Not found"
            sales_tax = re.search(r'Sales tax\s+([\d,\.]+)', text)
            sales_tax = sales_tax.group(1) if sales_tax else "Not found"
            total = re.search(r'Total\s+([\d,\.]+)\s+SGD', text)
            total = total.group(1) if total else "Not found"
            vendor = text.splitlines()[0].strip() if text.splitlines() else "Not found"
            gst_reg_no = re.search(r'GST registration number (\S+)', text)
            gst_reg_no = gst_reg_no.group(1) if gst_reg_no else "Not found"
            currency = re.search(r'Sales tax\s+([\d,\.]+)\s+([A-Za-z]{3})', text)
            currency = currency.group(2) if currency else "Not found"
            
            bill_to_index = text.find("Bill to:")
            if bill_to_index != -1:
                lines_after_bill_to = text[bill_to_index:].splitlines()
                sold_to = lines_after_bill_to[1].strip() if len(lines_after_bill_to) > 1 else "Not found"

            if bill_to_index != -1:
                lines_after_bill_to = text[bill_to_index:].splitlines()
                sold_to_address = lines_after_bill_to[2].strip() if len(lines_after_bill_to) > 2 else "Not found"

            currency = re.search(r'Sales tax\s+([\d,\.]+)\s+([A-Za-z]{3})', text)
            currency = currency.group(2) if currency else "Not found"

        return {
            "Title": title,
            "Vendor": vendor,
            "Vendor Address": "nil",
            "GST Reg. No.": gst_reg_no,
            "Sold To": sold_to,
            "Sold To Address": sold_to_address,
            "Invoice Number": invoice_number,
            "Invoice Date": date_time,
            "Currency": currency,
            "GST Amount": sales_tax,
            "Total Amount": total,

        }
    except Exception as e:
        return {"Error": str(e)}

def process_dfass_and_export():
    # folder_path = construct_folder_path("DFASS (SINGAPORE) PTE LTD")
    folder_path = './DFASS (SINGAPORE) PTE LTD'
    
    data = []
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(folder_path, file_name)
            extracted_data = extract_invoice_data_dfass(file_path)
            extracted_data["File Name"] = file_name
            data.append(extracted_data)

    dfass_data = pd.DataFrame(data)
    
    if dfass_data.empty:
        st.error("No DFASS data extracted.")
    else:
        # out_folder = construct_folder_path("output")
        # output_csv = os.path.join(out_folder, 'dfass_invoice_data.csv')
        output_csv = os.getcwd()+r'\output\dfass_invoice_data.xlsx'
        # dfass_data.to_csv(output_csv, index=False)  # Save as separate CSV
        with pd.ExcelWriter(output_csv) as writer:
            dfass_data.to_excel(writer, index=False)
        st.success(f"DFASS data successfully extracted! CSV file saved at {output_csv}")

def extract_invoice_data_kris(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            text = page.extract_text()
            
        title = text.splitlines()[0] if text.splitlines() else "Not found"
        lines = text.splitlines()

        # Extract vendor (third line)
        vendor = ""
        for i, line in enumerate(lines):
            if "Company" in line:
                if i + 1 < len(lines):
                    vendor = lines[i + 1]
                break

        # Process the third line (if it exists) and remove the last 5 words
        if len(lines) > 2:  # Ensure there are at least 3 lines
            third_line = lines[2]
            words = third_line.split()
            words = words[:-5]  # Remove the last 5 words
            vendor = ' '.join(words)  # Reassemble the line
                
        vendor_address = lines[6] if len(lines) > 6 else "Not found"

        sold_to = ""
        for i, line in enumerate(lines):
            if line.strip() == "Company":
                if i + 1 < len(lines):
                    sold_to = lines[i + 1]
                break
        
        sold_to_address = ""
        for i, line in enumerate(lines):
            if line.strip() == "Company":
                if i + 3 < len(lines):
                    sold_to_address = lines[i + 3]
                break

        currency_match = re.search(r'\((.*?)\)', text)
        currency = currency_match.group(1) if currency_match else "Not found"
        
        gst_reg_no = re.search(r'GST Reg No : (\S+)', text)
        gst_reg_no = gst_reg_no.group(1) if gst_reg_no else "Not found"

        tax_invoice_number = re.search(r'Document No : (\S+)', text)
        tax_invoice_number = tax_invoice_number.group(1) if tax_invoice_number else "Not found"

        tax_invoice_date = re.search(r'Date : (\S+)', text)
        tax_invoice_date = tax_invoice_date.group(1) if tax_invoice_date else "Not found"

        gst_amount = re.search(r'GST - NS (\S+)', text)
        gst_amount = gst_amount.group(1) if gst_amount else "Not found"

        # Extract Total
        total_amount = re.search(r'Total (\S+)', text)
        total_amount = total_amount.group(1) if total_amount else "Not found"

        # Return the extracted data
        return {
            "Title": title,
            "Vendor": vendor,
            "Vendor Address": vendor_address,
            "GST Reg. No.": gst_reg_no,
            "Sold To": sold_to,
            "Sold To Address": sold_to_address,
            "Invoice Number": tax_invoice_number,
            "Invoice Date": tax_invoice_date,
            "Currency": currency,
            "GST Amount": gst_amount,
            "Total Amount": total_amount,

        }

    except Exception as e:
        return None
        
def process_kris_and_export():
    # folder_path = construct_folder_path("KRIS+ PTE. LTD")
    folder_path = './KRIS+ PTE. LTD'
    
    data = []
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(folder_path, file_name)
            extracted_data = extract_invoice_data_kris(file_path)
            extracted_data["File Name"] = file_name
            data.append(extracted_data)

    kris_data = pd.DataFrame(data)
    
    if kris_data.empty:
        st.error("No kris data extracted.")
    else:
        # out_folder = construct_folder_path("output")
        # output_csv = os.path.join(out_folder, 'kris_invoice_data.csv')
        output_csv = os.getcwd()+r'\output\kris_invoice_data.xlsx'
        # kris_data.to_csv(output_csv, index=False)  # Save as separate CSV
        with pd.ExcelWriter(output_csv) as writer:
            kris_data.to_excel(writer, index=False)
        st.success(f"kris data successfully extracted! CSV file saved at {output_csv}")

def extract_invoice_data_apple(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            text = page.extract_text()
        
        # Common Fields
        title = text.splitlines()[0].strip() if text.strip() else "Not found"
        vendor = text.splitlines()[1].strip() if len(text.splitlines()) > 1 else "Not found"
        vendor_address = text.splitlines()[2].strip() if len(text.splitlines()) > 2 else "Not found"
        gst_reg_no = re.search(r'GST Reg\. No\. (\S+)', text)
        gst_reg_no = gst_reg_no.group(1) if gst_reg_no else "Not found"        
        sold_to = re.search(r'Apple Order Number: [A-Z0-9]+\s+(.+)', text)
        sold_to = sold_to.group(1).strip() if sold_to else "Not found"
        sold_to_address = re.search(r'(Credit Note Date|Tax Invoice Date): \d{2}/\d{2}/\d{4}\s+(.+)', text)
        sold_to_address = sold_to_address.group(2).strip() if sold_to_address else "Not found"
        tax_invoice_number = re.search(r'(Credit Note|Tax Invoice) Number: (\S+)', text)
        tax_invoice_number = tax_invoice_number.group(2) if tax_invoice_number else "Not found"
        tax_invoice_date = re.search(r'(Credit Note|Tax Invoice) Date: (\S+)', text)
        tax_invoice_date = tax_invoice_date.group(2) if tax_invoice_date else "Not found"
        currency = re.search(r'Total Value \(Incl\.GST\) (\S+)', text)
        currency = currency.group(1) if currency else "Not found"
        gst_amount = re.search(r'Terms and Conditions.*?GST.*?(\d+\.\d+)', text, re.DOTALL)
        gst_amount = gst_amount.group(1).strip() if gst_amount else "Not found"
        total_amount = re.search(r'Total Value \(Incl\.GST\).*?([\d.,]+)$', text, re.MULTILINE)
        total_amount = total_amount.group(1) if total_amount else "Not found"

        return {
            "Title": title,
            "Vendor": vendor,
            "Vendor Address": vendor_address,
            "GST Reg. No.": gst_reg_no,
            "Sold To": sold_to,
            "Sold To Address": sold_to_address,
            "Invoice Number": tax_invoice_number,
            "Invoice Date": tax_invoice_date,
            "Currency": currency,
            "GST Amount": gst_amount,
            "Total Amount": total_amount,
        }
    except Exception as e:
        return {"Error": str(e)}

def process_apple_and_export():
    # folder_path = construct_folder_path("Apple")
    folder_path = './Apple'
    
    data = []
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(folder_path, file_name)
            extracted_data = extract_invoice_data_apple(file_path)
            extracted_data["File Name"] = file_name
            data.append(extracted_data)

    apple_data = pd.DataFrame(data)
    
    if apple_data.empty:
        st.error("No Apple data extracted.")
    else:
        # out_folder = construct_folder_path("output")
        # output_csv = os.path.join(out_folder, 'apple_invoice_data.csv')
        output_csv = os.getcwd()+r'\output\apple_invoice_data.xlsx'
        # apple_data.to_csv(output_csv, index=False)  # Save as separate CSV
        with pd.ExcelWriter(output_csv) as writer:
            apple_data.to_excel(writer, index=False)
        st.success(f"Apple data successfully extracted! CSV file saved at {output_csv}")

def extract_invoice_data_ban(pdf_path):
    try:
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)
            page = reader.pages[0]
            text = page.extract_text()

        title = text.splitlines()[2].strip() if len(text.splitlines()) > 1 else "Not found"

        vendor = re.search(r'Beneficiary Name : (.+)', text)
        vendor = vendor.group(1).strip() if vendor else "Not found"
        ### 
        try:
            vendor_address = lines[3]
        except:
            vendor_address = 'Not Found'

        try:
            gst_reg_no = re.findall(r'M\d-\d{7}-\d', text)[0].strip()
        except:
            gst_reg_no = 'Not Found'
        ###
        sold_to = re.search(r'BILL TO:\s+(.+\n.+\n(.+))', text)
        sold_to = sold_to.group(2).strip() if sold_to else "Not found"
        sold_to_address = re.search(r'BILL TO:\s+(.+\n.+\n.+\n(.+))', text)
        sold_to_address = sold_to_address.group(2).strip() if sold_to_address else "Not found"
        tax_invoice_number = re.search(r'NUMBER INV DATE\s+([A-Za-z0-9]+)', text)
        tax_invoice_number = tax_invoice_number.group(1) if tax_invoice_number else "Not found"
        tax_invoice_date = re.search(r'NUMBER INV DATE\s+[A-Za-z0-9]+\s+(\d{2}/\d{2}/\d{4})', text)
        tax_invoice_date = tax_invoice_date.group(1) if tax_invoice_date else "Not found"
        currency = re.search(r'WAREHOUSE CURRENCY TERMS DUE DATE\s+([A-Za-z]+)\s+(\S+)', text)
        currency = currency.group(2).strip() if currency else "Not found"
        gst_amount = re.search(r'GST @ 9%\s+\$(\d+\.\d+)', text)
        gst_amount = gst_amount.group(1) if gst_amount else "Not found"
        lines = text.splitlines()
        last_six_lines = "\n".join(lines[-6:])  # Get the last 6 lines
        total_amount = re.search(r'TOTAL \$(\d+\.\d+)', last_six_lines)
        total_amount = total_amount.group(1) if total_amount else "Not found"

        return {
            "Title": title,
            "Vendor": vendor,
            "Vendor Address": vendor_address,
            "GST Reg. No.": gst_reg_no,
            "Sold To": sold_to,
            "Sold To Address": sold_to_address,
            "Invoice Number": tax_invoice_number,
            "Invoice Date": tax_invoice_date,
            "Currency": currency,
            "GST Amount": gst_amount,
            "Total Amount": total_amount,
        }
    except Exception as e:
        return {"Error": str(e)}

def process_ban_and_export():
    # folder_path = construct_folder_path("BAN LEONG TECHNOLOGIES LTD")
    folder_path = './BAN LEONG TECHNOLOGIES LTD'
    
    data = []
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(folder_path, file_name)
            extracted_data = extract_invoice_data_ban(file_path)
            extracted_data["File Name"] = file_name
            data.append(extracted_data)

    ban_data = pd.DataFrame(data)
    
    if ban_data.empty:
        st.error("No Ban Leong Tech data extracted.")
    else:
        # out_folder = construct_folder_path("output")
        # output_csv = os.path.join(out_folder, 'ban_leong_invoice_data.csv')
        output_csv = os.getcwd()+r'\output\ban_leong_invoice_data.xlsx'
        # ban_data.to_csv(output_csv, index=False)  # Save as separate CSV
        with pd.ExcelWriter(output_csv) as writer:
            ban_data.to_excel(writer, index=False)
        st.success(f"Ban Leong Tech data successfully extracted! CSV file saved at {output_csv}")

def extract_invoice_data_digihub(pdf_path):
    try:
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)
            page = reader.pages[0]
            text = page.extract_text()

        title = text.splitlines()[2].strip() if len(text.splitlines()) > 1 else "Not found"

        vendor = re.search(r'Beneficiary Name : (.+)', text)
        vendor = vendor.group(1).strip() if vendor else "Not found"
        sold_to = re.search(r'BILL TO:\s+(.+\n.+\n(.+))', text)
        sold_to = sold_to.group(2).strip() if sold_to else "Not found"
        sold_to_address = re.search(r'BILL TO:\s+(.+\n.+\n.+\n(.+))', text)
        sold_to_address = sold_to_address.group(2).strip() if sold_to_address else "Not found"       
        tax_invoice_number = re.search(r'NUMBER INV DATE\s+([A-Za-z0-9]+)', text)
        tax_invoice_number = tax_invoice_number.group(1) if tax_invoice_number else "Not found"
        tax_invoice_date = re.search(r'NUMBER INV DATE\s+[A-Za-z0-9]+\s+(\d{2}/\d{2}/\d{4})', text)
        tax_invoice_date = tax_invoice_date.group(1) if tax_invoice_date else "Not found"
        currency = re.search(r'WAREHOUSE CURRENCY TERMS DUE DATE\s+([A-Za-z]+)\s+(\S+)', text)
        currency = currency.group(2).strip() if currency else "Not found"
        gst_amount = re.search(r'GST @ 9%\s+\$(\d+\.\d+)', text)
        gst_amount = gst_amount.group(1) if gst_amount else "Not found"
        lines = text.splitlines()
        last_six_lines = "\n".join(lines[-6:])  # Get the last 6 lines
        total_amount = re.search(r'TOTAL \$(\d+\.\d+)', last_six_lines)
        total_amount = total_amount.group(1) if total_amount else "Not found"

        return {
            "Title": title,
            "Vendor": vendor,
            "Vendor Address": "nil",
            "GST Reg. No.": "nil",
            "Sold To": sold_to,
            "Sold To Address": sold_to_address,
            "Invoice Number": tax_invoice_number,
            "Invoice Date": tax_invoice_date,
            "Currency": currency,
            "GST Amount": gst_amount,
            "Total Amount": total_amount,
        }
    except Exception as e:
        return {"Error": str(e)}

def process_digihub_and_export():
    # folder_path = construct_folder_path("DIGITAL HUB PTE LTD")
    folder_path = './DIGITAL HUB PTE LTD'
    
    data = []
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(folder_path, file_name)
            extracted_data = extract_invoice_data_digihub(file_path)
            extracted_data["File Name"] = file_name
            data.append(extracted_data)

    digihub_data = pd.DataFrame(data)
    
    if digihub_data.empty:
        st.error("No Digital Hub data extracted.")
    else:
        # out_folder = construct_folder_path("output")
        # output_csv = os.path.join(out_folder, 'digital_hub_invoice_data.csv')
        output_csv = os.getcwd()+r'\output\digital_hub_invoice_data.xlsx'
        # digihub_data.to_csv(output_csv, index=False)  # Save as separate CSV
        with pd.ExcelWriter(output_csv) as writer:
            digihub_data.to_excel(writer, index=False)
        st.success(f"Digital Hub data successfully extracted! CSV file saved at {output_csv}")

def extract_invoice_data_consys(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            text = page.extract_text()

        lines = text.splitlines()
        title = lines[3].strip() if len(lines) > 3 else "Not found"
        vendor = lines[0].strip() if len(lines) > 0 else "Not found"
        vendor_address = lines[1].strip() if len(lines) > 1 else "Not found"
        gst_reg_no = re.search(r'GST No\. : (\S+)', text)
        gst_reg_no = gst_reg_no.group(1) if gst_reg_no else "Not found"
        customer_index = text.find("CUSTOMER: SHIPPING:")
        
        if customer_index != -1:
            sold_to = lines[lines.index("CUSTOMER: SHIPPING:") + 1].strip() if len(lines) > lines.index("CUSTOMER: SHIPPING:") + 1 else "Not found"
            sold_to_address = lines[lines.index("CUSTOMER: SHIPPING:") + 2].strip() if len(lines) > lines.index("CUSTOMER: SHIPPING:") + 2 else "Not found"
        else:
            sold_to = sold_to_address = "Not found"

        date_invoice_index = text.find("P . O . N O . O R D E R E D BY ACCOUNT NO. PAGE PAYMENT TERMS DATE INVOICE NO.")
        if date_invoice_index != -1:
            invoice_line = lines[lines.index("P . O . N O . O R D E R E D BY ACCOUNT NO. PAGE PAYMENT TERMS DATE INVOICE NO.") + 1] if len(lines) > lines.index("P . O . N O . O R D E R E D BY ACCOUNT NO. PAGE PAYMENT TERMS DATE INVOICE NO.") + 1 else ""
            
            invoice_date = re.search(r'\d{2}/\d{2}/\d{4}', invoice_line)
            invoice_number = re.search(r'\d+$', invoice_line)
            invoice_date = invoice_date.group(0) if invoice_date else "Not found"
            invoice_number = invoice_number.group(0) if invoice_number else "Not found"
        else:
            invoice_date = invoice_number = "Not found"

        currency = re.search(r'NETT TOTAL:\s+([A-Za-z]{3})', text)
        currency = currency.group(1) if currency else "Not found"
        gst_amount = re.search(r'GST 9%\s+(\d+\.\d+)', text)
        gst_amount = gst_amount.group(1) if gst_amount else "Not found"
        total_amount = re.search(r'NETT TOTAL:.*\s+(\d+\.\d+)', text)
        total_amount = total_amount.group(1) if total_amount else "Not found"

        return {
            "Title": title,
            "Vendor": vendor,
            "Vendor Address": vendor_address,
            "GST Reg. No.": gst_reg_no,
            "Sold To": sold_to,
            "Sold To Address": sold_to_address,
            "Invoice Number": invoice_number,
            "Invoice Date": invoice_date,
            "Currency": currency,
            "GST Amount": gst_amount,
            "Total Amount": total_amount,
        }
    except Exception as e:
        return {"Error": str(e)}

    # folder_path = construct_folder_path("CONVERGENT SYSTEMS")
    folder_path = './CONVERGENT SYSTEMS'
    
    data = []
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(folder_path, file_name)
            extracted_data = extract_invoice_data_consys(file_path)
            extracted_data["File Name"] = file_name
            data.append(extracted_data)

    consys_data = pd.DataFrame(data)
    
    if consys_data.empty:
        st.error("No Convergent Systems data extracted.")
    else:
        # out_folder = construct_folder_path("output")
        # output_csv = os.path.join(out_folder, 'convergent_systems_invoice_data.csv')
        output_csv = os.getcwd()+r'\output\convergent_systems_invoice_data.xlsx'
        # consys_data.to_csv(output_csv, index=False)  # Save as separate CSV
        with pd.ExcelWriter(output_csv) as writer:
            consys_data.to_excel(writer, index=False)
        st.success(f"Convergent Systems data successfully extracted! CSV file saved at {output_csv}")

def extract_invoice_data_ifactory(pdf_path):
    try:
        doc = fitz.open(pdf_path)  
      
        page = doc.load_page(0)  
        text = page.get_text("text")  


        lines = text.splitlines()  

        title = "Not found"
        for i, line in enumerate(lines):
            if "DATE :" in line and i + 1 < len(lines):
                title = lines[i + 1].strip()
                break

        vendor = lines[0].strip() if len(lines) > 0 else "Not found"

        vendor_address = "Not found"
        for i, line in enumerate(lines):
            if "UPC CODE" in line and i + 1 < len(lines):
                vendor_address = lines[i + 1].strip()
                break

        gst_reg_no = lines[4].strip() if len(lines) > 4 else "Not found"

        sold_to = "Not found"
        for i, line in enumerate(lines):
            if "TAX INVOICE" in line and i + 1 < len(lines):
                sold_to = lines[i + 1].strip()
                break

        sold_to_address = "Not found"
        if sold_to != "Not found" and i + 2 < len(lines):
            sold_to_address = lines[i + 2].strip()

        invoice_number = "Not found"
        for i, line in enumerate(lines):
            if "NO. :" == line.strip() and i + 1 < len(lines):  
                invoice_number = lines[i + 1].strip()  
                break

        invoice_date = "Not found"
        for i, line in enumerate(lines):
            if "PG NO. :" in line and i - 1 >= 0:
                invoice_date = lines[i - 1].strip() 
                break

        currency = "Not found"
        for line in lines:
            if "AMOUNT" in line:
                match = re.search(r'AMOUNT\s+([A-Za-z])', line)
                if match:
                    currency = match.group(1)
                    if currency == 'S':
                        currency = 'SGD'
                break

        gst_amount = lines[-2].strip() if len(lines) > 1 else "Not found"
        total_amount = lines[-1].strip() if len(lines) > 0 else "Not found"

        return {
            "Title": title,
            "Vendor": vendor,
            "Vendor Address": vendor_address,
            "GST Reg. No.": gst_reg_no,
            "Sold To": sold_to,
            "Sold To Address": sold_to_address,
            "Invoice Number": invoice_number,
            "Invoice Date": invoice_date,
            "Currency": currency,
            "GST Amount": gst_amount,
            "Total Amount": total_amount,
        }
    except Exception as e:
        return {"Error": str(e)}

def process_ifactory_and_export():
    folder_path = './iFactory Asia Pte Ltd'
    
    data = []
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(folder_path, file_name)
            extracted_data = extract_invoice_data_ifactory(file_path)
            extracted_data["File Name"] = file_name
            data.append(extracted_data)

    ifactory_data = pd.DataFrame(data)
    
    if ifactory_data.empty:
        st.error("No iFactory Asia data extracted.")
    else:
        output_csv = os.getcwd()+r'\output\ifactory_invoice_data.xlsx'
        with pd.ExcelWriter(output_csv) as writer:
            ifactory_data.to_excel(writer, index=False)
        st.success(f"iFactory Asia data successfully extracted! CSV file saved at {output_csv}")

def extract_invoice_data_pivene(pdf_path):
    try:
        doc = fitz.open(pdf_path)  
        
        page = doc.load_page(0)  
        text = page.get_text("text")  

        lines = text.splitlines()

        vendor = "Not found"
        for i, line in enumerate(lines):
            if line.startswith("Account Name: "):
                vendor = (line[14:]).strip()
                break

        vendor_address = "Not found"
        for i, line in enumerate(lines):
            if line.startswith("Tel:  "):
                vendor_address = lines[i - 1].strip()
                break

        gst_reg_no = "Not found"
        for i, line in enumerate(lines):
            if line.startswith("Business / GST Registration No.: "):
                gst_reg_no = (line[34:]).strip()
                break

        gst_amount = "Not found"
        for i, line in enumerate(lines):
            if line == ('GST Amount'):
                gst_amount = (lines[i + 3]).strip()
                break

        invoice_date = "Not found"
        for i, line in enumerate(lines):
            if line == ('Document Date'):
                invoice_date = (lines[i + 1]).strip()
                break

        total_amount = "Not found"
        for i, line in enumerate(lines):
            if line == ('Grand Total (SGD)'):
                total_amount = (lines[i + 1]).strip()
                break

        currency = "Not found"
        for i, line in enumerate(lines):
            if line == ('Grand Total (SGD)'):
                currency = (line[13:-1]).strip()
                break

        if len(re.findall(r'Tax Invoice', text))>0:
            title = 'Tax Invoice'
            sold_to = get_value(lines, 'Bill To', -8)
            sold_to_address = get_value(lines, 'Bill To', -6)
            invoice_number = get_value(lines, 'Tax Invoice No.', 1)
        else:
            title = 'Credit Note'
            sold_to = ''
            sold_to_address = ''
            invoice_number = ''

        return {
            'Title': title,
            "Vendor": vendor,
            'Vendor Address': vendor_address,
            "GST Reg. No.": gst_reg_no,
            "Sold To": sold_to,
            "Sold To Address": sold_to_address,
            "Invoice Number": invoice_number,
            "Invoice Date": invoice_date,
            "Currency": currency,
            "GST Amount": gst_amount,
            "Total Amount": total_amount,
            }

    except Exception as e:
        return {"Error": str(e)}

def process_pivene_and_export():
    folder_path = './PIVENE PTE LTD'
    
    data = []
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(folder_path, file_name)
            extracted_data = extract_invoice_data_pivene(file_path)
            extracted_data["File Name"] = file_name
            data.append(extracted_data)

    pivene_data = pd.DataFrame(data)
    
    if pivene_data.empty:
        st.error("No Pivene data extracted.")
    else:
        output_csv = os.getcwd()+r'\output\pivene_invoice_data.xlsx'
        with pd.ExcelWriter(output_csv) as writer:
            pivene_data.to_excel(writer, index=False)
        st.success(f"Pivene data successfully extracted! CSV file saved at {output_csv}")

####################################
#### Aggregate All Invoice Data ####
####################################

def process_invoice_data(invoice_type, extract_function, folder_name):
    try:
        folder_path = construct_folder_path(folder_name)
        print(folder_path)
        if not os.path.exists(folder_path):  # Check if folder exists
            return None

        data = []
        for file_name in os.listdir(folder_path):
            if file_name.endswith(".pdf"):
                file_path = os.path.join(folder_path, file_name)
                extracted_data = extract_function(file_path)
                extracted_data["File Name"] = file_name
                data.append(extracted_data)

        return pd.DataFrame(data)

    except Exception as e:
        st.warning(f"Error processing {invoice_type} invoices: {e}")
        return None
    
def aggregate_all_invoice_data():
    all_data = []

    # Process each source
    crystalwines_data = process_invoice_data("CRYSTALWINES", extract_invoice_data_crystalwines, "CRYSTAL WINES PTE LTD")
    if crystalwines_data is not None: all_data.append(crystalwines_data)

    setelco_data = process_invoice_data("SETELCO", extract_invoice_data_setelco, "SETELCO COMMUNICATIONS")
    if setelco_data is not None: all_data.append(setelco_data)
    
    dhl_data = process_invoice_data("DHL", extract_invoice_data_dhl, "DHL EXPRESS (SINGAPORE) PTE LTD")
    if dhl_data is not None: all_data.append(dhl_data)

    dfass_data = process_invoice_data("DFASS", extract_invoice_data_dfass, "DFASS (SINGAPORE) PTE LTD")
    if dfass_data is not None: all_data.append(dfass_data)

    kris_data = process_invoice_data("Kris+", extract_invoice_data_kris, "KRIS+ PTE. LTD")
    if kris_data is not None: all_data.append(kris_data)

    apple_data = process_invoice_data("Apple", extract_invoice_data_apple, "Apple")
    if apple_data is not None: all_data.append(apple_data)

    ban_data = process_invoice_data("Ban Leong", extract_invoice_data_ban, "BAN LEONG TECHNOLOGIES LTD")
    if ban_data is not None: all_data.append(ban_data)

    digihub_data = process_invoice_data("Digital Hub", extract_invoice_data_digihub, "DIGITAL HUB PTE LTD")
    if digihub_data is not None: all_data.append(digihub_data)

    consys_data = process_invoice_data("Convergent Systems", extract_invoice_data_consys, "CONVERGENT SYSTEMS")
    if consys_data is not None: all_data.append(consys_data)

    ifactory_data = process_invoice_data("iFactory", extract_invoice_data_ifactory, "iFactory Asia Pte Ltd")
    if ifactory_data is not None: all_data.append(ifactory_data)

    pivene_data = process_invoice_data("Pivene", extract_invoice_data_pivene, "PIVENE PTE LTD")
    if pivene_data is not None: all_data.append(pivene_data)

    # Combine all extracted data into a single DataFrame
    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        return combined_df
    else:
        return all_data
    
def save_combined_csv():
    combined_df = aggregate_all_invoice_data()

    # Convert DataFrame to CSV
    csv_buffer = io.StringIO()
    combined_df.to_csv(csv_buffer, index=False)
    csv_bytes = csv_buffer.getvalue().encode('utf-8')

    # Show success message and provide download button
    st.success("CSV file generated!")
    st.download_button(
        label="Download CSV",
        data=csv_bytes,
        file_name="invoice_data.csv",
        mime="text/csv"
    )

create_folder_structure(folder_list)

# Streamlit app layout
st.set_page_config(page_title="Automated Invoice Processor")
st.title("Automated Invoice Processing")

if st.button("Click here to start the process"):
    save_combined_csv()