import time
import os
import re
import fitz  # PyMuPDF
import requests
import pandas as pd
import hashlib
from datetime import datetime
from tqdm import tqdm  # Advanced Progress Bar Library
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

# 1. Excel File settings
EXCEL_FILE = "input_data.xlsx"
if not os.path.exists(EXCEL_FILE):
    print(f"Error: '{EXCEL_FILE}' file nahi mili! Pehle is naam ki excel file banayein.")
    exit()

df = pd.read_excel(EXCEL_FILE)
valid_rows = df[df['Registration_No'].notna() & df['Link'].notna()]
total_vehicles_count = len(valid_rows)

# 2. Folder Structure Setup
MAIN_DOWNLOAD_DIR = os.path.join(os.path.abspath("."), "Downloaded_Images")
os.makedirs(MAIN_DOWNLOAD_DIR, exist_ok=True)

current_timestamp = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
BATCH_FOLDER_NAME = f"Batch_Vehicles_{total_vehicles_count}__{current_timestamp}"
CURRENT_BATCH_DIR = os.path.join(MAIN_DOWNLOAD_DIR, BATCH_FOLDER_NAME)
os.makedirs(CURRENT_BATCH_DIR, exist_ok=True)

print(f"[*] Main Folder: Downloaded_Images")
print(f"[*] Current Run Batch Folder: {BATCH_FOLDER_NAME}\n")
print(f"[*] Starting Engine for {total_vehicles_count} Vehicles...\n")

report_data = []

# 3. Chrome Headless Mode Setup
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--disable-gpu")
driver = webdriver.Chrome(options=chrome_options)

def extract_claim_from_text(text):
    match = re.search(r'([A-Z]{2}\d{8})', str(text))
    return match.group(1) if match else None

def extract_and_clean_pdf(pdf_path, output_folder, pdf_base_name):
    """PDF se single correct image extract karti hai aur duplicates/overlays block karti hai"""
    try:
        if os.path.getsize(pdf_path) == 0:
            if os.path.exists(pdf_path): os.remove(pdf_path)
            return 0
        
        doc = fitz.open(pdf_path)
        img_count = 0

        for page_index in range(len(doc)):
            page = doc[page_index]
            images = page.get_images(full=True)
            total_images_on_page = len(images)

            for img_index, img in enumerate(images):
                # 🎯 CRITICAL FIX: Agar ek page par multi-images (tukde ya layers) hain,
                # toh hum sirf PEHLI image (index 0) ko hi accept karenge. 
                # Baki saare overlay background components strictly skip ho jayenge.
                if total_images_on_page > 1 and img_index > 0:
                    continue

                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]

                img_hash = hashlib.md5(image_bytes).hexdigest()[:6]
                image_name = f"{pdf_base_name}_p{page_index+1}_{img_hash}.{image_ext}"
                image_path = os.path.join(output_folder, image_name)

                if os.path.exists(image_path):
                    continue  # Exact duplicate content data ko overwrite hone se bachayein

                with open(image_path, "wb") as f:
                    f.write(image_bytes)

                img_count += 1

        doc.close()
        if os.path.exists(pdf_path): os.remove(pdf_path)
        return img_count
            
    except Exception:
        if os.path.exists(pdf_path): os.remove(pdf_path)
        return 0

try:
    # MAIN VEHICLE PROGRESS BAR SETUP
    vehicle_progress_bar = tqdm(
        valid_rows.iterrows(), 
        total=total_vehicles_count,
        desc="[+] Overall Bulk Process", 
        bar_format="{l_bar}{bar:40}{r_bar}{bar:-10b}",
        colour="green"
    )

    for index, row in vehicle_progress_bar:
        reg_no = str(row['Registration_No']).strip()
        portal_url = str(row['Link']).strip()
        
        if "http" not in portal_url:
            continue
            
        vehicle_progress_bar.set_description(f"[-] Processing: {reg_no}")
        
        reg_folder = os.path.join(CURRENT_BATCH_DIR, reg_no)
        os.makedirs(reg_folder, exist_ok=True)
        
        # Report Trackers
        row_total_files = 0
        row_pdf_count = 0
        row_other_count = 0
        row_downloaded_files = 0
        row_failed_files = 0
        row_extracted_images_total = 0
        row_direct_images_saved = 0
        row_attempts = 1
        
        # Exact Timestamp for log
        row_time_stamp = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        
        driver.get(portal_url)
        time.sleep(6)
        
        current_page_url = driver.current_url
        discovered_claim_number = extract_claim_from_text(current_page_url) or extract_claim_from_text(portal_url)

        selenium_cookies = driver.get_cookies()
        session = requests.Session()
        for cookie in selenium_cookies:
            session.cookies.set(cookie['name'], cookie['value'])
            
        user_agent = driver.execute_script("return navigator.userAgent;")
        session.headers.update({"User-Agent": user_agent})
        
        try:
            dropdown_element = driver.find_element(By.ID, "selectedDocument")
            dropdown = Select(dropdown_element)
            options = [opt.text for opt in dropdown.options if opt.text.strip()]
            
            row_total_files = len(options)
            
            if not discovered_claim_number:
                for opt_name in options:
                    found_claim = extract_claim_from_text(opt_name)
                    if found_claim:
                        discovered_claim_number = found_claim
                        break
            
            if not discovered_claim_number:
                discovered_claim_number = "CV00168307"
            
            for doc_name in options:
                if doc_name.lower().endswith('.pdf'):
                    row_pdf_count += 1
                else:
                    row_other_count += 1

            # SUB-LEVEL PROGRESS BAR
            file_progress_bar = tqdm(
                options, 
                desc=f"   |-- Files ({reg_no})", 
                leave=False, 
                bar_format="        {l_bar}{bar:20}{r_bar}",
                colour="cyan"
            )

            for doc_name in file_progress_bar:
                file_progress_bar.set_description(f"   |-- Syncing: {doc_name[:20]}...")
                
                detected_claim = extract_claim_from_text(doc_name) or discovered_claim_number
                doc_name_lower = doc_name.lower()
                
                temp_file_name = f"{reg_no}_{doc_name}"
                file_save_path = os.path.join(reg_folder, temp_file_name)
                
                download_url = f"https://icma.royalsundaram.in/DocumentsViewer/viewdocuments.do?do=fetchDocumentInternalCall&Dataclass=EcmsClaims&DocIndex={doc_name}&proposalCode=&inwardCode=&claimNumber={detected_claim}&documentType=FINAL_SURVEY"
                
                file_success = False
                for attempt in range(1, 4):
                    row_attempts = max(row_attempts, attempt)
                    try:
                        response = session.get(download_url, timeout=60, stream=True)
                        
                        if response.status_code == 200:
                            with open(file_save_path, "wb") as f:
                                for chunk in response.iter_content(chunk_size=1024 * 1024):
                                    if chunk:
                                        f.write(chunk)
                            
                            if os.path.exists(file_save_path) and os.path.getsize(file_save_path) > 100:
                                if doc_name_lower.endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp')):
                                    row_downloaded_files += 1
                                    row_direct_images_saved += 1
                                    file_success = True
                                    break
                                else:
                                    pdf_name_without_ext = os.path.splitext(doc_name)[0]
                                    pdf_base_identifier = f"{reg_no}_{pdf_name_without_ext}"
                                    imgs_saved = extract_and_clean_pdf(file_save_path, reg_folder, pdf_base_identifier)
                                    
                                    if imgs_saved > 0:
                                        row_downloaded_files += 1
                                        row_extracted_images_total += imgs_saved
                                        file_success = True
                                        break
                                        
                        time.sleep(4)
                    except Exception:
                        time.sleep(4)
                
                if not file_success:
                    row_failed_files += 1
                    if os.path.exists(file_save_path): os.remove(file_save_path)
                        
                time.sleep(2)
                
        except Exception:
            row_failed_files = row_total_files
            
        report_data.append({
            "Registration_No": reg_no,
            "Link": portal_url,
            "Timestamp": row_time_stamp,
            "Total_Source_Files": row_total_files,
            "Total_PDF_Count": row_pdf_count,
            "Total_Other_Files_Count": row_other_count,
            "Downloaded_Files_Success": row_downloaded_files,
            "Failed_Files_Count": row_failed_files,
            "Total_Extracted_Images_From_PDF": row_extracted_images_total,
            "Total_Direct_Images_Saved": row_direct_images_saved,
            "Total_Images_In_Folder": (row_extracted_images_total + row_direct_images_saved),
            "Max_Attempts_Used": row_attempts
        })
            
except Exception as e:
    print(f"\nAn unexpected error occurred: {e}")

finally:
    if report_data:
        report_df = pd.DataFrame(report_data)
        
        report_file_name = f"Execution_Report_{current_timestamp}.csv"
        report_csv_path = os.path.join(CURRENT_BATCH_DIR, report_file_name)
        
        report_df.to_csv(report_csv_path, index=False)
        print("\n" + "="*60)
        print(f"[!] PROCESS COMPLETE!")
        print(f"[!] Dynamic Report Saved: {report_csv_path}")
        print("="*60)
        
    driver.quit()
