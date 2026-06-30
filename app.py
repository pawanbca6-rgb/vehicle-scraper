import streamlit as st
import time
import os
import re
import fitz  # PyMuPDF
import requests
import pandas as pd
import hashlib
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import zipfile

st.set_page_config(page_title="Vehicle Document Extractor", layout="wide")

st.title("🚗 Vehicle Document Extractor Engine")
st.write("Upload an Excel sheet containing `Registration_No` and portal `Link` columns to automatically extract and download images.")

uploaded_file = st.file_uploader("Choose your input_data.xlsx file", type=["xlsx"])

def extract_claim_from_text(text):
    match = re.search(r'([A-Z]{2}\d{8})', str(text))
    return match.group(1) if match else None

def extract_and_clean_pdf(pdf_path, output_folder, pdf_base_name):
    try:
        if os.path.getsize(pdf_path) == 0:
            if os.path.exists(pdf_path): os.remove(pdf_path)
            return 0
        
        doc = fitz.open(pdf_path)
        img_count = 0
        for page_index in range(len(doc)):
            page = doc[page_index]
            images = page.get_images(full=True)
            for img_index, img in enumerate(images):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]

                img_hash = hashlib.md5(image_bytes).hexdigest()[:6]
                image_name = f"{pdf_base_name}_p{page_index+1}_i{img_index+1}_{img_hash}.{image_ext}"
                image_path = os.path.join(output_folder, image_name)

                if os.path.exists(image_path):
                    image_name = f"{pdf_base_name}_p{page_index+1}_i{img_index+1}_dup_{img_count}.{image_ext}"
                    image_path = os.path.join(output_folder, image_name)

                with open(image_path, "wb") as f:
                    f.write(image_bytes)
                img_count += 1
        doc.close()
        if os.path.exists(pdf_path): os.remove(pdf_path)
        return img_count
    except Exception:
        if os.path.exists(pdf_path): os.remove(pdf_path)
        return 0

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    # Validation check for required columns
    if 'Registration_No' not in df.columns or 'Link' not in df.columns:
        st.error("Error: Excel must contain 'Registration_No' and 'Link' columns.")
    else:
        valid_rows = df[df['Registration_No'].notna() & df['Link'].notna()]
        total_vehicles_count = len(valid_rows)
        st.success(f"Successfully loaded {total_vehicles_count} valid rows to process.")
        
        if st.button("🚀 Start Bulk Extraction Process"):
            # Temporary local engine directories inside container
            CURRENT_BATCH_DIR = os.path.abspath("Downloaded_Images")
            if os.path.exists(CURRENT_BATCH_DIR):
                import shutil
                shutil.rmtree(CURRENT_BATCH_DIR) # clear previous run
            os.makedirs(CURRENT_BATCH_DIR, exist_ok=True)
            
            # CRITICAL: Production Headless Chrome Settings for Linux Environments
            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_argument("--headless=new")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--window-size=1920,1080")
            
            # Points to the locations generated inside our Dockerfile environment
            chrome_options.binary_location = "/usr/bin/chromium"
            
            try:
                driver = webdriver.Chrome(options=chrome_options)
            except Exception as chrome_err:
                st.error(f"Failed to initialize Chrome Driver: {chrome_err}")
                st.stop()
            
            report_data = []
            
            # Progress tracking metrics
            main_progress = st.progress(0)
            status_text = st.empty()
            
            for idx, (index, row) in enumerate(valid_rows.iterrows()):
                reg_no = str(row['Registration_No']).strip()
                portal_url = str(row['Link']).strip()
                
                if "http" not in portal_url:
                    continue
                    
                status_text.markdown(f"**⏳ Processing ({idx+1}/{total_vehicles_count}):** `{reg_no}`")
                
                reg_folder = os.path.join(CURRENT_BATCH_DIR, reg_no)
                os.makedirs(reg_folder, exist_ok=True)
                
                row_total_files = 0
                row_pdf_count = 0
                row_other_count = 0
                row_downloaded_files = 0
                row_failed_files = 0
                row_extracted_images_total = 0
                row_direct_images_saved = 0
                row_attempts = 1
                row_time_stamp = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
                
                try:
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
                                                if chunk: f.write(chunk)
                                        
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
                        "Registration_No": reg_no, "Link": portal_url, "Timestamp": row_time_stamp,
                        "Total_Source_Files": row_total_files, "Total_PDF_Count": row_pdf_count,
                        "Total_Other_Files_Count": row_other_count, "Downloaded_Files_Success": row_downloaded_files,
                        "Failed_Files_Count": row_failed_files, "Total_Extracted_Images_From_PDF": row_extracted_images_total,
                        "Total_Direct_Images_Saved": row_direct_images_saved, 
                        "Total_Images_In_Folder": (row_extracted_images_total + row_direct_images_saved),
                        "Max_Attempts_Used": row_attempts
                    })
                except Exception as e:
                    st.warning(f"Error skipping vehicle row {reg_no}: {e}")
                
                main_progress.progress((idx + 1) / total_vehicles_count)
                
            driver.quit()
            status_text.text("✅ All Tasks Finished! Generating download files...")
            
            # Compiling Outputs for User download
            if report_data:
                report_df = pd.DataFrame(report_data)
                report_csv_path = "Execution_Report.csv"
                report_df.to_csv(report_csv_path, index=False)
                
                zip_path = "Downloaded_Images.zip"
                with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for root, dirs, files in os.walk(CURRENT_BATCH_DIR):
                        for file in files:
                            zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), os.path.dirname(CURRENT_BATCH_DIR)))
                
                st.balloons()
                
                # Visual Downloader Buttons
                col1, col2 = st.columns(2)
                with col1:
                    with open(zip_path, "rb") as fp:
                        st.download_button(label="📥 Download All Images (ZIP)", data=fp, file_name="extracted_vehicle_images.zip", mime="application/zip", use_container_width=True)
                with col2:
                    with open(report_csv_path, "rb") as rp:
                        st.download_button(label="📊 Download CSV Execution Report", data=rp, file_name="Execution_Report.csv", mime="text/csv", use_container_width=True)