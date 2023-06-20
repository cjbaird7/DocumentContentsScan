import os
from datetime import datetime
from datetime import timedelta
from openpyxl import Workbook
import time
import pandas as pd
import PyPDF2
from docx import Document

def read_pdf(file_path):
    with open(file_path, 'rb') as file:
        pdf = PyPDF2.PdfFileReader(file)
        text = ''
        for page_num in range(pdf.getNumPages()):
            text += pdf.getPage(page_num).extractText()
        return text

def read_docx(file_path):
    doc = Document(file_path)
    text = ' '.join([para.text for para in doc.paragraphs])
    return text

def read_excel(file_path):
    df = pd.read_excel(file_path)
    return df.to_string()

def clean_text(text):
    # Illegal characters to remove
    illegal_chars = [
        chr(0x00), chr(0x01), chr(0x02), chr(0x03), chr(0x04), chr(0x05), chr(0x06), chr(0x07), 
        chr(0x08), chr(0x0B), chr(0x0C), chr(0x0E), chr(0x0F), chr(0x10), chr(0x11), chr(0x12), 
        chr(0x13), chr(0x14), chr(0x15), chr(0x16), chr(0x17), chr(0x18), chr(0x19), chr(0x1A),
        chr(0x1B), chr(0x1C), chr(0x1D), chr(0x1E), chr(0x1F)
    ]
    
    # Create a translation table that maps each illegal character into None
    remove_illegal_chars = dict.fromkeys(map(ord, illegal_chars), None)
    
    return text.translate(remove_illegal_chars)


def split_text_to_excel_columns(text, max_chars=32767):
    '''Splits a text into several parts, each with a maximum number of characters'''
    # Illegal characters to remove
    illegal_chars = [
        chr(0x00), chr(0x01), chr(0x02), chr(0x03), chr(0x04), chr(0x05), chr(0x06), chr(0x07), 
        chr(0x08), chr(0x0B), chr(0x0C), chr(0x0E), chr(0x0F), chr(0x10), chr(0x11), chr(0x12), 
        chr(0x13), chr(0x14), chr(0x15), chr(0x16), chr(0x17), chr(0x18), chr(0x19), chr(0x1A),
        chr(0x1B), chr(0x1C), chr(0x1D), chr(0x1E), chr(0x1F)
    ]
    
    # Create a translation table that maps each illegal character into None
    remove_illegal_chars = dict.fromkeys(map(ord, illegal_chars), None)
    
    text = text.translate(remove_illegal_chars)

    return [(text[i:i+max_chars]) for i in range(0, len(text), max_chars)]

def main():
    start_time = time.time() #start timer
    last_save_time = time.time()  # Time of the last save
    files_processed = 0  # Number of files processed

    dir_paths = ["F:\\QUS\\STAUS"]
    base_output_file = "C:\\Users\\baicla\\Documents\\file_contents.xlsx"  # The output Excel file

    output_file = base_output_file
    count = 1
    while os.path.exists(output_file):
        base_name, ext = os.path.splitext(base_output_file)
        output_file = f"{base_name}_{count}{ext}"
        count += 1

    workbook = Workbook()
    sheet = workbook.active
    warnings_sheet = workbook.create_sheet('Warnings')
    sheet.append(['File Name', 'File Path', 'Date Created', 'Date Last Modified', 'File Contents 1', 'File Contents 2', 'File Contents 3'])

    for dir_path in dir_paths:
        for foldername, subfolders, filenames in os.walk(dir_path):
            for filename in filenames:
                if filename.startswith('~$'): # Check if the filename starts with ~$
                    continue # Skip this file and go to the next one
                full_path = os.path.join(foldername, filename)
                file_path = r"\\?\\" + full_path

                try:
                    file_extension = os.path.splitext(filename)[1]
                    if file_extension == '.pdf':
                        file_content = read_pdf(file_path)
                    elif file_extension in ['.doc', '.docx']:
                        file_content = read_docx(file_path)
                    elif file_extension in ['.xls', '.xlsx']:
                        file_content = read_excel(file_path)
                    else:
                        continue  # Skip file if it's not pdf/doc/docx/xls/xlsx

                except PyPDF2.errors.PdfReadError as e:
                    print(f"Failed to parse PDF file: {file_path}")
                    print(f"Error: {str(e)}")
                    cleaned_error = clean_text(str(e))  # Clean the error message
                    warnings_sheet.append([cleaned_error, file_path]) # Append error and filepath to warnings worksheet
                    continue  # Skip to the next file

                except Exception as e:
                    print(f"Failed to parse file: {file_path}")
                    print(f"Error: {str(e)}")
                    cleaned_error = clean_text(str(e).encode('utf-8', errors='ignore').decode('utf-8'))  # Clean the error message
                    warnings_sheet.append([cleaned_error, file_path])  # Append cleaned error and filepath to warnings worksheet
                    continue  # Skip to the next file

                # Replace consecutive spaces with a single space
                # Replace line breaks with a semi-colon
                file_content = clean_text(' '.join(file_content.split()).replace('\n', ';'))  # Clean the file content

                contents_columns = split_text_to_excel_columns(file_content)
                
                creation_time = os.path.getctime(file_path)
                modification_time = os.path.getmtime(file_path)

                row = [
                    filename,
                    file_path,
                    datetime.fromtimestamp(creation_time).strftime('%Y-%m-%d %H:%M:%S'),
                    datetime.fromtimestamp(modification_time).strftime('%Y-%m-%d %H:%M:%S'),
                    *contents_columns[:3]
                ]
                sheet.append(row)
                del row  # delete row from memory
                import gc
                gc.collect()  # call garbage collector to free up memory

                files_processed += 1
                print(f"Files processed: {files_processed}")

                # Save the workbook every 5 minutes
                current_time = time.time()
                if current_time - last_save_time > timedelta(minutes=5).total_seconds():
                    print("Saving workbook...")
                    workbook.save(filename=output_file)
                    last_save_time = current_time
                    print("Workbook saved.")

    # Save the workbook at the end of the script as well
    workbook.save(filename=output_file)

    end_time = time.time() #End timer
    elapsed_time = end_time - start_time
    print(f"Done. Task finished in {elapsed_time} seconds")

if __name__ == "__main__":
    main()

