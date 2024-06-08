# %%
import os
import PyPDF2
import pandas as pd

def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text_by_page = {}
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            text_by_page[page_num + 1] = page.extract_text()
    return text_by_page

def find_keywords_in_text(text_by_page, keywords):
    found_keywords = []
    for page_num, text in text_by_page.items():
        for keyword in keywords:
            if keyword.lower() in text.lower():
                found_keywords.append((keyword, page_num))
    return found_keywords

def process_pdfs_in_folder(folder_path, keywords):
    data = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            text_by_page = extract_text_from_pdf(pdf_path)
            found_keywords = find_keywords_in_text(text_by_page, keywords)
            if found_keywords:
                for keyword, page_num in found_keywords:
                    data.append({
                        'Filename': filename,
                        'Keyword': keyword,
                        'Page': page_num
                    })
    return data

def save_to_excel(data, output_file):
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)

if __name__ == "__main__":
    folder_path = '../multiequipos/pdfs/'
    keywords = ['interiores', 'app', 'CNQX', 'portones', 'andenes']  # Add your keywords here
    output_file = 'output.xlsx'
    
    data = process_pdfs_in_folder(folder_path, keywords)
    save_to_excel(data, output_file)
    print(f'Excel file saved as {output_file}')
