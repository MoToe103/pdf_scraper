import os
import PyPDF2
import re
from collections import Counter
import pandas as pd

folder_path = r"C:\Users\mbfly\My Drive\Professional\Resume\PDF Keywords"


def extract_text_from_pdfs(folder_path):
    pdf_texts = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf'):
            file_path = os.path.join(folder_path, filename)
            with open(file_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                text = ''
                for page in reader.pages:
                    text += page.extract_text()
                pdf_texts.append(text)
    return pdf_texts


def count_words(texts):
    word_count = Counter()
    for text in texts:
        words = re.findall(r'\b\w+\b', text.lower())
        word_count.update(words)
    return word_count


def write_to_excel(word_count, output_file):
    df = pd.DataFrame(word_count.items(), columns=['Word', 'Frequency'])
    df = df.sort_values(by='Frequency', ascending=False)
    df.to_excel(output_file, index=False)


def main(folder_path, output_file):
    pdf_texts = extract_text_from_pdfs(folder_path)
    word_count = count_words(pdf_texts)
    write_to_excel(word_count, output_file)
    print(f"Word frequencies have been written to {output_file}")


if __name__ == "__main__":
    folder_path = input("Enter the folder path containing PDFs: ")
    output_file = "resume_word_frequencies.xlsx"
    main(folder_path, output_file)
