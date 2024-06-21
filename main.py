import streamlit as st
import pandas as pd
import docx
import re
from io import BytesIO
from typing import Set, List

def extract_codes_from_excel_or_csv(file) -> Set[str]:
    if file.name.endswith('.xlsx'):
        df = pd.read_excel(file)
    elif file.name.endswith('.csv'):
        df = pd.read_csv(file)
    else:
        st.error("Unsupported file format!")
        return set()
    codes = set(df.iloc[:, 0].astype(str))
    return codes

def extract_codes_from_text_or_word(file, pattern: str) -> Set[str]:
    if file.name.endswith('.docx'):
        doc = docx.Document(file)
        text = "\n".join([para.text for para in doc.paragraphs])
    elif file.name.endswith('.txt'):
        text = file.read().decode('utf-8')
    else:
        st.error("Unsupported file format!")
        return set()
    codes = set(re.findall(pattern, text))
    return codes

def compare_codes(excel_codes: Set[str], word_codes: Set[str]) -> (Set[str], Set[str], Set[str]):
    matching_codes = excel_codes.intersection(word_codes)
    codes_only_in_excel = excel_codes - word_codes
    codes_only_in_word = word_codes - excel_codes
    return matching_codes, codes_only_in_excel, codes_only_in_word

def main():
    st.title("Code Comparison Tool")

    st.write("""
    Upload an Excel/CSV file containing codes and a Word/TXT document with text that may contain these codes.
    The application will compare the codes and list the matching ones and those that are only in one of the files.
    """)

    code_file = st.file_uploader("Upload Excel/CSV File for Codes", type=["xlsx", "csv"], accept_multiple_files=False)
    text_file = st.file_uploader("Upload Word/TXT File to Search Through", type=["docx", "txt"], accept_multiple_files=False)
    custom_pattern = st.text_input("Specify Custom RegEx Pattern (See this **[link](https://datasciencedojo.com/blog/regular-expression-101/)** or this [link](https://chatgpt.com/))", value=r'INT-\d+')

    if code_file and text_file:
        # Extract codes from the uploaded files
        excel_codes = extract_codes_from_excel_or_csv(code_file)
        word_codes = extract_codes_from_text_or_word(text_file, custom_pattern)

        # Compare codes
        matching_codes, codes_only_in_excel, codes_only_in_word = compare_codes(excel_codes, word_codes)

        # Equalize lengths by padding shorter lists with None
        max_len = max(len(matching_codes), len(codes_only_in_excel), len(codes_only_in_word))
        
        matching_codes_list = list(matching_codes) + [None] * (max_len - len(matching_codes))
        codes_only_in_excel_list = list(codes_only_in_excel) + [None] * (max_len - len(codes_only_in_excel))
        codes_only_in_word_list = list(codes_only_in_word) + [None] * (max_len - len(codes_only_in_word))

        # Create a DataFrame with equal-length lists
        result_df = pd.DataFrame({
            "Matching codes": matching_codes_list,
            "Codes only in Excel/CSV": codes_only_in_excel_list,
            "Codes only in Word/TXT": codes_only_in_word_list
        })

        # Provide option to download the results
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            result_df.to_excel(writer, index=False)

        processed_data = output.getvalue()

        st.download_button(label='📥 Download Results',
                           data=processed_data,
                           file_name='comparison_result.xlsx')
        


        col1, col2, col3 = st.columns(3)

        with col1:
            # Display results in the app
            st.write("### Matching Codes")
            st.dataframe(result_df[['Matching codes']].dropna())

        with col2:
            st.write("### Codes Only in Excel/CSV")
            st.dataframe(result_df[['Codes only in Excel/CSV']].dropna())

        with col3:
            st.write("### Codes Only in Word/TXT")
            st.dataframe(result_df[['Codes only in Word/TXT']].dropna())


if __name__ == "__main__":
    main()
