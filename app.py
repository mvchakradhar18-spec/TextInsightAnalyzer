import streamlit as st
import pandas as pd
import requests
import re
import io

from collections import Counter
from bs4 import BeautifulSoup
from pypdf import PdfReader
from docx import Document
from openpyxl import load_workbook
from urllib.parse import urlparse


stop_words = {
    "the", "is", "are", "was", "were", "a", "an", "and",
    "or", "in", "on", "of", "to", "for", "with", "this",
    "that", "it", "as", "at", "by", "be", "been", "from",
    "have", "has", "had", "he", "she", "they", "we", "you",
    "i", "my", "your", "their", "our", "but", "not", "do",
    "does", "did", "can", "will", "would", "could"
}


def read_document(file):
    extension = file.name.split(".")[-1].lower()
    text = ""

    if extension == "pdf":
        pdf = PdfReader(io.BytesIO(file.read()))

        for page in pdf.pages:
            text += (page.extract_text() or "") + " "

    elif extension == "docx":
        doc = Document(io.BytesIO(file.read()))

        for paragraph in doc.paragraphs:
            text += paragraph.text + " "

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text += cell.text + " "

    elif extension == "xlsx":
        workbook = load_workbook(
            io.BytesIO(file.read()),
            data_only=True
        )

        for sheet in workbook.worksheets:
            for row in sheet.iter_rows(values_only=True):
                for cell in row:
                    if cell is not None:
                        text += str(cell) + " "

    return text


def read_webpage(url):
    parsed_url = urlparse(url)

    if parsed_url.scheme not in ["http", "https"]:
        raise ValueError("Please enter a valid webpage URL.")

    if not parsed_url.netloc:
        raise ValueError("Invalid webpage URL.")

    response = requests.get(
        url,
        headers={"User-Agent": "Mozilla/5.0"},
        timeout=15
    )

    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")

    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()

    text = soup.get_text(separator=" ")

    return text


def count_words(text):
    words = re.findall(r"[a-zA-Z]+", text.lower())

    filtered_words = []

    for word in words:
        if word not in stop_words:
            filtered_words.append(word)

    word_count = Counter(filtered_words)

    return word_count


st.title("TextInsight Analyzer")

st.write(
    "Upload PDF, Word or Excel files, or enter a webpage URL "
    "to analyze frequently used words."
)

st.header("Upload Documents")

files = st.file_uploader(
    "Choose your files",
    type=["pdf", "docx", "xlsx"],
    accept_multiple_files=True
)

st.header("Webpage URL")

url = st.text_input(
    "Enter webpage URL",
    placeholder="https://example.com"
)


if st.button("Analyze"):

    all_text = ""

    if files:
        for file in files:
            try:
                text = read_document(file)
                all_text += text + " "
                st.success(file.name + " read successfully")

            except Exception as e:
                st.error("Unable to read " + file.name + ": " + str(e))

    if url.strip():
        try:
            webpage_text = read_webpage(url.strip())
            all_text += webpage_text + " "
            st.success("Webpage read successfully")

        except Exception as e:
            st.error("Unable to read webpage: " + str(e))

    if all_text.strip():

        word_count = count_words(all_text)

        total_words = sum(word_count.values())
        unique_words = len(word_count)

        st.header("Analysis Results")

        col1, col2 = st.columns(2)

        col1.metric("Total Words", total_words)
        col2.metric("Unique Words", unique_words)

        top_words = word_count.most_common(10)

        result = pd.DataFrame(
            top_words,
            columns=["Word", "Frequency"]
        )

        st.subheader("Top 10 Frequently Used Words")

        st.dataframe(result, use_container_width=True)

        st.subheader("Word Frequency Chart")

        st.bar_chart(result.set_index("Word"))

        csv_file = result.to_csv(index=False).encode("utf-8")

        st.download_button(
            label="Download CSV",
            data=csv_file,
            file_name="textinsight_word_frequency.csv",
            mime="text/csv"
        )

    else:
        st.warning("Please upload a document or enter a webpage URL.")
