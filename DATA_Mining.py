import sys
# import fitz  # Membaca file PDF
import nltk  # Natural Language Toolkit
import re  # Menghapus karakter angka.
import string  # Menghapus karakter tanda baca.
import matplotlib.pyplot as plt  # Menggambarkan Frekuensi Kemunculan
import docx  # Membaca File docx
# from docx2pdf import convert  # Meng-convert docx ke pdf
import numpy as np
from Sastrawi.Stemmer.StemmerFactory import StemmerFactory  # Stemming Sastrawi
from Sastrawi.StopWordRemover.StopWordRemoverFactory import StopWordRemoverFactory, StopWordRemover, ArrayDictionary  # Stopword Sastrawi
import os  # Membaca file di dalam folder
from sklearn.metrics.pairwise import cosine_similarity
import PyPDF2
from sklearn.feature_extraction.text import TfidfVectorizer

#GUI Pyqt5
from PyQt5 import QtCore,QtWidgets
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi

class ShowGUI(QMainWindow):
    def __init__(self):
        super(ShowGUI, self).__init__()
        loadUi('Gui.ui', self)
        self.file = ""
        self.query = ""
        self.frequensi = {}
        self.actionOpen_3.triggered.connect(self.OpenClick)
        self.buttonQuery.clicked.connect(self.MasukanQuery)
        self.buttonsearch.clicked.connect(self.showFrequensi)
        self.msg = QMessageBox()
        self.msg.setIcon(QMessageBox.Information)
        self.msg.setStandardButtons(QMessageBox.Ok)
        self.msg.setMinimumHeight(100)
        self.msg.setMinimumWidth(100)
        self.msg.setStyleSheet("QLabel{min-width: 200px; min-height: 200px;}")

    def OpenClick(self):
        self.file = str(QFileDialog.getExistingDirectory(
            self, "Select Directory"))
        # self.label_path.setText(str(self.file))
        count = 0
        # Iterate directory
        for path in os.listdir(self.file):
            # check if current path is a file
            if os.path.isfile(os.path.join(self.file, path)):
                # check if file is pdf or docx
                if path.endswith(".pdf"):
                    count += 1
                elif path.endswith(".docx"):
                    count += 1
        self.label_jumlah.setText(str(count))

    def MasukanQuery(self):
        self.query = self.edit_query.toPlainText()
        print(self.query)
        if self.query == "":
            self.msg.about(self, "Information", "Query is empty")
        else:
            if self.file == "":
                self.msg.about(self, "Information", "File is empty")
            else:
                self.main()

    def showFrequensi(self):
        freq = self.edit_search.toPlainText()
        if self.frequensi != {}:
            if freq in self.frequensi:
                self.msg.about(self,"Information",str(self.frequensi[freq]))
            else:
                self.msg.about(self,"Information", "Term Tidak DI Temukan!")
        else:
            self.edit_search.setText("")

    def relevant_document(self, folder_path):
        #membuat stemmer objek
        factory = StemmerFactory()
        stemmer = factory.create_stemmer()

        #membaca documen dari folder
        documents = []
        filenames = []
        for filename in os.listdir(folder_path):
            if filename.endswith(".pdf"):
                filenames.append(filename)
                with open(os.path.join(folder_path, filename), "rb") as f:
                    pdf_reader = PyPDF2.PdfFileReader(f)
                    num_pages = pdf_reader.getNumPages()
                    document = ""
                    for page_num in range(num_pages):
                        page = pdf_reader.getPage(page_num)
                        document += page.extractText()
                    documents.append(document)

            elif filename.endswith(".docx"):
                filenames.append(filename)
                doc = docx.Document(os.path.join(folder_path, filename))
                document = ""
                for paragraph in doc.paragraphs:
                    document += paragraph.text
                documents.append(document)
        #proses stemming
        proses_dokumen = []
        for document in documents:
            document = document.lower()
            document = re.sub(r'\W', ' ', str(document))
            # remove all single characters
            document = re.sub(r'\s+[a-zA-Z]\s+', ' ', document)
            # Remove single characters from the start
            document = re.sub(r'\^[a-zA-Z]\s+', ' ', document)
            # Substituting multiple spaces with single space
            document = re.sub(r'\s+', ' ', document, flags=re.I)
            # Removing prefixed 'b'
            document = re.sub(r'^b\s+', '', document)
            #add stopword
            # factory = StopWordRemoverFactory()
            # more_stopword = ['aku','yang','ke','di','dengan']
            # factory.extend(more_stopword)
            # stopword = StopWordRemoverFactory().create_stop_word_remover()
            # document = stopword.remove(document)
            #Stemming
            proses_dokumen1 = " ".join([stemmer.stem(word) for word in document.split()])
            proses_dokumen.append(proses_dokumen1)

            # Create a vocabulary of all the unique terms in the documents
            vocab = set()
            for document in proses_dokumen:
                for word in document.split():
                    vocab.add(word)

            vocab = list(vocab)

        # Create a matrix of document vectors, with each row representing a document and each column representing a term
        matrix = np.zeros((len(proses_dokumen),len(vocab)))
        # Calculate the term frequency of each term in each document
        for i,document in enumerate(proses_dokumen):
            for j,term in enumerate(vocab):
                count = document.split().count(term)
                matrix[i,j] = document.split().count(term)
                if term in self.frequensi:
                    self.frequensi[term] += count
                else:
                    self.frequensi[term] = count
        #print term frequncy
        for term,frequency in self.frequensi.items():
            print(f"Term: {term}, Frequncy: {frequency}")

        # TF-IDF
        tfidf_vectorize = TfidfVectorizer()
        tfidf_matrix = tfidf_vectorize.fit_transform(document)

        #cosine
        similarity = cosine_similarity(tfidf_matrix)
        print(similarity)
    def main(self):
        # Define Folder path
        folder_path = self.file
        #define query
        query = self.query

        most_relevan_file = self.relevant_document(folder_path)
        print("Most Relevant Document: ")
        self.show_file.clear()
        for similarity, document, filename in most_relevan_file:
            if np.isnan(similarity):
                similarity = 0
                similarity_persen = round(similarity * 100)
                similarity = similarity
                self.show_file.append(f"Document: {filename}(similarity:{similarity})({similarity_persen}%)")


app = QtWidgets.QApplication(sys.argv)
window = ShowGUI()
window.setWindowTitle('Information Retrieval')
window.show()
sys.exit(app.exec_())
