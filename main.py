import docx                                                                                         # pip install python-docx
from nltk.stem import SnowballStemmer                                                               # pip install nltk
from nltk.corpus import stopwords
from nltk import tokenize, FreqDist
from tika import parser                                                                             # pip install tika
from langdetect import detect                                                                       # pip install langdetect
import re
import pandas as pd                                                                                 # pip install pandas
import os

fdist1 = ""
language = {
    "nl": "dutch",
    "en": "english",
    "es": "spanish",
    "fr": "french",
    "de": "german",
    "it": "italian"
    }

directory_path = input('What is the path of the directory? ')                                       # step 1: Select file
directory = os.fsencode(directory_path)

report = docx.Document()
report.add_heading(f'Analysis {os.path.basename(directory_path)}', 0)

for file in os.listdir(directory):
    document_path = os.path.join(directory, file)
    document = parser.from_file(document_path)
    document = document['content']                                                                  # step 2: retrieve text from file
    content = re.sub("[^a-zA-Z1-9|^-]", " ", document).lower()                                      # step 3: Delete all punctuation/upper case letters
    content_words = tokenize.word_tokenize(content)                                                 # step 4: Split words into list
    language_name = language[detect(content)]
    content_words_core = [w for w in content_words if w not in stopwords.words(language_name)]      # step 5: Delete adverbs
    stemmed_words = [SnowballStemmer(language_name).stem(word) for word in content_words_core]      # step 6: Group different forms of a word to a single item
    for words in content_words_core:
        fdist1 = FreqDist(stemmed_words)                                                            # step 7: Count occurrence of words
    top_10_words = pd.DataFrame(fdist1.most_common(10), columns=['Word', 'Count'])                  # step 8: Put top 10 words in table
    filename = os.fsdecode(file)

    title = report.add_heading(filename, level=1)
    text_language = report.add_paragraph(f'Language: {language_name.capitalize()}')
    table = report.add_table(top_10_words.shape[0]+1, top_10_words.shape[1])                        # step 9: Add template table to .docx file
    for j in range(top_10_words.shape[-1]):                                                         # step 10: Add headers to table
        table.cell(0, j).text = top_10_words.columns[j]
    for i in range(top_10_words.shape[0]):                                                          # step 11: Add data to table
        for j in range(top_10_words.shape[-1]):
            table.cell(i+1, j).text = str(top_10_words.values[i, j])
        table.style = 'Light Shading'                                                               # step 12: Change style of table


report.save('report.docx')                                                                          # step 13: Save document
