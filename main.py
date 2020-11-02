from docx import Document                                                                       # pip install python-docx
from nltk.stem import WordNetLemmatizer                                                         # pip install nltk
from nltk.corpus import stopwords
from nltk import tokenize, FreqDist
import re
import pandas as pd                                                                             # pip install pandas

wnl = WordNetLemmatizer()


filename = input('What is the path of the file (including complete filename)? ')                # step 1: Select .docx file


document = Document(filename)                                                                   # step 2: retrieve text from .docx file


for content in document.paragraphs:
    content = re.sub("[^a-zA-Z1-9|^-]", " ", content.text).lower()                              # step 3: Delete all punctuation/upper case letters
    content_words = tokenize.word_tokenize(content)                                             # step 4: Split words into list
    content_words_core = [w for w in content_words if w not in stopwords.words("english")]      # step 5: Delete adverbs
    stemmed_words = [wnl.lemmatize(word) for word in content_words_core]                        # step 6: Group different forms of a word to a single item
    for words in content_words_core:
        fdist1 = FreqDist(stemmed_words)                                                        # step 7: Count occurrence of words
top_10_words = pd.DataFrame(fdist1.most_common(10), columns=['Word', 'Count'])                  # step 8: Put top 10 words in table

t = document.add_table(top_10_words.shape[0]+1, top_10_words.shape[1])                          # step 9: Add template table to .docx file
for j in range(top_10_words.shape[-1]):                                                         # step 10: Add headers to table
    t.cell(0, j).text = top_10_words.columns[j]
for i in range(top_10_words.shape[0]):                                                          # step 11: Add data to table
    for j in range(top_10_words.shape[-1]):
        t.cell(i+1, j).text = str(top_10_words.values[i, j])
t.style = 'Light Shading'                                                                       # step 12: Change style of table


document.save(f'rapport.{filename}')                                                            # step 13: Save document
