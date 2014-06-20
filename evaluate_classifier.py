'''
Created on Jun 12, 2014

@author: Alejandro Montes García
@author: Julia Efremova
@license: GPL v2
@organization: Eindhoven University of Technology
'''
import random
import xlrd
import os.path
import nltk
from nltk.classify import apply_features
from nltk.corpus import stopwords
import string

# Relevant features in current literature should be added here
def document_features(document):
    document_words = set(document.split())
    features = { }
    for word in word_features:
        features['contains(%s)' % word] = (word in document_words)
    return features

# Path to excel file
wb = xlrd.open_workbook(os.path.join('/home/TUE/amontes/Documents/excelfile.xlsx'))

sh = wb.sheet_by_index(1)
categorized = []
uncategorized = []

# Exclude punctuation symbols
exclude = set(string.punctuation)

# For every registry in the file the text is extracted and put into an array separating documents that are classified or need to be classified
for i in range(1, sh.nrows):
    text = sh.cell(i, 7).value
    category = sh.cell(i, 11).value
    if sh.cell(i, 8).value:
        text += sh.cell(i, 8).value
    if sh.cell(i, 9).value:
        text += sh.cell(i, 9).value
    if not text or type(text) is float:
        continue
    else:
        text = ''.join(ch for ch in text if ch not in exclude)
    if category:
        categorized.append((text, category))
    else:
        uncategorized.append((text, category))

# Shuffle the categorized documents so that the order in which they are in the excel file does not affect the results
random.shuffle(categorized)

all_words = []
for d in categorized:
    all_words += d[0].split()

# Take the 2000 most relevant words
all_words = nltk.FreqDist(w.lower() for w in all_words if (not w in stopwords.words('dutch') and len(w) > 1))
word_features = all_words.keys()[:2000]

# Take the 500 first registers for testing and the rest for training
test_set = apply_features(document_features, categorized[:500])
train_set = apply_features(document_features, categorized[500:])

# Train the classifier
classifier = nltk.NaiveBayesClassifier.train(train_set)

print (nltk.classify.accuracy(classifier, test_set) * 100)