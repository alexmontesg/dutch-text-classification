'''
Created on Jun 12, 2014

@author: Alejandro Montes Garcia
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
from nltk.collocations import BigramCollocationFinder
from nltk.metrics import BigramAssocMeasures

def getBigrams(words, score_fn=BigramAssocMeasures.chi_sq, n=200):
    bigram_finder = BigramCollocationFinder.from_words(words)
    return bigram_finder.nbest(score_fn, n)

# Relevant features in current literature should be added here
def document_features(document):
    words = document.split()
    document_words = words
    for bigram in getBigrams(words):
        document_words.append(bigram[0] + ":" + bigram[1])
    document_words = set(document_words)
    features = { }
    for word in word_features:
        features['%s' % word] = (word in document_words)
    return features

# Path to excel file
wb = xlrd.open_workbook(os.path.join('/home/TUE/amontes/Dropbox/alejandro-julia/Database/notary_acts_19062014 - classified.xlsx'))

sh = wb.sheet_by_index(0)
documents = []

# Exclude punctuation symbols
exclude = set(string.punctuation)

# For every registry in the file the text is extracted and put into an array separating documents that are classified or need to be classified
for i in range(1, sh.nrows):
    text = sh.cell(i, 1).value
    category = sh.cell(i, 0).value
    text = ''.join(ch for ch in text if ch not in exclude)
    documents.append((text.lower(), category))

# Shuffle the documents so that the order in which they are in the excel file does not affect the results
random.shuffle(documents)

all_words = []
for d in documents:
    words = []
    for word in d[0].split():
        if(len(word) > 1 and not word in stopwords.words('dutch')):
            words.append(word)
    all_words += words + getBigrams(words)

# Take the 2000 most relevant words
all_words = nltk.FreqDist(all_words)
word_features = all_words.keys()[:2000]

n_word = 0
for w in word_features:
    if type(w) is tuple:
        word_features[n_word] = word_features[n_word][0] + ":" + word_features[n_word][1]
        print word_features[n_word]
    n_word += 1

# Take the 1000 first registers for testing and the rest for training
test_set = apply_features(document_features, documents[1000:2000])
train_set = apply_features(document_features, documents[:1000])

# Train the classifier
classifier = nltk.NaiveBayesClassifier.train(train_set)

print (nltk.classify.accuracy(classifier, test_set) * 100)