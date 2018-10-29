import xlrd
import re
import nltk
from nltk.tokenize import RegexpTokenizer
from nltk import bigrams, trigrams
from nltk.stem.lancaster import LancasterStemmer
from nltk.stem import WordNetLemmatizer
from xlrd import open_workbook
from xlutils.copy import copy
from collections import Counter
from nltk.corpus import wordnet
# needed imports
from matplotlib import pyplot as plt
from scipy.cluster.hierarchy import dendrogram, linkage, fcluster
import scipy
import numpy as np

rb = xlrd.open_workbook("C:/Users/Youngeun Kang/Desktop/iPhone4.xls")
s = rb.sheet_by_index(0)
rb_write = open_workbook("C:/Users/Youngeun Kang/Desktop/iPhone4_pre.xls")
wb = copy(rb_write)
s_wb = wb.get_sheet(0)
rb_w = open_workbook("C:/Users/Youngeun Kang/Desktop/iPhone4_index.xls")
ww = copy(rb_write)
s_ww = wb.get_sheet(0)

row_val = s.nrows
i = 0

stopwords = nltk.corpus.stopwords.words('english')
tokenizer = RegexpTokenizer("[\w’]+", flags=re.UNICODE)
wordnet_lemmatizer = WordNetLemmatizer()

item = ''
while i < row_val:
    cell_val = s.cell_value(i, 3)
    item += cell_val
    i += 1
    print(i)


#pattern = r'''(?x) ([A-Z]\.)+ | \w+(-\w+)* | \$?\d+(\.\d+)?%? | \.\.\. | [][.,;"'?():-_`]'''
pattern = '\w+|\$[\d\.]+|\S+/+&'
tokens = nltk.regexp_tokenize(item, pattern)
print(len(tokens))

tokens = [token.lower() for token in tokens if len(token) > 2]
tokens = [token for token in tokens if token not in stopwords]

#ls = LancasterStemmer()
#tokens = [ls.stem(token) for token in tokens]

tokens = [wordnet_lemmatizer.lemmatize(token) for token in tokens]
tags_en = nltk.pos_tag(tokens)

tokens_pos = ['/'.join(t[:-1]) for t in nltk.pos_tag(tokens) if ((t[1] == 'NN') & (t[0] not in stopwords))]
print(tokens_pos)
#or (t[1] == 'NNS') or (t[1] == 'NNP') or (t[1] == 'NNPS')

C = Counter(tokens_pos)
sorted_C = sorted(C.items(), key=lambda x:x[1], reverse=True)[:20]
print(sorted_C)

tokens_pos_ = list(set(tokens_pos))  # exclude redundancy
print(len(tokens_pos_))

min_count = 10
word_dictionary = [word for word, freq in C.items() if freq >= min_count]
print(word_dictionary)
print(len(word_dictionary))

#tokens_pos_ = list(set(word_dictionary))  # exclude redundancy
#print(tokens_pos_)
#print(len(tokens_pos_))

#tokens_lem = [wordnet_lemmatizer.lemmatize(token) for token in tokens_pos_]
#print(tokens_lem)

i = 0
ar = len(word_dictionary)
syn_list = []
index = 0
while i < ar:
    try:
        s = wordnet.synsets(word_dictionary[i])
        for ss in s:
            lemma, pos, synset_index_str = ss.name().lower().rsplit('.', 2)
            if lemma == word_dictionary[i]:
                syn_list.append(ss)
                break
            else:
                print(lemma)
                print(word_dictionary[i])
    except:
        print(word_dictionary[i] + ": synset error happen.")
    i += 1

i = 0
for syn in syn_list:
    lemma, pos, synset_index_str = syn.name().lower().rsplit('.', 2)
    s_wb.write(i, 0, lemma)
    i += 1
wb.save("C:/Users/Youngeun Kang/Desktop/iPhone4_pre.xls")

x = np.zeros((len(syn_list), len(syn_list)))
i = 0
j = 0

while i < len(syn_list):
    while j < len(syn_list):
        if (None == syn_list[i].wup_similarity(syn_list[j])) or (None == syn_list[j].wup_similarity(syn_list[i])):
            x[i][j] = 100
        elif i == j:
            x[i][j] = 0
            if ((2/(syn_list[i].wup_similarity(syn_list[j])+syn_list[j].wup_similarity(syn_list[i])))-1) != 0:
                print(syn_list[i].name())
                print(syn_list[j].name())
                print((2/(syn_list[i].wup_similarity(syn_list[j])+syn_list[j].wup_similarity(syn_list[i])))-1)
        else:
            x[i][j] = (2/(syn_list[i].wup_similarity(syn_list[j])+syn_list[j].wup_similarity(syn_list[i])))-1
        j += 1
    print(i)
    i += 1
    j = 0

K = linkage(scipy.spatial.distance.squareform(x))
#y = M[np.triu_indices(n,1)]

# calculate full dendrogram
plt.figure(figsize=(25, 10))
plt.title('Hierarchical Clustering Dendrogram')
plt.xlabel('sample index')
plt.ylabel('distance')
dendrogram(
    K,
    leaf_rotation=90.,  # rotates the x axis labels
    leaf_font_size=8.,  # font size for the x axis labels
)
plt.show()

Z1 = dendrogram(
    K,
    leaf_rotation=90.,  # rotates the x axis labels
    leaf_font_size=8.,  # font size for the x axis labels
)

i = 0
for ii in Z1['ivl']:
    s_ww.write(i, 0, ii)
    print(ii)
    i += 1
ww.save("C:/Users/Youngeun Kang/Desktop/iPhone4_index.xls")  # 엑셀 파일 저장

print("debug")

#cb = wordnet.synset('book')
#ib = wordnet.synset('car')
#print(cb.wup_similarity(ib))

"""
tab = "    "
for synset in wordnet.synsets('car'):
    print("{}:".format(synset.name()))
    print(tab+"definition: {}".format(synset.definition()))
    print(tab+"pos: {}".format(synset.pos()))
    for e in synset.examples():
        print("    "+"example: {}".format(e))
    print()
"""