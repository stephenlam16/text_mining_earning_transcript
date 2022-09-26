#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Sep 25 14:00:54 2022

@author: Stephen
"""
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Sep 23 09:26:11 2022

@author: Stephen
"""
import sys
sys.modules[__name__].__dict__.clear()


import re
import textract
import string
import docx2txt
import pickle

#text = textract.process("/Users/Stephen/Desktop/python/text mining/transcripts/BOC Hong Kong (Holdings) Limited_Earnings Call_2022-03-29_English.docx")


#import glob
#text = ''
#for file in glob.glob('/Users/Stephen/Desktop/python/text mining/transcripts/*.docx'):
#    text += docx2txt.process(file)

transcripts_list = ['BOC Hong Kong (Holdings) Limited_Earnings Call_2020-08-31_English',
             'BOC Hong Kong (Holdings) Limited_Earnings Call_2021-03-30_English',
             'BOC Hong Kong (Holdings) Limited_Earnings Call_2021-08-30_English',
             'BOC Hong Kong (Holdings) Limited_Earnings Call_2022-03-29_English']

#file = open("/Users/Stephen/Desktop/python/text mining/transcripts/BOC Hong Kong (Holdings) Limited_Earnings Call_2020-08-31_English.docx", "rb")



#file = open("/Users/Stephen/Desktop/python/text mining/transcripts/"
#              + "BOC Hong Kong (Holdings) Limited_Earnings Call_2020-08-31_English" + ".docx", "rb")



def upload_transcripts(transcript):
    '''Returns transcript data specifically from scrapsfromtheloft.com.'''
    file = textract.process("/Users/Stephen/Desktop/python/text mining/transcripts/"
                            + transcript + ".docx")
    return(file)

    

text = [upload_transcripts(transcript) for transcript in transcripts_list]


data3 = {}


data = {}
for key in transcripts_list:
    for value in text:
        data[key] = value
        text.remove(value)
        break 
    
data.keys()

data['BOC Hong Kong (Holdings) Limited_Earnings Call_2020-08-31_English']




#### define a function to locate when a specific word appear the nth time
def find_nth(haystack, needle, n):
    start = haystack.find(needle)
    while start >= 0 and n > 1:
        start = haystack.find(needle, start+len(needle))
        n -= 1
    return start




#print(text)

for key in data:
    text = data[key]
    
    #### text is coded as bytes, for easier viewing decode it into string
    text2 = text.decode("utf-8") 


    tablecontent_position = text2.find("Table of Contents")

    next_n_follow = 600

    #print(text2[tablecontent_position:tablecontent_position+next_n_follow])

    list_of_next_follow = text2[tablecontent_position:tablecontent_position+next_n_follow]

    #### delete the boilerplate term of conditions from S&P
    SnP_terms = find_nth(text2, "These materials have been prepared solely for information", 1)
    text3 = text2[:SnP_terms]


    #### seperating the earning transcript into different sections
    #### Management presentation and Q&A
    if "Presentation" in list_of_next_follow:
    
        Presentation_start = find_nth(text2, "Presentation", 2) 
    
        if "Question and Answer" in list_of_next_follow:
    
            #### create a new sublist
            QnA_start = find_nth(text2, "Question and Answer", 2)
            Presentation_section = text3[Presentation_start:QnA_start]
            QnA_section = text3[QnA_start:]
        
        else:
            Presentation_section = text3[Presentation_start:]
    
    
    else:
    
        if "Question and Answer" in list_of_next_follow:
            QnA_start = find_nth(text2, "Question and Answer", 2)
            QnA_section = text3[QnA_start:]
    
        else:    
            pass

    Transcript_section_name = [key + 'Presentation', key + 'Question and Answer']
    
    Sections = [Presentation_section, QnA_section]
    
    data2 = {}
    for key2 in Transcript_section_name:
        for value2 in Sections:
            data2[key2] = value2
            Sections.remove(value2)
            break 
    
    
    data3 = {**data3, **data2}
    
    data2. clear()
    
    text2 = ''
    tablecontent_position = ''
    next_n_follow = ''
    list_of_next_follow = ''
    SnP_terms = ''
    text3 = ''
    Presentation_start = ''
    QnA_start = ''
    Presentation_section = ''
    QnA_section = ''
    Transcript_section_name = ''
    Sections = ''
    
    #del data[key]
    


data4 = {}

#### delete snp terms
SnP_words = ['spglobal', 'S&P Global', 'Copyright', 'EARNINGS CALL']
    
for key3 in data3:
    text4 = data3[key3]
    
    
    for i in SnP_words:
        text4 = '\n'.join(line for line in text4.split('\n') 
                          if i not in line)
             
    
    data4[key3] = text4
    
    text4 = ''
    
    
        
        
import pandas as pd
pd.set_option('max_colwidth',150)


data_df = pd.DataFrame.from_dict([data4]).transpose()
data_df.columns = ['transcript']
data_df = data_df.sort_index()
data_df

# data_df.transcript.loc['Presentation']



    
################ Cleaning transcripts   
### Start removing S&P company lines/copyright trademark


#### list of words that we want to remove
# Apply a first round of text cleaning techniques
import re
import string
import os


def clean_text_round1(text):
    '''Make text lowercase, remove text in square brackets, remove punctuation and remove words containing numbers.'''
    text = text.lower()
    text = re.sub('\[.*?\]', '', text)
    text = re.sub('[%s]' % re.escape(string.punctuation), '', text)
    text = re.sub('\w*\d\w*', '', text)
    return text

round1 = lambda x: clean_text_round1(x)

data_cleanrd1 = pd.DataFrame(data_df.transcript.apply(round1))

data_cleanrd1


# Apply a second round of cleaning
def clean_text_round2(text):
    '''Get rid of some additional punctuation and non-sensical text that was missed the first time around.'''
    text = re.sub('[‘’“”…]', '', text)
    text = re.sub('\n', ' ', text)
    return text

round2 = lambda x: clean_text_round2(x)

data_cleanrd2 = pd.DataFrame(data_cleanrd1.transcript.apply(round2))


#data_cleanrd2

data_forview = data_cleanrd2.transcript.loc['BOC Hong Kong (Holdings) Limited_Earnings Call_2020-08-31_EnglishPresentation']

print(data_forview)





################################
##### work in progress

from sklearn.feature_extraction.text import CountVectorizer

cv = CountVectorizer(stop_words='english')
data_cv = cv.fit_transform(data_cleanrd2.transcript)
data_dtm = pd.DataFrame(data_cv.toarray(), columns=cv.get_feature_names())
data_dtm.index = data_cleanrd2.index
data_dtm


#### data analysis
top_dict = {}
for c in data_dtm.columns:
    top = data_dtm[c].sort_values(ascending=False).head(30)
    top_dict[c]= list(zip(top.index, top.values))

top_dict

# Import the necessary modules for LDA with gensim
# Terminal / Anaconda Navigator: conda install -c conda-forge gensim
from gensim import matutils, models
import scipy.sparse


tdm = data_dtm.transpose()
tdm.head()


# We're going to put the term-document matrix into a new gensim format, from df --> sparse matrix --> gensim corpus
sparse_counts = scipy.sparse.csr_matrix(tdm)
corpus = matutils.Sparse2Corpus(sparse_counts)

id2word = dict((v, k) for k, v in cv.vocabulary_.items())

lda = models.LdaModel(corpus=corpus, id2word=id2word, num_topics=1, passes=1)

lda.print_topics()
    
    
    
    
    
    