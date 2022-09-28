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
import pandas as pd
import os



#### Name of the transcripts that we want to upload to python
transcripts_list = ['BOC Hong Kong (Holdings) Limited_Earnings Call_2020-08-31_English',
             'BOC Hong Kong (Holdings) Limited_Earnings Call_2021-03-30_English',
             'BOC Hong Kong (Holdings) Limited_Earnings Call_2021-08-30_English',
             'BOC Hong Kong (Holdings) Limited_Earnings Call_2022-03-29_English']


#### define a function to read the word documents and name the newly uploaded files
def upload_transcripts(transcript):
    '''Returns transcript data specifically from scrapsfromtheloft.com.'''
    file = textract.process("/Users/Stephen/Desktop/python/text mining/transcripts/"
                            + transcript + ".docx")
    return(file)

    
#### upload the files into a list
text = [upload_transcripts(transcript) for transcript in transcripts_list]


### transform the company transcript list into dictionary
data = {}
for key in transcripts_list:
    for value in text:
        data[key] = value
        text.remove(value)
        break 
    
### for viewing   
data.keys()

### extract one transcript for inspection
data['BOC Hong Kong (Holdings) Limited_Earnings Call_2020-08-31_English']

### create empty dictionary for later use
data3 = {}


#### define a function to locate when a specific word appear the nth time
def find_nth(haystack, needle, n):
    start = haystack.find(needle)
    while start >= 0 and n > 1:
        start = haystack.find(needle, start+len(needle))
        n -= 1
    return start


#### seperate each transcripts into two sections: Presentation and Question and Answer

#### loop through each items in the dictionary
for key in data:
    
    ### extract an item from dictionary, one by one
    text = data[key]
    
    #### text is coded as bytes, for easier viewing decode it into string
    text2 = text.decode("utf-8") 

    ### locate the position of the phrase "Table of Contents" in order
    ### to check if the earning transcripts contains both presentation and QnA, or only one
    tablecontent_position = text2.find("Table of Contents")

    ### set the number of characters we want to look before and after the phrase table of content
    next_n_follow = 600
    
    ### extract the +600 - -600 words around table of content
    list_of_next_follow = text2[tablecontent_position:tablecontent_position+next_n_follow]

    #### delete the boilerplate term of conditions from S&P
    SnP_terms = find_nth(text2, "These materials have been prepared solely for information", 1)
    text3 = text2[:SnP_terms]


    #### seperating the earning transcript into different sections
    #### Management presentation and Q&A
    
    #### Case 1: if there is presentation
    if "Presentation" in list_of_next_follow:
    
        Presentation_start = find_nth(text2, "Presentation", 2) 
    
        ### Case 1.1: also have QnA
        if "Question and Answer" in list_of_next_follow:
    
            #### create a new sublist
            QnA_start = find_nth(text2, "Question and Answer", 2)
            Presentation_section = text3[Presentation_start:QnA_start]
            QnA_section = text3[QnA_start:]
        
        ### Case 1.2: No QnA
        else:
            Presentation_section = text3[Presentation_start:]
    
    
    else:
        ### Case 2: Only have QnA
        if "Question and Answer" in list_of_next_follow:
            QnA_start = find_nth(text2, "Question and Answer", 2)
            QnA_section = text3[QnA_start:]
    
        ### Case 3: Neither QnA and presentation
        else:    
            pass
        
    ### merging the extracted sections into a new dataframe (lefjoin concept)
    
    ### If no presentation
    if Presentation_section == "":
        
        ### Neither present and QnA
        if QnA_section == "":
            pass
        
        ### Only QnA
        else:
            data2 = {}
            data2[key + ' ' + 'Question and Answer'] = QnA_section
            
            data3 = {**data3, **data2}
        
            data2. clear()
    
    ### If have presentation
    else:
        
        ### Only presentation, no QnA
        if QnA_section == "":
            
            data2 = {}
            data2[key + ' ' + 'Presentation'] = Presentation_section
            
            data3 = {**data3, **data2}
        
            data2. clear()
         
        ### Both presentation and QnA
        else:
            ### create names for the two parts
            Transcript_section_name = [key + ' ' + 'Presentation', key + ' ' + 'Question and Answer']
    
            ### Group the two parts into a list to create a new dictionary
            Sections = [Presentation_section, QnA_section]
    
            ### Create a dictionary
            data2 = {}
            for key2 in Transcript_section_name:
                for value2 in Sections:
                    data2[key2] = value2
                    Sections.remove(value2)
                    break 
            
            ### Combine old and new dictionary
            data3 = {**data3, **data2}
            
            ### Delete items that has to be used in loop
            data2. clear()
    
    ### Delete items that has to be used in loop
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
    
    ### end of function
    

##### Deleting lists that contains the following SnP words
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
####################    

print (data4.keys())

print(data4['BOC Hong Kong (Holdings) Limited_Earnings Call_2020-08-31_English Presentation'])

    
        
###### Transform the dictionary to pandas dataframe        
pd.set_option('max_colwidth',150)


data_df = pd.DataFrame.from_dict([data4]).transpose()
data_df.columns = ['transcript']
data_df = data_df.sort_index()
### data_df

# data_df.transcript.loc['Presentation']



    
################ Cleaning transcripts   
### Start removing S&P company lines/copyright trademark


#### list of words that we want to remove
# Apply a first round of text cleaning techniques

###### define function to remove words 
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

#data_forview = data_cleanrd2.transcript.loc['BOC Hong Kong (Holdings) Limited_Earnings Call_2020-08-31_English Presentation']

#print(data_forview)

###### further cleaning using lemmatization


import nltk
from nltk.stem import PorterStemmer
#nltk.download()

w_tokenizer = nltk.tokenize.WhitespaceTokenizer()


ps =PorterStemmer()


def lemmatize_text(text):
    return [ps.stem(w) for w in w_tokenizer.tokenize(text)]

data_cleanrd3 = pd.DataFrame(data_cleanrd2.transcript.apply(lemmatize_text))




#### another method for lemmatization

# lemmatizer = nltk.stem.WordNetLemmatizer()

#def lemmatize_text(text):
#    return [lemmatizer.lemmatize(w) for w in w_tokenizer.tokenize(text)]

#data_cleanrd3 = data_cleanrd2.text.apply(lemmatize_text)








# Apply a second round of cleaning
def clean_text_round3(text):
    '''Lemmatization'''
    text2 = " ".join([lemma(word) for word in text.split()])
    return text2


round3 = lambda x: clean_text_round3(x)

data_cleanrd3 = pd.DataFrame(data_cleanrd2.transcript.apply(round3))



df.applymap(foo_bar)




















################################
##### work in progress

#### Document word matrix
from sklearn.feature_extraction.text import CountVectorizer

cv = CountVectorizer(stop_words='english')
data_cv = cv.fit_transform(data_cleanrd2.transcript)
data_dtm = pd.DataFrame(data_cv.toarray(), columns=cv.get_feature_names())
data_dtm.index = data_cleanrd2.index
data_dtm

#### data analysis

top_dict = {}
for c in data_dtm.columns:
    top = data_dtm[c].sort_values(ascending=False).head(1)
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
    
    
    
    
    
    