#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Sep 25 14:00:54 2022

@author: Stephen
"""
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
from top2vec import Top2Vec
import transformers
import scipy
from transformers import AutoTokenizer, AutoModelForSequenceClassification
from scipy.special import softmax
from textblob import TextBlob



'''
transcripts_list = ['BOC Hong Kong (Holdings) Limited, H1 2019 Earnings Call, Aug 30, 2019',
                    'BOC Hong Kong (Holdings) Limited_Earnings Call_2019-08-30_English']

transcripts_list = ['BOC Hong Kong (Holdings) Limited, H1 2019 Earnings Call, Aug 30, 2019']

transcripts_list = ['BOC Hong Kong (Holdings) Limited_Earnings Call_2019-08-30_English']

'''

#### Name of the transcripts that we want to upload to python
transcripts_list = ['BOC Hong Kong (Holdings) Limited_Earnings Call_2013-03-26_English',
             'BOC Hong Kong (Holdings) Limited, 2016 Earnings Call, Mar 31, 2017',
             'BOC Hong Kong (Holdings) Limited, 2017 Earnings Call, Mar 29, 2018',
             'BOC Hong Kong (Holdings) Limited, 2018 Earnings Call, Mar 29, 2019',
             'BOC Hong Kong (Holdings) Limited, 2020 Earnings Call, Mar 30, 2021',
             'BOC Hong Kong (Holdings) Limited, 2021 Earnings Call, Mar 29, 2022',
             'BOC Hong Kong (Holdings) Limited, H1 2017 Earnings Call, Aug 30, 2017',
             'BOC Hong Kong (Holdings) Limited, H1 2018 Earnings Call, Aug 29, 2018',
             'BOC Hong Kong (Holdings) Limited, H1 2019 Earnings Call, Aug 30, 2019',
             'BOC Hong Kong (Holdings) Limited, H1 2021 Earnings Call, Aug 30, 2021',
             'BOC Hong Kong (Holdings) Limited, H1 2022 Earnings Call, Aug 30, 2022',
             'BOC Hong Kong (Holdings) Limited, Q2 2020 Earnings Call, Aug 31, 2020',
             'BOC Hong Kong (Holdings) Limited, Q4 2019 Earnings Call, Mar 27, 2020',
             'BOC Hong Kong Holdings Ltd., 2010 Earnings Call, Mar 24, 2011',
             'BOC Hong Kong Holdings Ltd., 2011 Earnings Call, Mar 29, 2012',
             'BOC Hong Kong Holdings Ltd., 2012 Earnings Call, Mar 26, 2013',
             'BOC Hong Kong Holdings Ltd., 2013 Earnings Call, Mar 26, 2014',
             'BOC Hong Kong Holdings Ltd., 2014 Earnings Call, Mar 25, 2015',
             'BOC Hong Kong Holdings Ltd., 2015 Earnings Call, Mar 30, 2016',
             'BOC Hong Kong Holdings Ltd., H1 2011 Earnings Call, Aug 24, 2011',
             'BOC Hong Kong Holdings Ltd., H1 2012 Earnings Call, Aug 23, 2012',
             'BOC Hong Kong Holdings Ltd., H1 2013 Earnings Call, Aug 29, 2013',
             'BOC Hong Kong Holdings Ltd., H1 2014 Earnings Call, Aug 19, 2014',
             'BOC Hong Kong Holdings Ltd., H1 2015 Earnings Call, Aug 28, 2015',
             'BOC Hong Kong Holdings Ltd., H1 2016 Earnings Call, Aug 30, 2016',
             'Hang Seng Bank Limited_Earnings Call_2010-08-02_English',
             'Hang Seng Bank Limited_Earnings Call_2011-08-01_English',
             'Hang Seng Bank Limited_Earnings Call_2012-02-27_English',
             'Hang Seng Bank Limited_Earnings Call_2012-07-30_English',
             'Hang Seng Bank Limited_Earnings Call_2013-03-04_English',
             'Hang Seng Bank Limited_Earnings Call_2013-08-05_English',
             'Hang Seng Bank Limited_Earnings Call_2014-02-24_English',
             'Hang Seng Bank Limited_Earnings Call_2014-08-04_English',
             'Hang Seng Bank Limited_Earnings Call_2015-02-23_English',
             'Hang Seng Bank Limited_Earnings Call_2016-02-22_English',
             'Hang Seng Bank Limited_Earnings Call_2016-08-03_English',
             'Hang Seng Bank Limited_Earnings Call_2017-02-21_English',
             'Hang Seng Bank Limited_Earnings Call_2017-07-31_English',
             'Hang Seng Bank Limited_Earnings Call_2018-02-20_English',
             'Hang Seng Bank Limited_Earnings Call_2018-08-06_English',
             'Hang Seng Bank Limited_Earnings Call_2019-02-19_English',
             'Hang Seng Bank Limited_Earnings Call_2019-08-05_English',
             'Hang Seng Bank Limited_Earnings Call_2020-02-18_English',
             'Hang Seng Bank Limited_Earnings Call_2020-08-03_English',
             'Hang Seng Bank Limited_Earnings Call_2021-02-23_English',
             'Hang Seng Bank Limited_Earnings Call_2021-08-02_English',
             'Hang Seng Bank Limited_Earnings Call_2022-02-22_English',
             'Hang Seng Bank Limited_Earnings Call_2022-08-01_English',
             'China Merchants Bank Co. Ltd._Earnings Call_2011-08-31_English',
             'China Merchants Bank Co. Ltd._Earnings Call_2012-03-30_English',
             'China Merchants Bank Co. Ltd._Earnings Call_2012-08-20_English',
             'China Merchants Bank Co. Ltd._Earnings Call_2013-04-02_English',
             'China Merchants Bank Co. Ltd._Earnings Call_2013-08-19_English',
             'China Merchants Bank Co. Ltd._Earnings Call_2014-03-31_English',
             'China Merchants Bank Co. Ltd._Earnings Call_2015-03-19_English',
             'China Merchants Bank Co. Ltd._Earnings Call_2016-03-31_English',
             'China Merchants Bank Co. Ltd._Earnings Call_2020-03-23_English',
             'China Merchants Bank Co. Ltd._Earnings Call_2020-08-31_English',
             'China Merchants Bank Co. Ltd._Earnings Call_2021-03-22_English',
             'China Merchants Bank Co. Ltd._Earnings Call_2021-04-26_English',
             'China Merchants Bank Co. Ltd._Earnings Call_2022-03-21_English',
             'China Merchants Bank Co. Ltd._Earnings Call_2022-08-22_English',
             'Industrial and Commercial Bank of China Limited_Earnings Call_2011-03-30_English',
             'Industrial and Commercial Bank of China Limited_Earnings Call_2011-08-25_English',
             'Industrial and Commercial Bank of China Limited_Earnings Call_2012-03-29_English',
             'Industrial and Commercial Bank of China Limited_Earnings Call_2012-08-30_English',
             'Industrial and Commercial Bank of China Limited_Earnings Call_2013-03-27_English',
             'Industrial and Commercial Bank of China Limited_Earnings Call_2013-04-26_English',
             'Industrial and Commercial Bank of China Limited_Earnings Call_2013-08-29_English',
             'Industrial and Commercial Bank of China Limited_Earnings Call_2014-03-27_English',
             'Industrial and Commercial Bank of China Limited_Earnings Call_2015-03-26_English',
             'Industrial and Commercial Bank of China Limited_Earnings Call_2015-10-30_English',
             'Industrial and Commercial Bank of China Limited_Earnings Call_2016-04-28_English',
             'Industrial and Commercial Bank of China Limited_Earnings Call_2017-08-30_English',
             'Industrial and Commercial Bank of China Limited_Earnings Call_2018-10-30_English',
             'Industrial and Commercial Bank of China Limited_Earnings Call_2019-04-29_English',
             'Industrial and Commercial Bank of China Limited_Earnings Call_2020-08-31_English',
             'Industrial and Commercial Bank of China Limited_Earnings Call_2020-10-30_English',
             'Industrial and Commercial Bank of China Limited_Earnings Call_2021-03-26_English',
             'Industrial and Commercial Bank of China Limited_Earnings Call_2021-08-27_English',
             'Industrial and Commercial Bank of China Limited_Earnings Call_2022-08-30_English',
             'Bank of China Limited_Earnings Call_2011-08-24_English',
             'Bank of China Limited_Earnings Call_2012-03-30_English',
             'Bank of China Limited_Earnings Call_2012-08-24_English',
             'Bank of China Limited_Earnings Call_2013-03-26_English',
             'Bank of China Limited_Earnings Call_2013-08-29_English',
             'Bank of China Limited_Earnings Call_2014-03-26_English',
             'Bank of China Limited_Earnings Call_2014-08-19_English',
             'Bank of China Limited_Earnings Call_2015-03-25_English',
             'Bank of China Limited_Earnings Call_2015-08-28_English',
             'Bank of China Limited_Earnings Call_2016-03-30_English',
             'Bank of China Limited_Earnings Call_2016-08-30_English',
             'Bank of China Limited_Earnings Call_2017-03-31_English',
             'Bank of China Limited_Earnings Call_2017-08-30_English',
             'Bank of China Limited_Earnings Call_2020-03-27_English',
             'Bank of China Limited_Earnings Call_2021-03-30_English',
             'Bank of China Limited_Earnings Call_2021-08-30_English',
             'Bank of China Limited_Earnings Call_2022-03-29_English',
             'Bank of China Limited_Earnings Call_2022-08-31_English',
             'AMMB Holdings Berhad, 2016 Earnings Call, May 27, 2016',
             'AMMB Holdings Berhad, 2017 Earnings Call, May 31, 2017',
             'AMMB Holdings Berhad, 2018 Earnings Call, May 31, 2018',
             'AMMB Holdings Berhad, H1 2017 Earnings Call, Nov 21, 2016',
             'AMMB Holdings Berhad, H1 2018 Earnings Call, Nov 28, 2017',
             'AMMB Holdings Berhad, H1 2019 Earnings Call, Nov 22, 2018',
             'AMMB Holdings Berhad, Nine Months 2017 Earnings Call, Feb 24, 2017',
             'AMMB Holdings Berhad, Nine Months 2018 Earnings Call, Feb 28, 2018',
             'AMMB Holdings Berhad, Nine Months 2019 Earnings Call, Feb 21, 2019',
             'AMMB Holdings Berhad, Q1 2017 Earnings Call, Aug 22, 2016',
             'AMMB Holdings Berhad, Q1 2018 Earnings Call, Aug 24, 2017',
             'AMMB Holdings Berhad, Q1 2019 Earnings Call, Aug 21, 2018',
             'AMMB Holdings Berhad, Q1 2020 Earnings Call, Aug 22, 2019']


#### define a function to read the word documents and name the newly uploaded files
def upload_transcripts(transcript):
    file = textract.process("/Users/Stephen/Desktop/python/text mining/transcripts/"
                            + transcript + ".docx")
    return(file)
'''

def upload_transcripts(transcript):
    file = textract.process("/Users/Stephen/Desktop/python/text mining/transcripts/testing/"
                            + transcript + ".docx")
    return(file)


'''


    
#### upload the files into a list
text = [upload_transcripts(transcript) for transcript in transcripts_list]

##text2 = [upload_transcripts_rtf(transcript) for transcript in transcripts_list_rtf]




### transform the company transcript list into dictionary
data = {}
for key in transcripts_list:
    for value in text:
        value2 = value.lower()
        data[key] = value2
        text.remove(value)
        value2=""
        break 


    
### for viewing   
data.keys()

### extract one transcript for inspection
#data['BOC Hong Kong (Holdings) Limited_Earnings Call_2014-03-26_English']
#data['BOC Hong Kong (Holdings) Limited_Earnings Call_2021-03-30_English']
#data['China Merchants Bank Co. Ltd._Earnings Call_2022-08-22_English']


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
    tablecontent_position = text2.find("contents")

    ### set the number of characters we want to look before and after the phrase table of content
    next_n_follow = 2000
    
    ### extract the +600 - -600 words around table of content
    list_of_next_follow = text2[tablecontent_position:tablecontent_position+next_n_follow]

    #### delete the boilerplate term of conditions from S&P
    SnP_terms = find_nth(text2, "these materials have been prepared solely for information", 1)
    SnP_terms2 = find_nth(text2, "the information in the transcripts", 1)
    
    ### create empty dataframe
    Presentation_section = ''
    QnA_section = ''
              
    
    if SnP_terms >= 0 and SnP_terms2 >= 0:
        
        if SnP_terms > SnP_terms2 :
            
            text3 = text2[:SnP_terms2]
        else:
            text3 = text2[:SnP_terms]
    
    elif SnP_terms >= 0:
        text3 = text2[:SnP_terms]
        
    elif SnP_terms2 >= 0:
        text3 = text2[:SnP_terms2]
    
    else: text3 = text2


    #### seperating the earning transcript into different sections
    #### Management presentation and Q&A
    
    #### Case 1: if there is presentation
    if "presentation" in list_of_next_follow:
    
        Presentation_start = find_nth(text2, "presentation", 2) 
    
        ### Case 1.1: also have QnA
        if "question and answer" in list_of_next_follow:
    
            #### create a new sublist
            QnA_start = find_nth(text2, "question and answer", 2)
            Presentation_section = text3[Presentation_start:QnA_start]
            QnA_section = text3[QnA_start:]
        
        ### Case 1.2: No QnA
        else:
            Presentation_section = text3[Presentation_start:]
    
    
    else:
        ### Case 2: Only have QnA
        if "question and answer" in list_of_next_follow:
            QnA_start = find_nth(text2, "question and answer", 2)
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
            data2[key + ' ' + 'question and answer'] = QnA_section
            
            data3 = {**data3, **data2}
        
            data2. clear()
    
    ### If have presentation
    else:
        
        ### Only presentation, no QnA
        if QnA_section == "":
            
            data2 = {}
            data2[key + ' ' + 'presentation'] = Presentation_section
            
            data3 = {**data3, **data2}
        
            data2. clear()
         
        ### Both presentation and QnA
        else:
            ### create names for the two parts
            Transcript_section_name = [key + ' ' + 'presentation', key + ' ' + 'question and answer']
    
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
    SnP_terms = ""
    SnP_terms2 = ""
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
SnP_words = ['spglobal', 's&p global', 'copyright', 'earnings call','www.spcapitaliq.com']
             
             
             
for key3 in data3:
    text4 = data3[key3]
    
    for i in SnP_words:
        text4 = '\n'.join(line for line in text4.split('\n') 
                          if i not in line)
            
    data4[key3] = text4
    text4 = ''
####################    

#print (data4.keys())

#print(data4['BOC Hong Kong (Holdings) Limited_Earnings Call_2020-08-31_English presentation'])


#data8 = {}
#data7 = data4['BOC Hong Kong (Holdings) Limited_Earnings Call_2020-08-31_English presentation']

#data7 = "adsdsdqweqwkjedjaskjdaskdjasdqwqwe\n\n\nbsdqsdsdsdsdsdsdsdsdwehjqwhe\n\ncdsdwdwdkqwueqweuussisi\n\n12312\n\n123456789010111213\n\n123"

#data8['BOC Hong Kong (Holdings) Limited_Earnings Call_2020-08-31_English presentation'] = data7



data5={}
data6={}

### problem \n\n and \n \n appears at the same time
### use the smaller number

#### loop through each items in the dictionary
for key in data4:
    
    iternation_num = 1
    
    ### extract an item from dictionary, one by one
    text = data4[key]
    
    #### text is coded as bytes, for easier viewing decode it into string
    #text2 = text.decode("utf-8") 
    text2 = text
    
    text4 = text2

    c=0    
    
    #b = find_nth(text4, "\n \n", 1) 
    #a = find_nth(text4, "\n\n", 1)

    #while find_nth(text4, "\n\n", 1) >= 0:
    while c >= 0:
        
        a = find_nth(text4, "\n\n", 1)
        b = find_nth(text4, "\n \n", 1) 

        if (a>=0 or b>=0) and b>a:
                
            ### a>0 and smaller than b
            if find_nth(text4, "\n\n", 1) >= 0:
    
                paragraph_location = find_nth(text4, "\n\n", 1)
                    
                text3 = text4[:paragraph_location]
                

                ### ignore new paragraphs if it is too short since it is most 
                ### likely just analyst names or companies name
                if len(text3) < 30:
                    ##### not sure if +4
                    text4 = text4[paragraph_location+2:]
                
                    c += 1
                
                else: 
                    data6[key + ' ' + str(iternation_num)] = text3
                
                    iternation_num = iternation_num +1
                    
                    text4 = text4[paragraph_location+2:]
        
                    #if find_nth(text4, "\n\n", 1) >= 0:
                    c += 1   
                
            ### a<0 but b>0
            else:

                paragraph_location = find_nth(text4, "\n \n", 1)
              
                text3 = text4[:paragraph_location]
                

                ### ignore new paragraphs if it is too short since it is most 
                ### likely just analyst names or companies name
                if len(text3) < 30:
                ##### not sure if +4
                    text4 = text4[paragraph_location+3:]
                
                    c += 1
                
                else: 
                    data6[key + ' ' + str(iternation_num)] = text3
                
                    iternation_num = iternation_num +1
                    
                    text4 = text4[paragraph_location+3:]
        
                    #if find_nth(text4, "\n\n", 1) >= 0:
                    c += 1  
                        
        elif (a>=0 or b>=0) and b<a:
                
            ## b>0 and smaller than a
            if find_nth(text4, "\n \n", 1) >= 0:
    
                paragraph_location = find_nth(text4, "\n \n", 1)
                    
                text3 = text4[:paragraph_location]
                

                ### ignore new paragraphs if it is too short since it is most 
                ### likely just analyst names or companies name
                if len(text3) < 30:
                    ##### not sure if +4
                    text4 = text4[paragraph_location+3:]
                
                    c += 1
                
                else: 
                    data6[key + ' ' + str(iternation_num)] = text3
                
                    iternation_num = iternation_num +1
                    
                    text4 = text4[paragraph_location+3:]
        
                    #if find_nth(text4, "\n\n", 1) >= 0:
                    c += 1   
                
            ### b<0 but a>0
            else:

                paragraph_location = find_nth(text4, "\n\n", 1)
              
                text3 = text4[:paragraph_location]
                

                ### ignore new paragraphs if it is too short since it is most 
                ### likely just analyst names or companies name
                if len(text3) < 30:
                    ##### not sure if +4
                    text4 = text4[paragraph_location+2:]
                
                    c += 1
                
                else: 
                    data6[key + ' ' + str(iternation_num)] = text3
                
                    iternation_num = iternation_num +1
                
                    text4 = text4[paragraph_location+2:]
    
                    #if find_nth(text4, "\n\n", 1) >= 0:
                    c += 1  
                        
        else:
            if len(text4) >30:
                    
                data6[key + ' ' + str(iternation_num)] = text4
                
                #iternation_num = iternation_num +1   
                    
            else:
                pass
            c = -1
                
        data5 = {**data5, **data6}
        data6.clear()
        text3=''
        a=''
        b=''
        
    text=''
    text2=''
    text4=''
    iternation_num=''
    c=''
                
        
 
###### Transform the dictionary to pandas dataframe        
pd.set_option('max_colwidth',1000)


data_df = pd.DataFrame.from_dict([data5]).transpose()
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
    text = re.sub('\[.*?\]', ' ', text)
    text = re.sub('[%s]' % re.escape(string.punctuation), '', text)
    text = re.sub('\w*\d\w*', ' ', text)
    return text

round1 = lambda x: clean_text_round1(x)

data_cleanrd1 = pd.DataFrame(data_df.transcript.apply(round1))

#data_cleanrd1


# Apply a second round of cleaning
def clean_text_round2(text):
    '''Get rid of some additional punctuation and non-sensical text that was missed the first time around.'''
    text = re.sub('[‘’“”…]', ' ', text)
    text = re.sub('\n', ' ', text)
    return text

round2 = lambda x: clean_text_round2(x)

data_cleanrd2 = pd.DataFrame(data_cleanrd1.transcript.apply(round2))


# Delete extra words
delete_word = [' hong ', ' kong ', ' gao ', 
             ' bank ',' banks ',' banking ',' cfo ',' china ', ' luo ', ' citigroup ', ' sachs ', ' lam ', ' icbc ',
             ' hsbc ', ' jiang ', ' analyst ', ' analysts ',' liu ',' tian ',
             ' andrew ',' mr ', ' morgan ', ' vice ', ' chief ', ' officer ', ' secretary ', ' office ', ' question ',
             ' answer ', ' thank ', ' hang ', " we ve ", ' goldman ', ' yi ', ' wang ',' ubs ',
             ' iii ', ' li ', ' ubs ', ' inc ', 'goldman ']


data_cleanrd3 = data_cleanrd2


#df = df.replace('old character','new character', regex=True)


for i in delete_word:    
    data_cleanrd3 = data_cleanrd3.replace(i, ' ', regex=True)
         

    


#### delete rows that are too short
data_cleanrd3 = data_cleanrd3[data_cleanrd2['transcript'].str.len()>100]
#data_cleanrd3 = data_cleanrd3.drop('index1', axis=1)


### extract only question and answer

data_cleanrd3_qna = data_cleanrd3
data_cleanrd3_qna['index1'] = data_cleanrd3_qna.index

data_cleanrd3_qna = data_cleanrd3_qna[data_cleanrd3_qna['index1'].str.contains('question')]
data_cleanrd3_qna = data_cleanrd3_qna.drop('index1', axis=1)


### extract only presentation
data_cleanrd3_presentation = data_cleanrd3
data_cleanrd3_presentation['index1'] = data_cleanrd3_presentation.index

data_cleanrd3_presentation = data_cleanrd3_presentation[data_cleanrd3_presentation['index1'].str.contains('presentation')]
data_cleanrd3_presentation = data_cleanrd3_presentation.drop('index1', axis=1)


### (sentiment)
### tranforming into list
sentiment_cleanrd1 = data_cleanrd3
#### work in progress

pol = lambda x: TextBlob(x).sentiment.polarity
sub = lambda x: TextBlob(x).sentiment.subjectivity

sentiment_cleanrd1['polarity'] = sentiment_cleanrd1['transcript'].apply(pol)
sentiment_cleanrd1['subjectivity'] = sentiment_cleanrd1['transcript'].apply(sub)


#sentiment_viewlist_1 = sentiment_cleanrd1[sentiment_cleanrd1['polarity'] < -0.2]
### combining the paragraphs back into a file for sentiment scoring - taking average
delete_last_word = lambda x: x.rsplit(' ', 1)[0]

sentiment_cleanrd1['index'] = sentiment_cleanrd1.index
sentiment_cleanrd1['index'] = sentiment_cleanrd1['index'].apply(delete_last_word)

sentiment_analysis2 = sentiment_cleanrd1[['index', 'polarity']]


sentiment_analysis3 = sentiment_analysis2.groupby(['index']).mean()
sentiment_analysis3['index'] = sentiment_analysis3.index



file_name = 'sentimentscorev1.xlsx'
sentiment_analysis3.to_excel(file_name)










#### try using top2vec
#from top2vec import Top2Vec

#docs = data_cleanrd3.transcript.tolist()
#docs = data_cleanrd3_qna.transcript.tolist()
#docs = data_cleanrd3_presentation.transcript.tolist()


#model = Top2Vec(docs)

model = Top2Vec(docs, embedding_model='universal-sentence-encoder',
                embedding_model_path='/Users/Stephen/Desktop/python/top2vec/universal-sentence-encoder_4')




topic_sizes, topic_nums = model.get_topic_sizes()

print(topic_sizes)

print(topic_nums)


#topic_words, words_score, topic_nums = model.get_topics(3)
topic_words, words_score, topic_nums = model.get_topics(10)




for words, score, num in zip(topic_words, words_score, topic_nums):
    print(num)
    print(f"Words: {words}")
    
topic_words, word_scores, topic_scores, topic_nums = model.search_topics(keywords=["profit"], num_topics=1)


topic_nums

topic_scores


for topic in topic_nums:
    model.generate_topic_wordcloud(topic)







###model = Top2Vec(docs)






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
    
    
    
    
    
    