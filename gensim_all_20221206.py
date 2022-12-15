#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Dec  6 11:35:49 2022

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
import gensim
import pickle



# folder path
dir_path = "/Users/Stephen/Desktop/python/text mining/transcriptsv2/apacbankswithtranscripts/"

# list to store files
transcripts_list = []

# Iterate directory
for path in os.listdir(dir_path):
    # check if current path is a file
    if os.path.isfile(os.path.join(dir_path, path)):
        transcripts_list.append(path)
    path = ''    
    
        
def upload_transcripts(transcript):
    file = textract.process("/Users/Stephen/Desktop/python/text mining/transcriptsv2/apacbankswithtranscripts/"
                            + transcript)
    return(file)


text = ''

text = [upload_transcripts(transcript) for transcript in transcripts_list]


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


#data_df = pd.DataFrame.from_dict([data4]).transpose()


data_df = pd.DataFrame.from_dict([data5]).transpose()
data_df.columns = ['transcript']
data_df = data_df.sort_index()
#paragraph_df = data_df




#pickle the dataframe
with open('/Users/Stephen/Desktop/python/text mining', 'wb') as f:
    pickle.dump(data_df, f)
    
path = "/Users/Stephen/Desktop/python/text mining" 
pickle.dump (data_df, open(path + "transcript_paragraphv1.p", "wb") )  


paragraph_df = pickle.load(open(path + "transcript_paragraphv1.p", "rb"))

### cleaning using gensim
pharagraph_gensim_cleanv1 = paragraph_df.transcript.apply(gensim.utils.simple_preprocess)



### model gensim
model = gensim.models.Word2Vec(
    window=15,
    min_count=2,
    workers=4,
)


model.build_vocab(pharagraph_gensim_cleanv1, progress_per=1000)

model.epochs = 60



model.train(pharagraph_gensim_cleanv1, total_examples=model.corpus_count, epochs=model.epochs)


model.wv.most_similar("uncertainty", topn =35)

model.wv.most_similar("uncertain", topn =35)

model.wv.most_similar("profit", topn =30)

model.wv.most_similar("liquidity", topn =30)

















    
