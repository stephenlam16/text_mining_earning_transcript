#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Nov 28 11:07:39 2022

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
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    







