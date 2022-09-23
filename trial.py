#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Sep 23 09:26:11 2022

@author: Stephen
"""
import sys
sys.modules[__name__].__dict__.clear()

import textract
text = textract.process("/Users/Stephen/Desktop/python/text mining/BOCHK.docx")

print(text)


#### text is coded as bytes, for easier viewing decode it into string
text2 = text.decode("utf-8") 


tablecontent_position = text2.find("Table of Contents")

next_n_follow = 600

print(text2[tablecontent_position:tablecontent_position+next_n_follow])

list_of_next_follow = text2[tablecontent_position:tablecontent_position+next_n_follow]

#### define a function to locate when a specific word appear the nth time
def find_nth(haystack, needle, n):
    start = haystack.find(needle)
    while start >= 0 and n > 1:
        start = haystack.find(needle, start+len(needle))
        n -= 1
    return start


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
    
    
    
    
    
    
    
    