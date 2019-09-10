# -*- coding: utf-8 -*-
"""
Created on Fri Aug 30 06:40:06 2019

@author: samue
"""

import os
import string
import itertools
from glob import glob
from lxml import etree
from collections import Counter

#Define Lists here?
_corpusNames = []
_corpusTrees = []
_corpusElements = []
_corpusLyrics = [] #This is a list of lists: _corpusLyrics[0] contains a list of all the lyrics for _corpusTrees[0]


def _select(context, path):
    result = context.xpath(path)
    if len(result):
        return result[0]
    return None

def CollectCorpus(): #Fills _corpusNames & _corpusTrees
    for filename in glob(r'E:/Documents/UNE/MUSI366/Corpus\*.xml'):
        tree = etree.parse(filename)   
        _corpusTrees.append(tree)

def CollectLyrics():
    for filename in glob(r'E:/Documents/UNE/MUSI366/Corpus\*.xml'):
        tree = etree.parse(filename) 
        root = tree.getroot()
        _textList = []
        _syllabicList = []
        _wordParts = []
        _wordList = []
            
        for text in root.iter('text'):
            _textList.append(text.text)
                        
        for syllabic in root.iter('syllabic'):
            _syllabicList.append(syllabic.text)
        
        #Forms words from Syllables
        for text, syllabic in zip(_textList, _syllabicList):
            if(syllabic == 'begin'):
                _wordParts.append(text)
            elif(syllabic == 'middle'): 
                _wordParts.append(text)
            elif(syllabic == 'end'):
                _wordParts.append(text)
                _word = ''.join(_wordParts)
                _wordList.append(_word)
                _wordParts.clear()
            elif(syllabic == 'single'):
                _wordList.append(text)
                _wordParts.clear()
                #Adds completed words to _wordList
                
        #Converts all words to lowercase and removes punctuation
        _wordList = [''.join(character for character in word if character not in string.punctuation) for word in _wordList]
        _wordList = map(lambda x:x.lower(), _wordList)
        
        #Adds _wordList to _corpusLyrics
        _corpusLyrics.append(_wordList)
        
        #print (Counter(_wordList))    
        #print(_wordList)        
        #print('Text List = ', len(_textList), '~', 'Syllabic List = ', len(_syllabicList)) #Testing Output 
        #print(_textList[0], _syllabicList[0]) #Testing Output

def TextCount():
    _count = (list(itertools.chain.from_iterable(_corpusLyrics)))
    print('')
    print('Corpus total word count is: ', len(_count))
    print(Counter(_count))
         
CollectCorpus()
CollectLyrics()
TextCount()

#print(string.punctuation)
#print('Corpus Contains...')
#print('Corpus Trees', _corpusTrees)
#print('Total Corpus = ', len(_corpusTrees))
#print(len(_corpus))