import os, sys, csv, re, glob
from docx import *
from theorystats import *
from problemstats import *


#returns timing statistics for an individual theory (or exercise) lesson item as a dictionary
#subroutines: from theorystats.py
def theoryitemstats(docpath):
    avgwordstable = 0 #set to 0
    avgcorrcount = 0 #set to 0
    wordsperscreen = 0 #set to 0
    stats = getStats(docpath)
    screencount = stats[0]
    nextcount = stats[1]
    submitcount = stats[2]
    branchcount = stats[3]
    correctcount = stats[4]
    onscreenwordcount = len(onscreenText(docpath).split())
    tablecount = getTablecount(docpath)
    totalwordsintable = getTablewordcount(docpath)
    if screencount == 0: 
        wordsperscreen = onscreenwordcount
        screencount = 1 #there is at least one screen?
    else:
        wordsperscreen = float(onscreenwordcount)/screencount
    if branchcount == 0:
            avgcorrcount = 0
    else:
        avgcorrcount = float(correctcount)/branchcount
    if tablecount == 0:
        avgwordstable = 0
    else:
        avgwordstable = float(totalwordsintable)/tablecount
    
    return{
            'Next count': nextcount,
            'Submit count': submitcount,
            'Screen count': screencount,
            'On-screen word count': onscreenwordcount,
            'Words per screen': wordsperscreen,
            'Table count': tablecount,
            'Total number of words in tables': totalwordsintable,
            'Average number of words per table': avgwordstable,
            'Number of theory scripts':1  #looking at only one theory script
            }

			
# returns statistics for a problem script in a lesson as a dictionary
#subroutines: from problemstats.py
def probitemstats(docpath): 
#with open('82AproblemData2.csv', 'w') as csvfile:
    #csvfile.write('Lesson, Grade, ProblemCount, SolutionCount, IScount, HintCount, \
    #SecondlayerCount \n')
    
    probcount = 0 #set at 0
    solcount = 0 #set at 0
    inputcount = 0 #set at 0
    stats = getSolstats(docpath)   
    iscount = stats[0]
    hintcount = stats[1]
    secondlayercount = stats[2]
    #inputcount = getInputcount(docpath)
    tempprobc = probCount(docpath)
    tempsolc = solCount(docpath)
    if tempprobc==-1: #problem extracting words from script
        probcount = 'ERROR!'
    else:
        probcount += tempprobc
    if tempsolc==-1: #problem extracting words from script
        solcount = 'ERROR!'
    else:
        solcount += tempsolc
    
    return{
            'Problem statement word count': probcount,
            'Solution word count': solcount,
            'Interactive steps count': iscount,
            'Hints count': hintcount,
            'Second layers count': secondlayercount,
            'Number of problem scripts': 1 #looking at 1 problem script
            }