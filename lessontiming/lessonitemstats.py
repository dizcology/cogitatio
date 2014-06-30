from docx import *
import os, sys
import re
from comtypes.client import CreateObject
import comtypes.gen
import wave, contextlib
from Tkinter import Tk
from tkFileDialog import askopenfilename
import tempfile

mainlinestyles = ['Line','Normal','DefaultStyle','Onscreen']

def removeBracketed(text):
    '''Remove text enclosed in square brackets.  Regexes can't really handle
    nested brackets, so this does it manually.'''
    bc = 0
    temp = ''
    for char in text:
        if char == '[':
            bc += 1
        if bc == 0: temp += char
        if char == ']':
            bc -= 1
    return temp

def getLength(text,wavfn):
    '''Get the length, in seconds, of a wav file produced by applying a
    text-to-speech engine to the given text.'''
    with tempfile.NamedTemporaryFile() as f:
        wavfn = f.name

    engine = CreateObject("SAPI.SpVoice")
    stream = CreateObject("SAPI.SpFileStream")
    stream.Open(wavfn, comtypes.gen.SpeechLib.SSFMCreateForWrite)
    engine.AudioOutputStream = stream
    engine.speak(text)
    stream.Close()

    with contextlib.closing(wave.open(wavfn,'r')) as f:
        frames = f.getnframes()
        rate = f.getframerate()
        duration = frames / float(rate)

    os.remove(wavfn)
    return duration

def getStats(text,style):
    '''Get various statistics of the text.'''
    submittime = 0
    wtdc = 0
    nextc = 0
    shortcount = 0
    medcount = 0
    longcount = 0
    nonstandardsubmittime = 0
    longsubmittime = 0

    if 'submit' in text.lower():
        if re.search('[0-9]+:[0-9][0-9]',text):
            time = re.search('[0-9]+:[0-9][0-9]',text).group(0)
            time = time.split(':')
            stime = int(time[0])*60+int(time[1])
            submittime += stime
            nonstandardsubmittime += stime
            if stime >= 180: longsubmittime += stime
        elif re.search('[0-9]+ second',text.lower()):
            time = re.search('[0-9]+ second',text.lower()).group(0)
            stime = int(time.split(' ')[0])
            submittime += stime
            nonstandardsubmittime += stime
            if stime >= 180:
                longsubmittime += stime
        elif re.search('[0-9]+ minute',text.lower()):
            time = re.search('[0-9]+ minute',text.lower()).group(0)
            stime = int(time.split(' ')[0])*60
            submittime += stime
            nonstandardsubmittime += stime
            if stime >= 180:
                longsubmittime += stime
        elif 'long' in text.lower():
            submittime += 40
            longcount += 1
        elif 'medium' in text.lower(): 
            submittime += 20
            medcount += 1
        elif 'short' in text.lower():
            submittime += 10
            shortcount += 1
    if 'wtd' in text.lower() and not 'disappears' in text.lower():
        wtdc += 1
    if '[next' in text.lower():
        nextc += 1
    return [submittime,wtdc,nextc,shortcount,medcount,longcount,nonstandardsubmittime,longsubmittime]

def getBranchText(text,style,inNR):
    '''Get text in the "main" and "NoResponse" branches.'''
    MLtext = ''
    NRtext = ''

    if (style in mainlinestyles or inNR) and re.match('^[A-Z][0-9]* ',text) and not '(tutor)' in text and not 'student)' in text:
        text = text.replace(u'\u2019',"'")
        text = text.encode('ascii','ignore')
        if style in ['Line','BranchLine'] or inNR:
            text = re.sub('^[A-Z][0-9]* ','',text)

            text = removeBracketed(text)

            text = text.replace('  ',' ')
            text = text.replace('#','')

            if style == 'Line':
                MLtext += ' ' + text
            elif inNR:
                NRtext += ' ' + text

    return MLtext,NRtext

def getDocText(text,style):
    '''Get any dialogue text.'''
    doctext = ''

    text = text.replace(u'\u2019',"'")
    text = text.encode('ascii','ignore')
    if re.match('^[A-Z][0-9]* ',text) and not '(tutor)' in text and not 'student)' in text:
        # Dialogue texts starts with a single letter (e.g. 'T' or 'A').
        # Exclude lines containing '(tutor)' and 'student)' as a precaution
        # against counting the character definition lines near the top.
        text = re.sub('^[A-Z][0-9]* ','',text).split('/ /')[0]

        text = removeBracketed(text)

        text = text.replace('  ',' ')
        text = text.replace('#','')

        doctext += ' ' + text
    return doctext

def getOnscreenText(text,style):
    '''Get any unbracketed next that's not dialogue.'''
    doctext = ''

    text = text.replace(u'\u2019',"'")
    text = text.encode('ascii','ignore')
    if not re.match('^[A-Z][0-9]* ',text):
        text = re.sub('^[A-Z][0-9]* ','',text).split('/ /')[0]

        text = removeBracketed(text)

        text = text.replace('  ',' ')
        text = text.replace('#','')

        doctext += ' ' + text
    return doctext

def getlessonitemstats(itemfn):
    '''Calculate statsitics of the lesson item.'''
    doc = Document(itemfn)
    wavfn = itemfn.replace('docx','wav')

    submittime = 0
    wtdc = 0
    nextc = 0
    inNR = False
    onscreentext = ''
    doctext = ''
    MLtext = ''
    NRtext = ''
    branchtext = ''
    avgbranchlength = 0
    totalbranchlength = 0
    branchcount = 0
    corrcount = 0
    gotoNR = True
    inBranch = False
    branchnum = 0
    shortcount = 0
    medcount = 0
    longcount = 0
    nonstandardsubmittime = 0.
    longcustomtime = 0.
    avgcorrcount = 0.

    for par in doc.paragraphs[7:]:
        style = par.style
        text = ''

        for run in par.runs:
            if not run.strike:
                text += ' ' + run.text
        text = re.sub('^ ','',text)

        # Track if we're in a No Response branch
        if style == 'NoResponse' or style == 'SecondaryNoResponse': inNR = True
        elif style in mainlinestyles: inNR = False

        # Track if we're in a branch besides No Response.
        if style in ['Correct','Incorrect']:
            inBranch = True
        elif inNR: inBranch = False

        if re.match('^correct',text.lower()) or re.match('^incorrect',text.lower()) or re.match('^no response',text.lower()):
            branchcount += 1
            if re.match('^correct',text.lower()):
                avgcorrcount += 1

        # Any dialogue that either we know is part of a branch or is explicitly NOT.
        if inBranch or inNR or not style in mainlinestyles:
            if re.match('^[A-Z][0-9]* ',text):
                btext = text.replace(u'\u2019',"'")
                btext = btext.encode('ascii','ignore')
                btext = re.sub('^[A-Z][0-9]* ','',btext)

                text = removeBracketed(text)

                btext = btext.replace('  ',' ')
                btext = btext.replace('#','')

                branchtext += ' ' + btext

                if gotoNR and inNR:
                    branchtext += ' ' + btext
            if 'go to nr' in text.lower() or 'give nr' in text.lower():
                gotoNR = True
        else:
            gotoNR = False
            if branchcount>0:
                branchnum += 1
                totalbranchlength = getLength(branchtext,itemfn.replace('.docx',"-branch" + str(branchnum) + ".wav"))
                avgbranchlength += totalbranchlength/branchcount
            branchtext = ''
            branchcount = 0


        temp = getStats(text,style)
        submittime += temp[0]
        wtdc += temp[1]
        nextc += temp[2]
        shortcount += temp[3]
        medcount += temp[4]
        longcount += temp[5]
        nonstandardsubmittime += temp[6]
        longcustomtime += temp[7]

        doctext += getDocText(text,style)
        onscreentext += getOnscreenText(text,style)

        temp = getBranchText(text,style,inNR)
        MLtext += temp[0]
        NRtext += temp[1]

    if branchnum > 0: avgcorrcount /= branchnum
    TTStime = getLength(doctext,wavfn)
    MLtime = getLength(MLtext,wavfn.replace('.wav','-main.wav'))
    NRtime = getLength(NRtext,wavfn.replace('.wav','-NR.wav'))

    return {
            'submit time': submittime, 
            'WTD count': wtdc,
            'next count': nextc,
            'dialogue time (total)': TTStime, 
            'dialogue time (main branch)': MLtime,
            'dialogue time (NR branch)': NRtime,
            'average branch time': avgbranchlength,
            'onscreen text word count': len(onscreentext.split()),
            'short count': shortcount,
            'medium count': medcount,
            'long count': longcount,
            'nonstandard submit time': nonstandardsubmittime,
            'long submit time': longcustomtime,
            'corrects per branch': avgcorrcount,
            'branch count': branchnum,
            }

if __name__ == '__main__':
    Tk().withdraw()
    if len(sys.argv) > 1:
        filename = sys.argv[1]
    else:
        filename = askopenfilename(**{'title':'Select the script'})

    if not filename[-4:] == 'docx':
        try:
            raise Exception()
        except Exception as e:
            print >> sys.stderr, 'OS file must be of type *.docx' 
            exit(3)

    stats = getlessonitemstats(filename)

    for feat in stats:
        print feat + ':',
        if isinstance(stats[feat],int): print stats[feat]
        else: print '{0:.2f}'.format(stats[feat])

    sys.stdin.readline()
