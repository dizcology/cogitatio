from docx import *
import re
import sys,os,time
from comtypes.client import CreateObject
import comtypes.gen
import wave, contextlib
from Tkinter import Tk
from tkFileDialog import askopenfilename
from parseOSfile import parseOSfile
from lessonitemstats import getlessonitemstats
from subprocess import Popen

Tk().withdraw()
if len(sys.argv) > 1:
    filename = sys.argv[1]
else:
    filename = askopenfilename(**{'title':'Select the OS file'})

if not filename[-4:] == 'docx':
    try:
        raise Exception()
    except Exception as e:
        print >> sys.stderr, 'OS file must be of type *.docx' 
        exit(3)
lesson = re.findall('[0-9][0-9][0-9]',filename)[-1]
filepath = '/'.join(filename.split('/')[:-1]) + '/'

paths = parseOSfile(filename)

allitems = []
for fn in [f for f in os.listdir(filepath + 'Scripts/') if 'doc' in f and not '~' in f]:
    allitems += [fn.split('.doc')[0].encode('ascii')]

for path in ['weak + behind','weak + ontime']:
    allitems += [lesson + '-' + p for p in paths[path]]

for branch in paths['branches']:
    allitems += [lesson + '-' + i for b in branch for i in b] 
allitems = sorted(list(set(allitems)))

itemstats = {}

itemcoefficients = {
        'submit time': 0.302,
        'WTD count': 29.055,
        'next count': 5.602,
        'dialogue time (total)': 0.887,
        'dialogue time (main branch)': 0.443,
        'dialogue time (NR branch)': -0.344,
        'onscreen text word count': 0.114,
        'long submit time': -0.049,
        'corrects per branch': -3.133,
        'y-intercept': 20.293,
        'branch count': 1 
        }

lessoncoefficients = {
        'WTD count': 32.970,
        'next count': 3.004,
        'dialogue time (total)': 1.213,
        'onscreen text word count': -0.092,
        'medium count': 6.307,
        'nonstandard submit time': 0.290,
        'long submit time': -0.234,
        'corrects per branch': -72.396,
        'branch count': 0.,
        'total corrects': 0.,
        'y-intercept': 640.44
        }

def timeFormat(time):
    '''Format a time in seconds as mm:ss.'''
    minutes = int(time/60)
    seconds = int(round(time-minutes*60))
    return str(minutes) + ':' + str(seconds).zfill(2)

def predLength(stats,coefs):
    '''Calculate a prediction from a set of coefficients for the given set of variables.'''
    prediction = coefs['y-intercept']
    prediction += sum([stats[f]*coefs[f] for f in coefs if f != 'y-intercept'])
    return prediction

def lessonStats(itemstats):
    '''Aggregate lesson item statistics for a given path through the lesson.'''
    lessonstats = {}
    for i in itemstats:
        if 'corrects per branch' in i:
            i['total corrects'] = i['corrects per branch']*i['branch count']
    for feat in lessoncoefficients:
        lessonstats[feat] = 0
        for i in itemstats:
            if feat in i:
                lessonstats[feat] += i[feat]

    lessonstats['corrects per branch'] = lessonstats['total corrects']/lessonstats['branch count']

    return lessonstats

csvfilename = filepath + lesson + '_timing.csv'
warning = False
with open(csvfilename,'w') as csvfile:
    csvfile.write('item,time\n')
    for i in sorted(allitems):
        item = '-'.join(i.split('-')[1:])
        itemfile = filepath + 'Scripts/' + i + '.docx'

        if os.path.exists(itemfile.replace('docx','doc')) and not os.path.exists(itemfile):
            print >> sys.stderr, 'WARNING: script for item ' + item + ' is in *.doc format, not *.docx; skipping.'
            itemstats[item] = {}
            csvfile.write(i + ',,(incorrect file format)\n')
        elif not os.path.exists(itemfile):
            print >> sys.stderr, 'WARNING: Scripts/' + lesson + '-' + item + '.docx not found; skipping.'
            itemstats[item] = {}
            csvfile.write(i + ',,(file not found)\n')
        else:
            itemstats[item] = getlessonitemstats(itemfile)
            print i.ljust(15) + timeFormat(predLength(itemstats[item],itemcoefficients)).rjust(10)
            csvfile.write(i + ',' + timeFormat(predLength(itemstats[item],itemcoefficients)) + '\n')

    csvfile.write('\ndescription,time,path\n')

    branchpath = []
    for branch in paths['branches']:
        branchpath += max(branch,key = lambda x: sum([predLength(itemstats[i],itemcoefficients) for i in x]))
        # This isn't strictly correct -- proper way would be to try all possible lesson paths for all
        # possible branch paths, since the lesson timing model is not the sum over items of the item
        # timing model.  In practice, though, this should be more than good enough, and it's much simpler
        # if there are multiple branch points in paths['branches'].

    for path in ['weak + behind','weak + ontime']:
        pathstats = [itemstats[i] for i in (paths[path] + branchpath)]
        print path.ljust(15) + timeFormat(predLength(lessonStats(pathstats),lessoncoefficients)).rjust(10)
        csvfile.write(path + ',' + timeFormat(predLength(lessonStats(pathstats),lessoncoefficients)) + ',')
        csvfile.write('->'.join(sorted(paths[path]+branchpath)) + '\n')

Popen(csvfilename, shell=True)
