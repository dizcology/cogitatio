from docx import *  #pip install python-docx
import re
import sys,os,time
from comtypes.client import CreateObject
import comtypes.gen
#import wave, contextlib
from Tkinter import Tk
from tkFileDialog import askopenfilename
from parseOSfilestaar import parseOSfile
from lessonstats import *
from subprocess import Popen
import datetime


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

lesson = re.findall('[0-9][0-9][0-9][A-Z]{1,3}',filename)[-1].encode('ascii')

filepath = '/'.join(filename.split('/')[:-1]) + '/'

paths = parseOSfile(filename)

allitems = []
#remove the step below if there is no need to list all the items (excluding items in Stage 4)
for fn in [f for f in os.listdir(filepath + 'Scripts/') if 'doc' in f and not '~' in f]:
    allitems += [fn.split('.doc')[0].encode('ascii')]

for path in ['weak + behind','weak + ontime']:
    allitems += [p for p in paths[path]] #[lesson + '-' + p for p in paths[path]]

for branch in paths['branches']:
    allitems += [i for b in branch for i in b]  #[lesson + '-' + i for b in branch for i in b] 
allitems = sorted(list(set(allitems)))

#print allitems

itemstats = {}
#itemnotes = {}

lessoncoefficients = { 
        'y-intercept': -217.2910,  #-250.49496,
        'Next count': 5.4153,  #6.78153,
        #'Submit count': -4.78795,
        #'Screen count': 0,
        'On-screen word count': 0.4506,  #0.37606,
        #'Words per screen':  0,
        'Table count': 43.3979,  #28.94253,
        #'Total number of words in tables': 1.19803,
        #'Average number of words per table': 3.63042,
        'Number of theory scripts': 32.9849,  #32.92053,
        'Number of problem scripts': 235.3690,   #242.86896,
        'Problem statement word count': 2.1272,  #1.90171,
        #'Solution word count': 0.08366,
        'Interactive steps count': -12.0940,  #-12.59084,
        #'Hints count': -6.61074,
        'Second layers count': 6.3491  #6.04622
        #'Solution input field count': 0
        }

def timeFormat(time):
    '''Format a time in seconds as mm:ss.'''
    return str(datetime.timedelta(seconds=time))[0:7]


def predLength(stats,coefs):
    '''Calculate a prediction from a set of coefficients for the given set of variables.'''
    prediction = coefs['y-intercept']
    prediction += sum([stats[f]*coefs[f] for f in coefs if f != 'y-intercept'])
    return prediction


def lessonStats(itemstats):
    '''Aggregate lesson item statistics for a given path through the lesson.'''
    lessonstats = {}
    for feat in lessoncoefficients:
        lessonstats[feat] = 0
        for i in itemstats:
            if feat in i:
                lessonstats[feat] += i[feat]
    return lessonstats

csvfilename = filepath + lesson + '_timing_' + datetime.datetime.now().strftime("%Y%m%d_%H%M") + '.csv'
warning = False
with open(csvfilename,'w') as csvfile:
    
    for i in sorted(allitems):  # <-- NOT NEEDED TO LOOP OVER ALL ITEMS!
        #item = '-'.join(i.split('-')[1:])
        itemfile = filepath + 'Scripts/' + i + '.docx' #path of lesson item

        if os.path.exists(itemfile.replace('docx','doc')) and not os.path.exists(itemfile):
            print >> sys.stderr, 'WARNING: script for item ' + i + ' is in *.doc format, not *.docx; skipping.'
            #itemstats[item] = {}
            itemstats[i] = {}
            csvfile.write(i + ',,(incorrect file format)\n')
        elif not os.path.exists(itemfile):
            print >> sys.stderr, 'WARNING: Scripts/' + i + '.docx not found; skipping.'
            itemstats[i] = {}
            csvfile.write(i + ',,(file not found)\n')
        else:  
            if 'th' in i or 'ex' in i: #theory or exercise lesson item
                itemstats[i] = theoryitemstats(itemfile)
            elif 'pr' in i: #problem lesson item
                itemstats[i] = probitemstats(itemfile)
    
    
    
    branchpath = []
    for branch in paths['branches']: #apply different functions depending on theory or problem
        branchpath += max(branch,key = lambda x: sum([predLength(itemstats[i],itemcoefficients) for i in x]))
        # This isn't strictly correct -- proper way would be to try all possible lesson paths for all
        # possible branch paths, since the lesson timing model is not the sum over items of the item
        # timing model.  In practice, though, this should be more than good enough, and it's much simpler
        # if there are multiple branch points in paths['branches'].

    wopathstats = lessonStats([itemstats[i] for i in (paths['weak + ontime'] + branchpath)]) #weak + ontime path
    #wbpathstats = [itemstats[i] for i in (paths['weak + behind'] + branchpath)]

    for feat in wopathstats:
		if feat != 'y-intercept':
			print feat + ':',
			csvfile.write(feat + ':' + ',')
			if isinstance(wopathstats[feat],int): 
				print wopathstats[feat]
				csvfile.write(str(wopathstats[feat]) + '\n')
			else: 
				print '{0:.2f}'.format(wopathstats[feat])
				csvfile.write('{0:.2f}'.format(wopathstats[feat]))
    
    print '\n\n'
    csvfile.write('\n\n')
	
#for key in wopathstats:
#    print key, wopathstats[key]
    
    csvfile.write('\ndescription,time,path\n')
	
    for path in ['weak + behind','weak + ontime']:
        pathstats = [itemstats[i] for i in (paths[path] + branchpath)]
        print path.ljust(15) + timeFormat(predLength(lessonStats(pathstats),lessoncoefficients)).rjust(10)
        #newtime[path[7:]]=timeFormat(predLength(lessonStats(pathstats),lessoncoefficients))
        csvfile.write(path + ',' + timeFormat(predLength(lessonStats(pathstats),lessoncoefficients)) + ',')
        csvfile.write('->'.join(sorted(paths[path]+branchpath)) + '\n')