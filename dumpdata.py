#!/usr/bin/python

import cPickle as pickle

datafile='pullsheets.pkl'

for hrline in pickle.load( open( datafile, "rb" ) ):
    print 'HR', hrline
    
