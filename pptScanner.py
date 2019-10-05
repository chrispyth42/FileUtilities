#!/usr/bin/python3
#To read the archive that is a ppt file
from zipfile import ZipFile
#For easy string matching/replacing
import re
#For getting a directory
import os
import sys

#file path delimiter for windows/linux
delim = '\\' #or '/'

#Iterate through every slide xml contained in a powerpoint, and notify user of
#any instances of the search terms found
def scanPpt(searchTerm,fileName):
    #Open power point archive and get directory structure
    ppt = ZipFile(fileName,'r')
    ppDir = ppt.namelist()

    #Variable that matched slide numbers are returned to
    matchedSlides = list()

    #For each file that is in the slides directory, search its contents
    for f in ppDir:
        #Filters out every file that isn't a slide XML file
        if re.match(r'ppt/slides/[^/]+\.xml',f):
            #Read the ppt slide file from the ppt archive, and decode it to be a normal string
            slide = ppt.read(f).decode()
            #Remove all XML tags from the file
            slide = re.sub(r'</?[^>]+>','',slide)

            #If the search string is in the slide, add to list of successful slides
            if searchTerm.upper() in slide.upper():
                matchedSlides.append(f.split('/slide')[-1][:-4])
    
    #Return results
    if matchedSlides == list():
        return "\t...\t" + fileName.split('/')[-1]
    else:
        return "\t(" + ",".join(matchedSlides) + ")\t" + fileName.split('/')[-1] 

#Passes every ppt in a directory to the scanPpt function
def scanPptDirectory(searchTerm,path):
    #Get list of all ppts in the directory
    ppts = [f for f in os.listdir(path) if f.endswith(".pptx")]

    #Print the current path to serve as a header
    print(path)

    #Pass each file to the scanner function
    for ppt in ppts:
        print(scanPpt(searchTerm, (path + delim + ppt)))
    print('')

#Gets all subdirectories in a directory, and passes them up to scanPptDirectory, to scan each powerpoint within it
def scanPptDirectoryTree(searchTerm,path):
    #First do ppts in the root path provided
    scanPptDirectory(searchTerm,path)

    #Then do ppts in all child paths of that one
    for root, dirs, files in os.walk(path):
        for directory in dirs:
            scanPptDirectory(searchTerm,(root + delim + directory).replace(delim*2,delim)) #Replaces '\\' with '\'

#Passes command line arguments into the above functions
#syntax is
#   scanner.py [searchTerm] [path] [recursive?]
#if the recursive argument exists at all, it searches all powerpoints in a directory tree
def main():
    if len(sys.argv) > 2:
        searchTerm = sys.argv[1]
        path = sys.argv[2]
        if len(sys.argv) > 3:
            scanPptDirectoryTree(searchTerm,path)
        else:
            scanPptDirectory(searchTerm,path)

main()
