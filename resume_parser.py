from time import sleep
import win32com.client as win32
import glob,os
import sys

Files = []
extensions = ['*.doc','*.docx','*.rtf']
software = []


word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
for e in extensions:
    for infile in glob.glob( os.path.join('',e) ):
        Files.append(infile)
  
datascientist = []
D_map = {}
S_map = {}
      
write_file = open("outfile.txt", "w")
if ('Data_scientist keyterms.docx' in Files) & ('Software keyterms.docx' in Files):
    for infile in ['Data_scientist keyterms.docx','Software keyterms.docx']:
        m = {}
        doc = word.Documents.Open(os.getcwd()+'\\'+infile)
        for each_word in doc.Words:
            w = ""
            text = each_word.Text
            for i in text:
                if ((i>='a') & (i<='z')) |((i>='A') & (i<='Z')) | (i == '-') :
                    w+=i
                else:
                    break
            if w!="":
                m[w] = 1
        if infile == 'Software keyterms.docx':
            S_map = m
        else:
            D_map = m
        print "Done Parsing  ",infile
       # doc.Close(False)
    for infile in Files:
        m = {}

        if infile in ['Data_scientist keyterms.docx','Software keyterms.docx']:
            continue               
        else:
            choice = "Y"
            if choice == 'N':
                continue
            elif choice != 'Y':
                print 'Invalid input'
                print infile + ' is being identified'
            scount = 0
            mcount = 0
            doc = word.Documents.Open(os.getcwd()+'\\'+infile)
            for each_word in doc.Words:
                w = ""
                text = each_word.Text
                for i in text:
                    if ((i>='a') & (i<='z')) |((i>='A') & (i<='Z')) | (i == '-') :
                        w+=i
                    else:
                        break
                if w!="":
                    if w in m.keys():
                        m[w] += 1
                    else:
                        m[w] = 1
                        if w in S_map.keys():
                            #print('software',w)
                            scount += 1
                        if w in D_map.keys():
                            #print('datascientist',w)
                            mcount += 1
            print("Document Parsing Completed  ",infile)          
            spercent = int(float(scount)/float(len(S_map.keys()))*100)
            mpercent = int(float(mcount)/float(len(D_map.keys()))*100)
            write_file.write(str(infile)+"\t"+str(float(mpercent+spercent/2))+"%"+"\n")

else:
    print('Software keyterms.docx Or Data_scientist keyterms.docx Not Found')


print('Resume identified as software resume using "software keyterms.docx" are:')
for i in software:
    print(i)
print('Resume identified as Data Scientist resume using "Data_scientist keyterms.docx" are: ')
for i in datascientist:
    print(i)

