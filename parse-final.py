from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import nltk
import os
import shutil
import glob


path=r"D:\\resume\\resume\\" #this path contains the path which will hold all the resumes you hav downloaded
dest_path = r"D:\\WriteResume\\" #this path holds the path to the folder which will hold the resumes

def convertPDFToText(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()
    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)
    fp.close()
    device.close()
    string = retstr.getvalue()
    retstr.close()
    return string

skills=[]

print("Enter the skill set you want to extract from the candidate's CV")
print("Press 'END' when you are done")
print("Start entering the skills,press enter after each skill")
print("============================================================")

while True:
    skill=input()
    if skill.upper() == 'END':
        break
    else:
        skills.append(skill)

print(skills)
all_skills=""
for skill in skills:
    all_skills = all_skills+" "+skill

#process all files inside the resume corpus
def get_files(dir_path):
    files = []
    for file in os.listdir(dir_path):
        files.append(file)
    return files

def pre_process_cv(cv_text):
    cv_dict = {}
    for word in nltk.word_tokenize(cv_text):
        if word[0].isdigit():
            continue
        if not word[0].isalnum():
            continue
        cv_dict[word] = word
    return cv_dict

def squeeze_cv(cv_dict):
    main_cv = ""
    for key in cv_dict.keys():
        main_cv = main_cv + " " + key
    return main_cv

#start processing the strings one-by-one
#path_test=r"D:\\resume\\resume1\\Curriculum Vitae-Avik.docx"
#text = docxpy.process(path_test)
#print(text)

def parse_all_cv(dir_path):
    all_cv={}
    for resume in get_files(path):
        cv_dict={}
        squeezed=""
        file_path = path + resume
        read_pdf = convertPDFToText(file_path)  # contains the text from the pdf
        cv_dict = pre_process_cv(read_pdf)
        squeezed=squeeze_cv(cv_dict)
        all_cv[resume] = squeezed
    return all_cv

#rank all CVs according to the fuzzy-wuzzy library
from fuzzywuzzy import fuzz

def calc_scores(all_cv):
    max_score=len(skills)*100
    scored_cv={}
    score = 0
    temp_list=[]
    for resume in all_cv:
        score = 0
        for skill in skills:
            score = score + fuzz.token_set_ratio(all_cv[resume].lower(),skill.lower())
        scored_cv[resume] = (score/max_score)*100
    return scored_cv

scored_cv = calc_scores(parse_all_cv(path))
#sorted_cv = sorted(scored_cv.items(), key=operator.itemgetter(0),reverse=True)
def dict2tup(scored_cv):
    list_tuples=[]
    for key,value in scored_cv.items():
        list_tuples.append((key,value))
    return list_tuples

toSort = dict2tup(scored_cv)
from operator import itemgetter

def Reverse(tuples):
	new_tup = ()
	for k in reversed(tuples):
		new_tup = new_tup + (k,)
	return new_tup

sorted_dec = Reverse(sorted(toSort,key=itemgetter(1)))

def tup2dict(sorted_dec):
    ranked_cv={}
    for tup in sorted_dec:
        ranked_cv[tup[0]] = tup[1]
    return ranked_cv

ranked_cv = tup2dict(sorted_dec)
print("The resumes as ranked according to their employability")
print("==================================================================")
for cv in ranked_cv:
    print(cv,":",ranked_cv[cv]) #the final ranking of the cvs according to the skills
print("===================================================================")
#delete all the files inside a directory
def delete_files(dir_path):
    print("Deleting Files in folder...")
    files = glob.glob(dest_path+"*")
    for f in files:
        os.remove(f)

#enter the path where you want to write the resumes
def selectTopCv(dir_path):
    top = 0
    print("Please specify how many top resumes you want to extract...")
    top = int(input())
    print("=========================================")
    print("These are the top CVs")
    print("==========================================")
    files = glob.glob(dest_path+"*")
    if len(files) != 0:
        print("Do you want to flush the previous contents of the destination folder")
        print("Enter YES to delete and NO to keep")
        choice = input()
        if choice.lower() == "yes":
            print("Deleting Files...")
            delete_files(dir_path)
        else:
            print("Keeping the previously selected CV s")
    for i in range(top):
        shutil.copy(path+sorted_dec[i][0],dir_path+sorted_dec[i][0])
        print(sorted_dec[i][0])

selectTopCv(dest_path)
