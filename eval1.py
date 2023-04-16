import PyPDF2
from io import BytesIO
import pandas as pd

pdf = open('SE IT 2015.pdf','rb')
pdfReader = PyPDF2.PdfFileReader(pdf)

total = pdfReader.numPages

pageOne = pdfReader.getPage(1).extractText()

for i in range(0,total):
    page = pdfReader.getPage(i).extractText()
    j=1
    index = 393
    while j < 3:
        examSeatNo = page[index:page.find(" ",index)]
        details = {}
        details[examSeatNo] = {}
        details[examSeatNo]['examSeatNo'] = page[index:page.find(" ",index)]
        index+=12
        details[examSeatNo]['name'] = page[index:page.find("  ",index)]
        index+=40
        val=25
        while not(page[index+val].isdigit()):
            val+=1
        details[examSeatNo]['motherName'] = page[index:index + val].strip()
        index+=val
        #index+=25
        if page[index].isdigit():
            details[examSeatNo]['prn'] = page[index:page.find(" ",index)]
        else:
            tempIndex = page.find(" ",index) - 9
            details[examSeatNo]['prn'] = page[tempIndex:page.find(" ",index)]
            index+=2
        index+=292
        totalSem = 2
        if page[0:page.find("SEM.:2")] == -1:
            totalSem = 1
        k = 1
        while k <= totalSem:
            details[examSeatNo]['sem{}'.format(page[index:page.find(" ",index)])] = {}
            semNo = 'sem{}'.format(page[index:page.find(" ",index)])
            index+=5
            for m in range(1,11):
                subjectID = ""
                if page[index-1:index] == ' ':
                    subjectID = page[index:page.find(" ",index)]
                else:
                    subjectID = page[index-1:page.find(" ",index-1)]
                #subjectID = page[index:page.find(" ",index)]
                details[examSeatNo][semNo][subjectID] = {}
                details[examSeatNo][semNo][subjectID]['subjectId'] = subjectID
                index+=10
                findString = ""
                if page[index:index+1] == '0':
                    findString = "/"
                else:
                    findString = " "
                details[examSeatNo][semNo][subjectID]['insem'] = page[index:page.find(findString,index)]
                index+=9
                if page[index:index+1] == '0':
                    findString = "/"
                else:
                    findString = " "
                details[examSeatNo][semNo][subjectID]['theory'] = page[index:page.find(findString,index)]
                index+=18
                if page[index:index+1] == '0':
                    findString = "/"
                else:
                    findString = " "
                details[examSeatNo][semNo][subjectID]['termWork'] = page[index:page.find(findString,index)]
                index+=9
                if page[index:index+1] == '0':
                    findString = "/"
                else:
                    findString = " "
                details[examSeatNo][semNo][subjectID]['practical'] = page[index:page.find(findString,index)]
                index+=9
                if page[index:index+1] == '0':
                    findString = "/"
                else:
                    findString = " "
                details[examSeatNo][semNo][subjectID]['oral'] = page[index:page.find(findString,index)]
                index+=9
                details[examSeatNo][semNo][subjectID]['total'] = page[index:page.find(findString,index)]
                index+=5
                details[examSeatNo][semNo][subjectID]['credits'] = page[index:page.find(findString,index)]
                index+=5
                details[examSeatNo][semNo][subjectID]['grade'] = page[index:page.find(findString,index)]
                index+=5
                details[examSeatNo][semNo][subjectID]['gradePoints'] = page[index:page.find(findString,index)]
                index+=6
                details[examSeatNo][semNo][subjectID]['creditPonts'] = page[index:page.find(findString,index)]
                if m == 10:
                    index+=9
                else:
                    index+=5
            #index+=9
            k+=1
        #print(details)
        #print(pd.DataFrame(details))
        #index+=16
        #if page[index-1] == '-':
        details[examSeatNo]['pointer'] = page[page.find(":",index)+2:page.find(',',index)]
        index+=16
        if page[page.find(':',index) + 4] == 'R':
            details[examSeatNo]['totalCredits'] = page[page.find(':',index)+2:page.find('RESERVED',index)-7]
        elif page.find(':',index) + 4 == '.':
            details[examSeatNo]['totalCredits'] = page[page.find(':',index)+2:page.find('.',index)]
        else:
            details[examSeatNo]['totalCredits'] = page[page.find(':',index)+2:page.find('  ',index)]
        #else:
        #    details[examSeatNo]['pointer'] = page[index-4:index]

        index = page.find('DISTRIBUTION...........') + 24
        #index+=182

        print(details)
        j+=1
        #index+=2024




"""class Student:
    def __init__(examSeatNo, name):
        self.examSeatNo = examSeatNo
        self.name = name

class SubjectDetails:
    def __init__(self,subjectId,insem,theory,total,termWork,practical,oral,credit,grade,gradePoints,creditPoints):
        self.subjectId = subjectId
        self.insem = insem
        self.theory = theory
        self.total = total
        self.termWork = termWork
        self.practical = practical
        self.oral = oral
        self.credits = credits
        self.grade = grade
        self.gradePoints = gradePoints
        self.creditPonts = creditPoints

    def printDetails():
        print("{} {} {} {} {}".format(subjectId, insem,total, oral, grade))
        print(subjectId," ",insem)


class SubjectMarks:
    def extractMarks(self,pageOne,index):
        subCode = pageOne[index:pageOne.find(" ",index)]
        index += 10
        print(index)
        oeMarks = pageOne[index:pageOne.find("/",index)]
        index += 9
        print(index)
        thMarks = pageOne[index:pageOne.find("/",index)]
        index += 9
        print(index)
        totMarks = pageOne[index:pageOne.find("/",index)]
        index += 46
        print(index)
        tw = 0
        pr = 0
        oral = 0
        grd = pageOne[index:pageOne.find(" ",index)]
        index += 5
        print(index)
        grdPts = pageOne[index:pageOne.find(" ",index)]
        index += 6
        print(index)
        crdPts = pageOne[index:pageOne.find(" ",index)]
        return SubjectDetails(self,subCode,oeMarks,thMarks,totMarks,tw,pr,oral,grd,grdPts,crdPts)"""

"""index = 393
examSeatNo = pageOne[index:pageOne.find(" ",index)]
details = {}
details[examSeatNo] = {}
details[examSeatNo]['examSeatNo'] = pageOne[index:pageOne.find(" ",index)]
index+=8
details[examSeatNo]['name'] = pageOne[index:pageOne.find("  ",index)]
index+=65
details[examSeatNo]['prn'] = pageOne[index:pageOne.find("  ",index)]
index+=296
details[examSeatNo]['sem{}'.format(pageOne[index:pageOne.find(" ",index)])] = {}
"""
#print(details)

"""details = {
    pageOne[393:pageOne.find(" ",393)]: {
        'examSeatNo': pageOne[393:pageOne.find(" ",393)],
        'name': pageOne[405:pageOne.find("  ",405)],
        'prn': pageOne[470:pageOne.find(" ",470)],
        'sem{}'.format(pageOne[762:pageOne.find(" ",762)]): {
            pageOne[767:pageOne.find(" ",767)]: {
                'subjectId': pageOne[767:pageOne.find(" ",767)],
                'insem': pageOne[777:pageOne.find("/",777)],
                'theory': pageOne[786:pageOne.find("/",786)],
                'termWork': pageOne[802:pageOne.find(" ",802)],
                'practical': pageOne[811:pageOne.find(" ",811)],
                'oral': pageOne[820:pageOne.find(" ",820)],
                'total': pageOne[831:pageOne.find(" ",831)],
                'credits': pageOne[836:pageOne.find(" ",836)],
                'grade': pageOne[841:pageOne.find(" ",841)],
                'gradePoints': pageOne[846:pageOne.find(" ",846)],
                'creditPonts': pageOne[852:pageOne.find(" ",852)]
            }
        }
    }
}"""

#print(pd.DataFrame(details['student']['sem1']))
#print(details)

"""details = {'Student': {
        'examSeatNo':
        'name':
        'prn' :
        'sem1': {
            'subjectId':{
                'insem':
                'theory':
                'total':
                'termWork':
                'practical':
                'oral':
                'credit':
                'grade':
                'gradePoints':
                'creditPoints':
            }
        }
    }
}

SubjectMarks().extractMarks(pageOne,767).printDetails()"""


"""pdfContent = ""

def getPDFContent(path):
    content = ""
    num_pages = 10
    p = open(path,"rb")
    pdf = PyPDF2.PdfFileReader(p)
    for i in range(0,num_pages):
        content += pdf.getPage(i).extractText() + '\n'
    content = " ".join(content.replace(u"\xa0", " ").strip().split())
    return content

#pdfContent = BytesIO(getPDFContent("SE IT 2015.pdf").encode("ascii","ignore"))
for line in pdfContent:
    print(line.strip())"""

#get Exam Seat Number
"""examSeatNo = pageOne[393:pageOne.find(" ",393)]
print(examSeatNo)

#get Name
name = pageOne[405:pageOne.find("  ",405)]
print(name)

#get PRN
prn = pageOne[470:pageOne.find(" ",470)]
print(prn)

#get Sem Number
semNo = pageOne[762:pageOne.find(" ",762)]
print(semNo)


"""""" Marks Starts Now """"""

subCode = pageOne[767:pageOne.find(" ",767)]
print(subCode)

oeMarks = pageOne[777:pageOne.find("/",777)]
print(oeMarks)

thMarks = pageOne[786:pageOne.find("/",786)]
print(thMarks)

totMarks = pageOne[795:pageOne.find("/",795)]
print(totMarks)

grd = pageOne[841:pageOne.find(" ",841)]
print(grd)

grdPts = pageOne[846:pageOne.find(" ",846)]
print(grdPts)

grdPts = pageOne[852:pageOne.find(" ",852)]
print(grdPts)
print(pageOne.find("04"))"""

"""startIndex = 757
list = []
list.append(SubjectMarks(startIndex))
startIndex+=90
list.append(SubjectMarks(startIndex))
startIndex+=90
list.append(SubjectMarks(startIndex))
startIndex+=90
list.append(SubjectMarks(startIndex))
startIndex+=90
list.append(SubjectMarks(startIndex))
"""

#767
print(index)
print(pageOne.find("DISTRIBUTION..........."))
print(pdfReader.getPage(10).extractText())
