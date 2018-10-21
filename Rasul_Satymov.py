#Due to the complexity and unfamiliar structure of HTML code...
#... I had to write some complex methods for parsing the CS and EE links.
#Though the IE parsing method is quite pretty :D


from Tkinter import *
import urllib2
from selenium import webdriver
from bs4 import BeautifulSoup
import time
import docclass
import os

class GradesClairvoyant(Frame):
    def __init__(self, root):
        Frame.__init__(self, root)
        root.title("Grades Clairvoyant")
        self.root = root
        self.courses_grades = {}
        self.initUI()

    def initUI(self):
        self.main_frame = Frame()
        self.main_frame.pack()
        Label(self.main_frame, text = "Grades Clairvoyant! Look into the future... but be careful!",
              bg="slategray4", fg='white', font='harrington 28 bold', width=45).grid(columnspan=3)
        Label(self.main_frame, text = "Please upload your curriculum with the grades: ", font='arial 13')\
            .grid(pady=10, stick=E)
        self.file_name = Label(self.main_frame, width=8) #A Label to display the department selected
        self.file_name.grid(row=1, column=1)
        Button(self.main_frame, text = "Browse", font='arial 12', width=15, command=self.browse)\
            .grid(row=1, column=2, stick=W)

        #Separator " " " " " " " " " " " " " " " " " " " " " " " " " " " " " "
        Label(self.main_frame, text=" \" \" \""*40).grid(row=2, columnspan=3)

        Label(self.main_frame, text="Enter URLs for course descriptions:", font="arial 16")\
            .grid(stick=W, columnspan=3, padx=20)
        self.urls = Text(self.main_frame, width=80, height=7, bg='gray80')
        self.urls.grid(columnspan=3, stick=W, padx=20)
        self.urls.insert(END, 'http://www.sehir.edu.tr/en/Pages/Academic/Bolum.aspx?BID=12\n'
                              'http://www.sehir.edu.tr/en/Pages/Academic/Bolum.aspx?BID=13\n'
                              'http://www.sehir.edu.tr/en/Pages/Academic/Bolum.aspx?BID=14\n'
                              'http://www.sehir.edu.tr/en/Pages/Academic/Bolum.aspx?BID=32')
                              #The URLs are there by defualt because we work with only these URLs anyway
        Label(self.main_frame, text="Key colors:", font="arial 16").grid(stick=W, padx=20)

        # A new frame for a nice UI
        self.new_frame = Frame(self.main_frame)
        self.new_frame.grid(columnspan=3, stick=W, padx=20)

        #Key colors
        Label(self.new_frame, text="A", bg="springgreen3", width=8, font='arial 16').grid(row=1, column=0, padx=5)
        Label(self.new_frame, text="B", bg="olivedrab1", width=8, font='arial 16').grid(row=1, column=1, padx=5)
        Label(self.new_frame, text="C", bg="khaki1", width=8, font='arial 16').grid(row=1, column=2, padx=5)
        Label(self.new_frame, text="D", bg="red2", width=8, font='arial 16').grid(row=1, column=3, padx=5)
        Label(self.new_frame, text="F", bg="black", fg='white', width=8, font='arial 16').grid(row=1, column=4, padx=5)

        #I placed this button in the main_frame for nice UI
        Button(self.main_frame, text="Predict Grades", width=15, font="arial 12", command=self.predicting_grades)\
            .grid(row=6, column=2, stick=W, pady=8)

        #Separator " " " " " " " " " " " " " " " " " " " " " " " " " " " " " "
        Label(self.main_frame, text=" \" \" \"" * 40).grid(row=7, columnspan=3)

    def browse(self):
        import tkFileDialog
        self.file_opt = options = {}
        options['filetypes'] = [('Excel', '.xlsx'), ('All Files', '.*')]
        options['initialdir'] = os.getcwd()
        options['parent'] = self.root
        self.browse_window = (str(tkFileDialog.askopenfilename(**self.file_opt)))
        self.file_name.config(text="<< "+self.browse_window[-7:-5].upper()+" >>",font='arial 13')#displaying the name of the file chosen

    #This method is called when "Predict Grades" button is pressed
    def predicting_grades(self):
        self.courses_grades.clear()  # I clear the courses dictionary everytime "Predict Courses" button is pressed to avoid duplication
        print self.courses_grades

        #I call this method for each semester. The last argument is the semester
        self.reading_excel(13,7,1,1)
        self.reading_excel(13,7,10,2)
        self.reading_excel(24,6,1,3)
        self.reading_excel(24,6,10,4)
        self.reading_excel(35,6,1,5)
        self.reading_excel(35,6,10,6)
        if self.browse_window[-7:-5] == "cs": #cause for some iexplicable reasons CS excel file is a little different
            self.reading_excel(46,5,1,7)
            self.reading_excel(46,5,10,8)
        else:
            self.reading_excel(45,5,1,7)
            self.reading_excel(45,5,10,8)

        self.reading_webpage()
        self.naive_bayes() #training the classifier
        self.classifying() #predicting grades

    def reading_excel(self, row, row_minus, column, semester):
        from xlrd import open_workbook
        excel_file = open_workbook(self.browse_window)
        sheet = excel_file.sheet_by_index(0)
        for row_index in range(row - row_minus, row):
            for col_index in range(column - 1, column):
                self.courses_grades.setdefault(sheet.cell(row_index, col_index).value, [])
                #Course name as a key and an empty list as a value
            for col_index1 in range((column + 6) - 1, column + 6):
                if len(sheet.cell(row_index, col_index1).value)==0: #If the course has no grade...
                    self.courses_grades[sheet.cell(row_index, col_index).value].append(semester) #I only append in which semester it is
                    continue
                self.courses_grades[sheet.cell(row_index, col_index).value].append(sheet.cell(row_index, col_index1).value[0]) #append the grade without the sign
                self.courses_grades[sheet.cell(row_index, col_index).value].append(semester) #append in which semester it is

    def reading_webpage(self):
        #check which file is chosen and call the appropriate parsing method
        if self.browse_window[-7:-5] == "cs":
            self.url = "http://www.sehir.edu.tr/en/Pages/Academic/Bolum.aspx?BID=12"
            self.cs_parser()
        elif self.browse_window[-7:-5] == "ee":
            self.url = "http://www.sehir.edu.tr/en/Pages/Academic/Bolum.aspx?BID=13"
            self.ee_parser()
        else:
            self.url = "http://www.sehir.edu.tr/en/Pages/Academic/Bolum.aspx?BID=14"
            self.ie_parser()
        #a separate method for parsing the Core courses which is called no matter which excel file is chosen
        self.getting_the_core_courses()

    def getting_the_core_courses(self):
        #no need for Selenium. urllib2 is enough
        c = urllib2.urlopen('http://www.sehir.edu.tr/en/Pages/Academic/Bolum.aspx?BID=32')
        soup = BeautifulSoup(c.read(), "html.parser")
        core_list = [] #core courses list
        for span in soup('span', {'class': "BolumOzet"}):
            for i in range(4, len(span('p'))):
                if len(self.gettextonly(span('p')[i]).strip()) > 0:
                    core_list.append(self.gettextonly(span('p')[i]).strip().split())
        del core_list[1] #getting rid of professor names and other abnormalities
        del core_list[3] #P.S. professors aren't abnormalities but their names' strings here aren't needed
        del core_list[5]
        del core_list[6][-2:]
        for k in range(9, 29, 2):
            del core_list[k]
        for core in range(0, len(core_list)):
            if core_list[core][0] == 'UNI':
                self.courses_grades.setdefault(core_list[core][0]+" "+core_list[core][1], []) #i.e. { "UNI 207" : [] }
                self.courses_grades[core_list[core][0]+" "+core_list[core][1]].append(" ".join(core_list[core + 1]))
                if core_list[core][2] == "/": #some courses are like: UNI 123 / UNI 124
                    self.courses_grades.setdefault(core_list[core][0]+" "+core_list[core][4], [])
                    self.courses_grades[core_list[core][0]+" "+core_list[core][4]].append(" ".join(core_list[core + 1]))
        self.courses_grades.setdefault(core_list[24][0][1:]+" "+core_list[24][1], []) #unusual course name. i.e. *UNI 111 / UNI 112
        self.courses_grades[core_list[24][0][1:]+" "+core_list[24][1]].append(" ".join(core_list[25]))
        self.courses_grades.setdefault(core_list[24][0][1:] + " " + core_list[24][4], [])
        self.courses_grades[core_list[24][0][1:] + " " + core_list[24][4]].append(" ".join(core_list[25]))
        for i in range(1, 12, 2):
            #the Turkish language courses are like: *UNI 115/116/215 etc.
            self.courses_grades.setdefault(core_list[26][0][1:]+" "+core_list[26][i], [])
            self.courses_grades[core_list[26][0][1:]+" "+core_list[26][i]].append(" ".join(core_list[27]))

    def naive_bayes(self):
        self.my_train = docclass.naivebayes(docclass.getwords)
        for key in self.courses_grades:
            if len(self.courses_grades[key]) == 3 and type(self.courses_grades[key][2])==unicode:
                self.my_train.train(self.courses_grades[key][2], self.courses_grades[key][0]) #course description as doc and course grade as class
                self.my_train.setthreshold(self.courses_grades[key][0], 1.0) #as mentioned in the assignment, I set the threshold to 1
        #The way my dict is build is like this:
        #    course_name : [course_grade, course's_semester, course_description]
        #or  course_name : [course's_semester, course_description]
        #and I only feed naivebayes with the courses that do have a grade
        #that is why I check if the length is equal to 3

        #I also check if the 3rd item is a unicode text
        # because some courses (i.e. IE XXX : [7,8,8]) may be in several semesters and therefore apply to my length checking condition ^^

    def classifying(self):
        #I add new elements to the UI after predicting the grades
        Label(self.main_frame, text='Predicted Grades', font='arial 16 bold').grid(row = 8, stick=W, padx=86)
        self.frame_for_preidctions = Frame(self.main_frame)
        self.frame_for_preidctions.grid(row = 9, columnspan=3, pady=10)
        self.scrollbar = Scrollbar(self.frame_for_preidctions)
        self.scrollbar.pack(side=RIGHT, fill=Y) #I used PACK here for a proper placement of the Scrollbar
        self.text_widget = Text(self.frame_for_preidctions, width=100, height=13, yscrollcommand=self.scrollbar.set)
        self.text_widget.pack(side=LEFT, fill=BOTH)
        self.scrollbar.config(command=self.text_widget.yview)

        courses_with_semester = {}
        uni_courses = []
        departmental_electives = []
        for key in self.courses_grades:
            #feed the classify method only with courses that have no grades
            if len(self.courses_grades[key]) == 1 and type(self.courses_grades[key][0]) == unicode:
                if str(key)[:3] == "UNI":
                    uni_courses.append(key + "-->" + self.my_train.classify(self.courses_grades[key][0]))
                else:
                    departmental_electives.append(key + "-->" + self.my_train.classify(self.courses_grades[key][0]))
            elif len(self.courses_grades[key]) == 2 and type(self.courses_grades[key][1]) == unicode:
                if str(key)[:3] == "UNI":
                    uni_courses.append(key + "-->" + self.my_train.classify(self.courses_grades[key][1]))
                else:
                    courses_with_semester.setdefault('Semester '+str(self.courses_grades[key][0]), [])
                    courses_with_semester['Semester '+str(self.courses_grades[key][0])]\
                        .append(key + "-->" + self.my_train.classify(self.courses_grades[key][1]))

        #tag_configure methods for colored highlighting of the courses
        self.text_widget.tag_configure('A', background='springgreen3')
        self.text_widget.tag_configure('B', background='olivedrab1')
        self.text_widget.tag_configure('C', background='khaki1')
        self.text_widget.tag_configure('D', background='red2')
        self.text_widget.tag_configure('F', background='black', foreground = "white")

        #Displaying the courses and predictions with colors in Text Widget
        for semester in courses_with_semester:
            self.text_widget.insert(END, "\n" + semester + "\n\n")
            for course in courses_with_semester[semester]:
                self.text_widget.insert(END, course+'\n', (course[-1], 'recent', 'warning'))
        self.text_widget.insert(END, "\n"+'UNI Courses'+"\n\n")
        for course in uni_courses:
            self.text_widget.insert(END, course+'\n', (course[-1], 'recent', 'warning'))
        self.text_widget.insert(END, "\n"+"Departmental Electives"+"\n\n")
        for course in departmental_electives:
            self.text_widget.insert(END, course+"\n", (course[-1], 'recent', 'warning'))

    def cs_parser(self):
        driver = webdriver.Firefox()
        driver.get(self.url)
        button = driver.find_element_by_link_text("Course Descriptions")
        button.click()
        time.sleep(3) #Although it was recommended to set 5 seconds, I figured 3 seconds are more than enough
        html = driver.page_source
        driver.close()
        soup = BeautifulSoup(html, "html.parser")
        course_description_draft = []  # incomplete list of courses and their descriptions
        for i in soup('p', {'style': 'font-family:helvetica;font-size:9pt;color:rgb(0, 0, 0)'}):
            if i.span == None:
                pass
            else:
                for j in i.span:
                    course_description_draft.append(self.gettextonly(j).strip(' \n'))
        course_description = []  # nice list of courses and their descriptions
        for item in course_description_draft:
            if len(item) > 2:
                if item[:8] != 'Textbook' and not re.match('[0-9]', item[-4:]): #To remove the Textbook
                    if not re.match('[0-9]', item[:3]):
                        if len((re.sub(r'^https?:\/\/.*[\r\n]*', '', item, flags=re.MULTILINE).strip())) > 2:
                            course_description.append(
                                re.sub(r'^https?:\/\/.*[\r\n]*', '', item, flags=re.MULTILINE).strip().split())#to remove the links in course descriptions
        for k in range(0, len(course_description), 2):
            if re.match('[A-Z]', course_description[k][0]):
                self.courses_grades.setdefault(course_description[k][0] + ' ' + course_description[k][1], [])
                self.courses_grades[course_description[k][0] + ' ' + course_description[k][1]].append(
                    " ".join(course_description[k + 1]))

    def ee_parser(self):
        driver = webdriver.Firefox()
        driver.get(self.url)
        button = driver.find_element_by_link_text("Course Descriptions")
        button.click()
        time.sleep(3) #Although it was recommended to set 5 seconds, I figured 3 seconds are more than enough
        html = driver.page_source
        driver.close()
        soup = BeautifulSoup(html, "html.parser")
        ee_courses = []
        for i in soup('div', {'class':'ExternalClass40EE2489EF7943DEA08AF13FA56C59C7'}):
            for k in range(0, len(i('div'))):
                #to remove the Textbooks, links, prerequisite courses' info
                if self.gettextonly(i('div')[k]).strip()[:8] != 'Textbook' \
                        and self.gettextonly(i('div')[k]).strip()[:4] != 'http'\
                        and self.gettextonly(i('div')[k]).strip()[:9] != 'Text Book'\
                        and self.gettextonly(i('div')[k]).strip()[:13]!='(Prerequisite':
                    ee_courses.append(self.gettextonly(i('div')[k]).strip().split())

        for k in range(0, len(ee_courses)-1):
            if re.match('[0-9]', ee_courses[k][1]):
                self.courses_grades.setdefault(ee_courses[k][0]+" "+ee_courses[k][1], [])
                self.courses_grades[ee_courses[k][0]+" "+ee_courses[k][1]].append(" ".join(ee_courses[k+1]))

    def ie_parser(self):
        driver = webdriver.Firefox()
        driver.get(self.url)
        button = driver.find_element_by_link_text("Course Descriptions")
        button.click()
        time.sleep(3) #Although it was recommended to set 5 seconds, I figured 3 seconds are more than enough
        html = driver.page_source
        driver.close()
        soup = BeautifulSoup(html, "html.parser")
        ie_courses_draft = []
        for i in soup('div', {'class': 'ExternalClass66C86DAFFE5B42FE91CEE3CD4790CA72'}):
            #There were no links in IE courses' descriptions
            # and I figured that it won't make much difference if I leave the textbooks
            for linebreak in i('br'):
                linebreak.decompose()
            for k in i('div'):
                ie_courses_draft.append(k.strong.extract().get_text().split())
                ie_courses_draft.append(k.get_text().split())

        for item in range(len(ie_courses_draft)):
            if len(ie_courses_draft[item])>1:
                if re.match('[0-9]',ie_courses_draft[item][1]):
                    if ie_courses_draft[item][1] == '498':
                        continue
                    self.courses_grades.setdefault(ie_courses_draft[item][0]+" "+ie_courses_draft[item][1], [])
                    self.courses_grades[ie_courses_draft[item][0]+" "+ie_courses_draft[item][1]].append(" ".join(ie_courses_draft[item+1]))

    def gettextonly(self, soup):
        v = soup.string
        if v == None:
            c = soup.contents
            resulttext = ''
            for t in c:
                subtext = self.gettextonly(t)
                resulttext += subtext + '\n'
            return resulttext
        else:
            return v.strip()

if __name__ == '__main__':
    root = Tk()
    GradesClairvoyant(root)
    root.mainloop()
