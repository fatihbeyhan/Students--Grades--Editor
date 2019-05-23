from Tkinter import *
from ttk import *
import ttk
import tkFileDialog
from xlrd import open_workbook
from tempfile import TemporaryFile
from xlwt import Workbook


class Student: # Creating a class to represent each student with their informations

    def __init__(self,id,name,surname,section,dept,gpa,mp1,mp2,mp3,mt,fe):

        self.id = id
        self.name = name
        self.surname = surname
        self.section = section
        self.dept = dept
        self.gpa = gpa
        self.mp1 = mp1
        self.mp2 = mp2
        self.mp3 = mp3
        self.mt = mt
        self.fe = fe

class Student_Data: # Creating class to create objects for each student and other data that we need to use for further steps.

    StudentDict = {} # Creating a dictionary to keep data of students
    studentList = [] # Creating a student list by their IDs to use in other fucntions
    studentObjectList = [] # Creating a student object list to use in other functions

    def SelectFile(self,CheckFile): # Creating a function to call when "Select File" button clicked
        CheckFile.set("Yes")


        file_diractory = tkFileDialog.askopenfilename()
        book = open_workbook(file_diractory)

        sheet = book.sheet_by_index(0)

        for o in range(1, sheet.nrows):                                                   # Filling the student list + Creating an object for each student with their atributes + Filling the Student Object List
                                                                                          #
            namesplitlist = unicode.split(sheet.cell_value(o, 1), " ")                    #
            if len(namesplitlist) == 2:                                                   #
                nname = namesplitlist[0]                                                  # Filling the student list + Creating an object for each student with their atributes + Filling the Student Object List
                ssurname = namesplitlist[1]                                               #
            if len(namesplitlist) == 3:                                                   #
                nname = namesplitlist[0] +" "+ namesplitlist[1]                           #
                ssurname = namesplitlist[2]                                               #
                                                                                          # Filling the student list + Creating an object for each student with their atributes + Filling the Student Object List
            b = str(int(sheet.cell_value(o, 0)))                                          #
                                                                                          #
            Student_Data.studentList.append(b)                                            #
                                                                                          # Filling the student list + Creating an object for each student with their atributes + Filling the Student Object List
            b = Student(b,nname,ssurname,sheet.cell_value(o, 2),sheet.cell_value(o, 3),   #
            sheet.cell_value(o, 4),                                                       #
            sheet.cell_value(o, 5),                                                       #
            sheet.cell_value(o, 6),                                                       # Filling the student list + Creating an object for each student with their atributes + Filling the Student Object List
            sheet.cell_value(o, 7),                                                       #
            sheet.cell_value(o, 8),                                                       #
            sheet.cell_value(o, 9))                                                       #
                                                                                          # Filling the student list + Creating an object for each student with their atributes + Filling the Student Object List
                                                                                          #
                                                                                          #
                                                                                          #
            Student_Data.studentObjectList.append(b)                                      # Filling the student list + Creating an object for each student with their atributes + Filling the Student Object List




        for roww in range(1, sheet.nrows):                                                                  # Creating the keys as students' IDs and giving all other information as value.
            Student_Data.StudentDict[int(sheet.cell_value(roww, 0))] = [sheet.cell_value(roww, 5),
                                                                     sheet.cell_value(roww, 6),
                                                                     sheet.cell_value(roww, 7),
                                                                     sheet.cell_value(roww, 8),
                                                                     sheet.cell_value(roww, 9),             # Creating the keys as students' IDs and giving all other information as value.
                                                                     sheet.cell_value(roww, 0),
                                                                     sheet.cell_value(roww, 1),
                                                                     sheet.cell_value(roww, 2),
                                                                     sheet.cell_value(roww, 3),
                                                                     sheet.cell_value(roww, 4)]              # Creating the keys as students' IDs and giving all other information as value.



class GMT_Interface(Frame):

    def __init__(self,parent):
        Frame.__init__(self,parent)
        self.parent = parent
        self.initUI()


        self.CheckIfFileLoaded = StringVar()       # Creating the StringVar to check if file loaded or not in the further steps
        self.CheckIfFileLoaded.set("No")
        self.TreeViewIsClicked = StringVar()       # Creating the StringVar to check if they click on treeview or not
        self.TreeViewIsClicked.set("No")



    def initUI(self):

        self.bigFrame = Frame(self, borderwidth=12)
        self.infoframe = Frame(self.bigFrame, relief=RIDGE,borderwidth=2)
        self.editFrame = Frame(self,relief=RIDGE,borderwidth=12)

        self.InfoInsideFrame2 = Frame(self.infoframe)
        self.InfoInsideFrame3 = Frame(self.infoframe)
        self.InfoInsideFrame4 = Frame(self.infoframe)
        self.InfoInsideFrame5 = Frame(self.infoframe)
        self.InfoInsideFrame6 = Frame(self.infoframe)

        self.header = Label(self,text = "Grades Management Tool",anchor = CENTER,font=("Calibri",25, "bold"),background="springgreen",foreground='white')

        self.selectFileB = Button(self,text="Select File",command= self.SelectFileC)

        self.treeview = ttk.Treeview(self.bigFrame)
        self.treeview['columns'] = ('ID', 'NAME','SURNAME')
        self.treeview.heading('ID', text='ID', anchor=CENTER)
        self.treeview.heading('NAME', text='NAME', anchor=CENTER)
        self.treeview.heading('SURNAME', text='SURNAME', anchor=CENTER)


        self.treeview.column('#00', anchor=W, minwidth=00, stretch=0, width=0)

        self.treeview.column('#01', anchor=W, minwidth=50, stretch=1, width=103)
        self.treeview.column('#02', anchor=W, minwidth=50, stretch=1, width=120)
        self.treeview.column('#03', anchor=W, minwidth=50, stretch=1, width=120)
        self.treeview.bind("<ButtonPress-1>", self.TreeviewClick)


        self.showDataB = Button(self.bigFrame, text="--Show Data-->",command= self.ShowDataC)

        #INFOLABELS
        self.DataLabel = Label(self.infoframe,text="Student Details:",background = 'darkcyan',foreground="white",anchor=CENTER,font= ("Arial",12,"bold"))

        self.Name = Label(self.InfoInsideFrame2,text ="Name :",anchor=W,width=9,font= ("Arial",10,"bold"))
        self.ID = Label(self.InfoInsideFrame4, text="ID:",anchor=W,width=9,font= ("Arial",10,"bold"))
        self.Surname = Label(self.InfoInsideFrame3, text="Surname:",anchor=W,width=9,font= ("Arial",10,"bold"))
        self.Dept = Label(self.InfoInsideFrame5, text="Dept:",anchor=W,width=9,font= ("Arial",10,"bold"))
        self.GPA = Label(self.InfoInsideFrame6, text= "GPA:",anchor=W,width=9,font= ("Arial",10,"bold"))
        self.MP1 = Label(self.InfoInsideFrame2, text="MP1 Grade:",anchor=W,width=12,font= ("Arial",10,"bold"))
        self.MP3 = Label(self.InfoInsideFrame4, text="MP3 Grade:",anchor=W,width=12,font= ("Arial",10,"bold"))
        self.MP2 = Label(self.InfoInsideFrame3, text="MP2 Grade:",anchor=W,width=12,font= ("Arial",10,"bold"))
        self.MT = Label(self.InfoInsideFrame5, text="MT Grade:",anchor=W,width=12,font= ("Arial",10,"bold"))
        self.Final = Label(self.InfoInsideFrame6, text="Final Grade:",anchor=W,width=12,font= ("Arial",10,"bold"))

        self.NameL = Label(self.InfoInsideFrame2,anchor=W,width=13)
        self.IDL = Label(self.InfoInsideFrame4,width=13)
        self.SurnameL = Label(self.InfoInsideFrame3,width=13)
        self.DeptL = Label(self.InfoInsideFrame5,width=13)
        self.GPAL = Label(self.InfoInsideFrame6,width=13)
        self.MP1L = Label(self.InfoInsideFrame2,anchor=W,width=10)
        self.MP2L = Label(self.InfoInsideFrame3,anchor=W,width=10)
        self.MP3L = Label(self.InfoInsideFrame4,anchor=W,width=10)
        self.MTL = Label(self.InfoInsideFrame5,anchor=W,width=10)
        self.FinalL = Label(self.InfoInsideFrame6,anchor=W,width=10)


        #EDITING GRADES

        self.projectN = Label(self.editFrame, text= "Projects :")

        self.Grades = Label(self.editFrame,text="Grades :")

        self.mp1L = Label(self.editFrame,text ="MP1")
        self.mp1E = Entry(self.editFrame,width=12)

        self.mp2L = Label(self.editFrame, text="MP2")
        self.mp2E = Entry(self.editFrame,width=12)

        self.mp3L = Label(self.editFrame, text="MP3")
        self.mp3E = Entry(self.editFrame,width=12)

        self.mtL = Label(self.editFrame, text="MT")
        self.mtE = Entry(self.editFrame,width=12)


        self.feL = Label(self.editFrame, text="Final")
        self.feE = Entry(self.editFrame,width=12)

        self.saveGradesB = Button(self.editFrame,text="Save Grades",command= self.SaveGradesC,width=15)
        self.ExportAs = Label(self.editFrame,text="Export As:")

        self.FileTypeVar = StringVar()
        self.FileTypeVar.set('f')
        self.csvB = Checkbutton(self.editFrame, text='csv', variable=self.FileTypeVar, onvalue='.csv')
        self.txtB = Checkbutton(self.editFrame, text='txt', variable=self.FileTypeVar, onvalue='.txt')
        self.xlsB = Checkbutton(self.editFrame, text='xls', variable=self.FileTypeVar, onvalue='.xls')

        self.exportfileNameL = Label(self.editFrame,text="File Name:")
        self.exportfileNameE = Entry(self.editFrame,width=15)
        self.exportfileB = Button(self.editFrame,text="Export Data",command= self.ExportDataC) #command= self.ExportDataF  must pass the data to excel

        self.programMsg = Label(self,text="Program Messages...")

        #GRIDLAYOUT
        self.grid()
        self.header.grid(row=0,column=0,sticky=N+E+W,pady=0)
        self.selectFileB.grid(row=2,column=0,padx=70,pady=6,sticky=W+N+S)
        self.treeview.grid(row=0,column=0,padx=6)
        self.infoframe.grid(row=0,column=3,sticky=N+S,padx=6)

        self.InfoInsideFrame2.grid(row=1,column=0,pady=7)
        self.InfoInsideFrame3.grid(row=2,column=0,pady=7)
        self.InfoInsideFrame6.grid(row=5,column=0,pady=7)
        self.InfoInsideFrame4.grid(row=3,column=0,pady=7)
        self.InfoInsideFrame5.grid(row=4,column=0,pady=7)

        self.editFrame.grid(row=4,columnspan=1,column=0,sticky=E+W,padx=3)
        self.bigFrame.grid(row=3,column=0,sticky=EW)
        self.showDataB.grid(row=0,column=2,padx=3)
        self.DataLabel.grid(row=0,column=0,sticky=E+W)
        self.Name.grid(row=0,column=0,sticky=E)
        self.ID.grid(row=0,column=0,sticky=E)
        self.Surname.grid(row=0,column=0,sticky=E)
        self.Dept.grid(row=0,column=0,sticky=E)
        self.GPA.grid(row=0,column=0,sticky=E)
        self.MP1.grid(row=0,column=2,sticky=E)
        self.MP2.grid(row=0,column=2,sticky=E)
        self.MP3.grid(row=0,column=2,sticky=E)
        self.MT.grid(row=0,column=2,sticky=E)
        self.Final.grid(row=0,column=2,sticky=E)

        self.NameL.grid(row=0,column=1,sticky=E)
        self.IDL.grid(row=0,column=1,sticky=E)
        self.SurnameL.grid(row=0,column=1,sticky=E)
        self.DeptL.grid(row=0,column=1,sticky=E)
        self.GPAL.grid(row=0,column=1,sticky=E)
        self.MP1L.grid(row=0,column=3,sticky=E)
        self.MP2L.grid(row=0,column=3,sticky=E)
        self.MP3L.grid(row=0,column=3,sticky=E)
        self.MTL.grid(row=0,column=3,sticky=E)
        self.FinalL.grid(row=0,column=3,sticky=E)


        self.projectN.grid(row=0,column=0,pady=3)
        self.Grades.grid(row=1,column=0,pady=3)
        self.ExportAs.grid(row=2, column=0,pady=3)

        self.mp1L.grid(row=0,column=1)
        self.mp1E.grid(row=1,column=1,padx=6)

        self.mp2L.grid(row=0,column=2)
        self.mp2E.grid(row=1, column=2,padx=6)

        self.mp3L.grid(row=0,column=3)
        self.mp3E.grid(row=1,column=3,padx=6)

        self.mtL.grid(row=0,column=4)
        self.mtE.grid(row=1,column=4,padx=6)

        self.feL.grid(row=0,column=5)
        self.feE.grid(row=1,column=5,padx=6)

        self.saveGradesB.grid(row=0,column=7,rowspan=2,sticky=N+E+S,padx=(150, 0))

        self.csvB.grid(row=3,column=1)
        self.txtB.grid(row=4,column=1)
        self.xlsB.grid(row=5,column=1)

        self.exportfileNameL.grid(row=3,column=2)
        self.exportfileNameE.grid(row=3,column=3)
        self.exportfileB.grid(row=4,rowspan=2,columnspan=2,column=2,sticky=E+W)

        self.programMsg.grid(row=10,column=0,sticky=SW)

    def TreeviewClick(self,event):
        if self.CheckIfFileLoaded.get() == "No":
            self.programMsg.configure(text="INFO: Please Load the Files First!", background='red')
            self.TreeViewIsClicked.set("No")
        else:
            self.TreeViewIsClicked.set("Yes")




    def SelectFileC(self):
        try:
            Student_Data().SelectFile(self.CheckIfFileLoaded) # Calling the function that we defined in Student_Data class to create student list, object for each student and student objects list
            for student in Student_Data.studentObjectList:

                self.treeview.insert('', 'end', values=[student.id,
                                                                 student.name,
                                                                 student.surname, student.section,
                                                                 student.dept, \
                                                                 student.gpa, student.mp1,
                                                                 student.mp2, \
                                                                 student.mp3, student.mt,
                                                                 student.fe])

            self.programMsg.configure(text="INFO: File Loaded.", background='green')

        except:
            self.programMsg.configure(text="INFO: Loading Failed! Please Try Again.",background='red')




    def ShowDataC(self):
        if self.CheckIfFileLoaded.get() == "No": #Checking if excel file is loaded
            self.programMsg.configure(text="INFO: Please Load the Files First!", background='red')

        else:
            try:

                self.selectedRow = self.treeview.selection()
                self.idd = self.treeview.item(self.selectedRow)['values'][0]
                self.IDL.configure(text=(self.treeview.item(self.selectedRow))['values'][0])
                self.NameL.configure(text=(self.treeview.item(self.selectedRow))['values'][1])
                self.SurnameL.configure(text=(self.treeview.item(self.selectedRow))['values'][2])
                self.DeptL.configure(text=(self.treeview.item(self.selectedRow))['values'][4])
                self.GPAL.configure(text=(self.treeview.item(self.selectedRow))['values'][5])
                self.GPAL.configure(text=(self.treeview.item(self.selectedRow))['values'][5])
                self.MP1L.configure(text=  Student_Data.StudentDict[self.idd][0])
                self.MP2L.configure(text=  Student_Data.StudentDict[self.idd][1])
                self.MP3L.configure(text=  Student_Data.StudentDict[self.idd][2])
                self.MTL.configure(text=   Student_Data.StudentDict[self.idd][3])
                self.FinalL.configure(text= Student_Data.StudentDict[self.idd][4])

                self.mp1E.delete(0,"end")
                self.mp1E.insert(END,int(Student_Data.StudentDict[self.idd][0]))

                self.mp2E.delete(0, "end")
                self.mp2E.insert(END,int(Student_Data.StudentDict[self.idd][1]) )

                self.mp3E.delete(0,"end")
                self.mp3E.insert(END,int(Student_Data.StudentDict[self.idd][2]) )

                self.mtE.delete(0,"end")
                self.mtE.insert(END,int(Student_Data.StudentDict[self.idd][3]) )

                self.feE.delete(0,"end")
                self.feE.insert(END,int(Student_Data.StudentDict[self.idd][4]) )

                self.programMsg.configure(text="INFO: Student Info Being Displayed!", background='green')

            except:
                self.programMsg.configure(text="INFO: Please Select A Student First!",background='red')


    def SaveGradesC(self):
        if self.CheckIfFileLoaded.get() == "No": # Checkin if excel file is loaded
            self.programMsg.configure(text="INFO: Please Load the Files First!", background='red')
        else:
            if self.TreeViewIsClicked.get() == 'Yes' : # Checking if student is selected by the user. We have problem in here because the solution i find for this is not enough
                                                       # because if the user click just on the column headin but not the any students, this function will understand it as student selected :)
                try:

                    Student_Data.StudentDict[self.idd][0] = int(self.mp1E.get())
                    Student_Data.StudentDict[self.idd][1] = int(self.mp2E.get())
                    Student_Data.StudentDict[self.idd][2] = int(self.mp3E.get())
                    Student_Data.StudentDict[self.idd][3] = int(self.mtE.get())
                    Student_Data.StudentDict[self.idd][4] = int(self.feE.get())

                    self.MP1L.configure(text=Student_Data.StudentDict[self.idd][0])
                    self.MP2L.configure(text=Student_Data.StudentDict[self.idd][1])
                    self.MP3L.configure(text=Student_Data.StudentDict[self.idd][2])
                    self.MTL.configure(text=Student_Data.StudentDict[self.idd][3])
                    self.FinalL.configure(text=Student_Data.StudentDict[self.idd][4])

                    self.programMsg.configure(text="INFO: Grades are saved!", background='green')

                except:
                    self.programMsg.configure(text="INFO: Warning! The Type Of The Grade Is Incorrect!", background='red')

            else:
                self.programMsg.configure(text="INFO: Please Select A Student First!", background='red')

    def ExportDataC(self):

        if self.CheckIfFileLoaded.get() == "No":
            self.programMsg.configure(text="INFO: Please Load the Files First!", background='red')
        else:


            if self.FileTypeVar.get() == 'f': # Checking if the user selected any type of the file type to export
                self.programMsg.configure(text="INFO: Please provide a file type!",background='red')
            
            else:
            
                if self.exportfileNameE.get() == "":
                    self.programMsg.configure(text= "INFO: Please provide a file name!",background='red')
            
                else:
                    if self.FileTypeVar.get() == ".xls":
            
                        self.book = Workbook()
                        self.stheet = self.book.add_sheet('sheet 1')
                        self.stheet.write(0, 0, 'ID')
                        self.stheet.write(0, 1, 'NAME')
                        self.stheet.write(0, 2, 'SECTION')
                        self.stheet.write(0, 3, 'DEPT')
                        self.stheet.write(0, 4, 'GPA')
                        self.stheet.write(0, 5, 'MP1')
                        self.stheet.write(0, 6, 'MP2')
                        self.stheet.write(0, 7, 'MP3')
                        self.stheet.write(0, 8, 'MT')
                        self.stheet.write(0, 9, 'FINAL')
            
                        for i in range(1,len(Student_Data.studentList)):
                            self.stheet.row(i).write(0, Student_Data.StudentDict[Student_Data.StudentDict.keys()[i - 1]][5])
                            self.stheet.row(i).write(1, Student_Data.StudentDict[Student_Data.StudentDict.keys()[i - 1]][6])
                            self.stheet.row(i).write(2, Student_Data.StudentDict[Student_Data.StudentDict.keys()[i - 1]][7])
                            self.stheet.row(i).write(3, Student_Data.StudentDict[Student_Data.StudentDict.keys()[i - 1]][8])
                            self.stheet.row(i).write(4, Student_Data.StudentDict[Student_Data.StudentDict.keys()[i - 1]][9])
                            self.stheet.row(i).write(5, Student_Data.StudentDict[Student_Data.StudentDict.keys()[i - 1]][0])
                            self.stheet.row(i).write(6, Student_Data.StudentDict[Student_Data.StudentDict.keys()[i - 1]][1])
                            self.stheet.row(i).write(7, Student_Data.StudentDict[Student_Data.StudentDict.keys()[i - 1]][2])
                            self.stheet.row(i).write(8, Student_Data.StudentDict[Student_Data.StudentDict.keys()[i - 1]][3])
                            self.stheet.row(i).write(9, Student_Data.StudentDict[Student_Data.StudentDict.keys()[i - 1]][4])
            
                        self.book.save(self.exportfileNameE.get()+self.FileTypeVar.get())
                        self.book.save(TemporaryFile())
            
                        self.programMsg.configure(text= "INFO:"+" "+self.exportfileNameE.get()+self.FileTypeVar.get()+" "+ "saved!",background="green")
            
                    elif self.FileTypeVar.get() == ".csv":
                        self.programMsg.configure(text="INFO: Type not supported!",background="red")
            
            
                    elif self.FileTypeVar.get() == ".txt":
            
                        self.SavedTextFile= open(self.exportfileNameE.get()+".txt","w+")
            
                        for i in range(1,len(Student_Data.studentList)):
                            self.SavedTextFile.write(str(int(Student_Data.StudentDict[Student_Data.StudentDict.keys()[i - 1]][5]))+",  "+str(Student_Data.StudentDict[Student_Data.StudentDict.keys()[i - 1]][7])+",  "+str(Student_Data.StudentDict[Student_Data.StudentDict.keys()[i - 1]][8])+",  "+str(Student_Data.StudentDict[Student_Data.StudentDict.keys()[i - 1]][9])+",  "+str(Student_Data.StudentDict[Student_Data.StudentDict.keys()[i - 1]][0])+",  "+str(Student_Data.StudentDict[Student_Data.StudentDict.keys()[i - 1]][1])+",  "+str(Student_Data.StudentDict[Student_Data.StudentDict.keys()[i - 1]][2])+",  "+str(Student_Data.StudentDict[Student_Data.StudentDict.keys()[i - 1]][3])+",  "+str(Student_Data.StudentDict[Student_Data.StudentDict.keys()[i - 1]][4])+"\n"+"\n")
                        self.SavedTextFile.close()
                        self.programMsg.configure(text= "INFO:"+" "+self.exportfileNameE.get()+self.FileTypeVar.get()+" "+ "saved!",background="green")




def main():
    root = Tk()
    root.title("Grades Management Tool v2.0")
    root.geometry("800x520+280+100")
    app = GMT_Interface(root)
    root.mainloop()

main()

