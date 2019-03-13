from tkinter import *
import numpy as np
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
import os


class GUI(object):

    def __init__(self, Master):
        self.row = 0
        self.Files = [[]]
        self.Files.append([])
        self.num_of_Additional = 0
        self.PlateType = []
        self.Win = Master

    #Button Functions
    def End(self):
        self.Win.quit()

    def File(self,FileName, Type, Count):
        if (Count == 0):
            File = askopenfilename()
        else:
            File = askdirectory()
        FileName["text"] = os.path.basename(File)
        self.Files[0].append(File)
        if(Type == "R"):
            self.Files[1].append(Count)
        else:
            self.Files[1].append(Type)

    #Function to create headings and buttons
    def Organizer(self, Heading, Type):

        #Create a Prompt and a Frame
        Prompt = Label(self.Win, text = Heading)
        F = Frame(self.Win)

        #Display the prompt
        Prompt.grid(row = self.row, column = 0,pady = 5, padx = 100)
        self.row += 1

        #IF this is an FILE button, follow this code to display it
        if (Type == "FILE"):
            Count = self.num_of_Additional
            FileName = Button(F,text = "File Name", command = lambda: self.File(FileName,Type, Count))
            FileName.grid(row = 0, column = 2)

        #IF this isn't an FILE button, run this code to display the buttons
        else:
            Formats = [ "96 Well", "24 Well", "6 Well"]
            Format_Var = StringVar(self.Win)
            Format_Var.set("96 Well") # default value
            self.PlateType.append(Format_Var)
            FileName = OptionMenu(F, Format_Var, *Formats)
            FileName.grid(row = 0, column = 2)

        return F

    def run(self):

        self.Win.title("RNA Velocity ")

        self.T_Seq = self.Organizer("Please select the JMP Excel Output", "FILE")
        self.T_Seq.grid(row = self.row, column = 0)
        self.row += 1

        self.G_Seq = self.Organizer("Please select the type of Plate", "PLATE")
        self.G_Seq.grid(row = self.row, column = 0)
        self.row += 1

        #Done Button
        self.Done = Button(self.Win, text = "Done", command = self.End)
        self.Done.grid(row = self.row, column = 0, pady = 50)
        self.row += 1


def Get_Files():

    Root = Tk()
    Program = GUI(Root)
    Program.run()
    Root.mainloop()

    JMP_Excel = Program.Files[0][0]
    Plate_Format = Program.PlateType[0].get()

    #Add code submiting for additional commands

    Root.destroy()

    return( JMP_Excel, Plate_Format)
