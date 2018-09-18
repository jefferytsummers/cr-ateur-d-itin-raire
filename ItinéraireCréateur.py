
##### Imports #####
import json # for saving and loading files

from os import path

# Used for creating the PDFs
from pylatex import Document, LongTable, Command, MultiColumn, MultiRow, NoEscape

# Used for creating the GUI
from tkinter import Tk, Frame, Scrollbar, Label, Menu, \
                    BOTH, Y, X, TRUE, FALSE, NW, TclError
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
from tkinter.tix import *
from tkinter.ttk import Entry, Button, Checkbutton


class ModdableTableEntry(Entry):
    def __init__(self, master, numentries=6, textvariables='', justify='left', width=20):
        # Inherit from tkinter.Entry
        Entry.__init__(self, master)

        # Create a frame for the entry box
        entryFrame = Frame(master)
        entryFrame.pack(side=LEFT, fill=X, expand=1)

        # Create a frame for the checkboxes
        checkFrame = Frame(master)
        checkFrame.pack(side=RIGHT)

        # Initialize the variables stored in the moddable
        # entrybox
        textMods = ['Italique', 'Gras', 'Souligner']
        self.checkVars = [IntVar() for _ in range(3)]

        # Create and array to store entries
        self.entries = []
        
        # Entry boxes
        # Include multiple entries if needed
        if numentries != 0:
            for i in range(numentries):
                entryBox = Entry(entryFrame, textvariable=textvariables[i], width=width, justify=justify)
                entryBox.pack(side=LEFT, fill=X, expand=1, pady=1.1, padx=5)      
                # Append each entry box to the entries array
                self.entries.append(entryBox)
        # Check boxes
        for i in range(3):
            Checkbutton(checkFrame, variable=self.checkVars[i]).pack(side=RIGHT)
            Label(checkFrame, text=textMods[i]).pack(side=RIGHT)

    def state(self):
        # Display the state of the entry box
        return str(self.checkVars[0].get()) + ' ' + str(self.checkVars[1].get()) \
               + ' ' + str(self.checkVars[2].get())

class ModdableEntry(Entry):
    def __init__(self, master, numentries=0, textvariable='', justify='left', width=20):
        # Inherit from tkinter.Entry
        Entry.__init__(self, master)

        # Create a frame for the entry box
        entryFrame = Frame(master)
        entryFrame.pack(side=LEFT, fill=X, expand=1)

        # Create a frame for the checkboxes
        checkFrame = Frame(master)
        checkFrame.pack(side=RIGHT)

        # Initialize the variables stored in the moddable
        # entrybox
        textMods = ['Italique', 'Gras', 'Souligner']
        self.checkVars = [IntVar() for _ in range(3)]

        # Create and array to store entries
        self.entries = []
        
        # Entry boxes
        # Include multiple entries if needed
        if numentries != 0:
            for i in range(numentries):
                entryBox = Entry(entryFrame, textvariable=textvariable, width=width, justify=justify)
                entryBox.pack(side=LEFT, fill=X, expand=1, pady=2)      
                # Append each entry box to the entries array
                self.entries.append(entryBox)
        else:
            entryBox = Entry(entryFrame, textvariable=textvariable, width=width, justify=justify)
            entryBox.pack(side=LEFT, fill=X, expand=1, pady=2)
            self.entries.append(entryBox)

        # Check boxes
        for i in range(3):
            Checkbutton(checkFrame, variable=self.checkVars[i]).pack(side=RIGHT)
            Label(checkFrame, text=textMods[i]).pack(side=RIGHT)


    def state(self):
        # Display the state of the entry box
        state =str(self.checkVars[0].get()) + ' ' + str(self.checkVars[1].get()) \
                   + ' ' + str(self.checkVars[2].get())
        return state

    
    def row(self, entry):
        # Scan entries, if entries have not been deleted then store the contents
        # of individual entries as elements of the list 'row'
        # Return 'row'.
        row = []
        if entry in self.entries:
            try:
                row.append(entry.get())
            except TclError:
                pass
        return row


class App(Frame):
    def __init__(self, master):

        #----------------------------------------------------------------------------------------------------------------------------
        # Create 'itineraires' directory if it doesn't exist
        if not path.isdir('itineraires'):
            os.mkdir('itineraires')
        if not path.isdir('itinéraires en cours'):
            os.mkdir('itinéraires en cours')

        self.master = master
        # Inherit members from the Tk class 'Frame'
        Frame.__init__(self, master)

        # Create a scrollable window within the master frame.
        # With the exception of the toolbar, all other frame will lie within
        # this frame.
        self.mainFrame = ScrolledWindow(master, scrollbar=Y)
        self.mainFrame.pack(fill=BOTH, expand=1)

        # Initialize the toolbar. The toolbar contains four buttons:
        # 'Add Title', 'Add Subtitle', 'Add Table', and 'Generate PDF'
        self.toolBar = Menu(master)
        self.toolBar.add_command(label='Produire PDF',  \
                                 command=lambda: self.generate_pdf())
        self.toolBar.add_command(label='Enregistrer l\'itinéraire', \
                                 command=lambda: self.save_itinerary())
        self.toolBar.add_command(label='Itinéraire de Chargement', \
                                 command=lambda: self.load_itinerary())       

        #----------------------------------------------------------------------------------------------------------------------------

        # Create an entry for the name of the itinerary
        # This information is accessed by the 'save_itinerary' method
        saveFrame = Frame(self.mainFrame.window)
        saveFrame.pack(side=TOP)
        self.itinName = StringVar()
        Label(saveFrame, text='Nom de l\'itinéraire: ', relief=GROOVE).pack(side=LEFT)
        self.saveEntry = Entry(saveFrame, textvariable=self.itinName, width=50)
        self.saveEntry.pack(side=LEFT)

        #----------------------------------------------------------------------------------------------------------------------------

        # Create a frame to store the header information
        headFrame = Frame(self.mainFrame.window, width=self.mainFrame.winfo_reqwidth()-500)
        headFrame.pack(side=TOP, fill=X, expand=0)

        # First initialize a list of 9 StringVar that will be assigned to each
        # entry box
        self.headerInfo = [StringVar(value='') for _ in range(9)]
        headerLabels = ['Titre Principal de l\'événement', 'Durée de l\'événement Principal', 'Placement/Concentration', \
                        'ITINERAIRE DE', 'Voitures', 'Durée de l\'événement', \
                        'Secteur', 'Moyennes (pour placement uniquement)', 'Total Distance/Moment Idéal']
        # Initialize an array to hold all of the header entries
        # This array can be mutated by the save/load itinerary methods as well as the 
        # generate_pdf method.
        self.headerEntries = []

        # Create 9 entry boxes and 9 labels for the header info
        # Frames for the labels and entry boxes
        HLFrame = Frame(headFrame, width=50)
        HLFrame.pack(side=LEFT)
        HEFrame = Frame(headFrame)
        HEFrame.pack(side=LEFT, fill=X, expand=1, pady=2)
        for i in range(9):
            # Create a temporary frame to store the modded entries in
            tmpFrame = Frame(HEFrame)
            tmpFrame.pack(fill=X, expand=1)
            if i == 5 or i == 7 or i == 8:
                tmpEntry = ModdableEntry(tmpFrame, textvariable=self.headerInfo[i], \
                                         justify='right')
                # Add each header entry to the headerEntries array so that the values
                # stored in the entries can be accessed by the save, load, and generate_pdf
                # methods
                self.headerEntries.append(tmpEntry)
                Label(HLFrame, text=headerLabels[i], relief=GROOVE, \
                      justify='left').pack(fill=X, pady=2)
            else:
                tmpEntry = ModdableEntry(tmpFrame, textvariable=self.headerInfo[i])
                # Add each header entry to the headerEntries array so that the values
                # stored in the entries can be accessed by the save, load, and generate_pdf
                # methods
                self.headerEntries.append(tmpEntry)
                Label(HLFrame, text=headerLabels[i], relief=GROOVE, \
                      justify='left').pack(fill=X, pady=2)

        #----------------------------------------------------------------------------------------------------------------------------
        # Pack the table frame between the header frame and the footer frame
        self.tableFrame = Frame(self.mainFrame.window)
        self.tableFrame.pack(side=TOP)
        self.tableEntries = []
        self.tableData = []
        LBFrame = Frame(self.tableFrame)
        LBFrame.pack(side=TOP, expand=0)
        labelFrame = Frame(self.tableFrame)
        labelFrame.pack(side=TOP)
        tableLables = ['Time Control', 'Communes', 'Routes', 'Partiel Dist.', 'Total Dist.', 'Horaire Approx']
        for i in range(6):
            Label(LBFrame, text=tableLables[i], relief=GROOVE).pack(side=LEFT, padx=(0,12*((3/4)*i+6)))
        Button(LBFrame, text='Add Row', command=lambda: self.add_row()).pack(side=RIGHT)            
        #----------------------------------------------------------------------------------------------------------------------------

        # Create the footer
        # Create and pack the footer frame
        footerFrame = Frame(self.mainFrame.window)
        footerFrame.pack(side=BOTTOM, fill=X)
        # First create a list of 4 stringVars that will be assigned to each footer entry
        self.footerInfo = [StringVar(value='') for _ in range(4)]
        footerLabels = ['Cartes Michelin', 'A L\'ATTENTION DES SPECTATEURS', \
                        'Z.R. ouvertes avant l\'heure ideale de passage',
                        'Z.R. will be open prior to the target time']
        # Initialize an array to hold all of the header entries
        # This array can be mutated by the save/load itinerary methods as well as the 
        # generate_pdf method.
        self.footerEntries = []

        # Frames for the entry boxes and labels
        FLFrame = Frame(footerFrame)
        FLFrame.pack(side=LEFT)
        FEFrame = Frame(footerFrame)
        FEFrame.pack(side=LEFT, fill=X, expand=1)
        # Create 4 entry boxes and 9 labels for the footer info
        for i in range(4):
            # Temp frame to store the modded entries in
            tmpFrame = Frame(FEFrame)
            tmpFrame.pack(fill=X, expand=1)
            Label(FLFrame, text=footerLabels[i], relief=GROOVE, \
                justify='left').pack(fill=X, pady=2)
            tmpEntry = ModdableEntry(tmpFrame, textvariable=self.footerInfo[i])
            # Add each header entry to the headerEntries array so that the values
            # stored in the entries can be accessed by the save, load, and generate_pdf
            # methods
            self.footerEntries.append(tmpEntry)
        

    def reload(self):
        # Clear all frames and reinitialize
        # Called when loading a new file
        
        # Warn the user that all current progress will be lost
        answer = messagebox.askokcancel(title='Avertissement tout progrès sera perdu', \
                        message='Avertissement tout progrès sera perdu, veuillez \
                                enregistrer l\'itinéraire avant de continuer')
        if answer:
            # remove primary frame and reinitialize
            self.mainFrame.destroy()
            self.__init__(self.master)
            return True
        else:
            return False


    def add_row(self):
        # Create a temporary frame to store the modded entries in
        tmpFrame = Frame(self.tableFrame)
        tmpFrame.pack(fill=X, expand=1)
        newTableData = [StringVar() for _ in range(6)]
        self.tableData.append(newTableData)
        
        tmpEntry = ModdableTableEntry(tmpFrame, numentries=6, textvariables=newTableData)
        print(tmpEntry.state())
        self.tableEntries.append(tmpEntry)
        Button(tmpFrame, text='X', command=lambda: self.remove_row(newTableData, tmpFrame)).pack(side=RIGHT)
    

    def remove_row(self, data, frame):
        # Delete a row of a given index in the table
        self.tableData.remove(data)
        frame.destroy()

    def save_itinerary(self):

        # Let user know that the itinerary will be saved and will erase the current
        # itinerary
        if self.itinName.get():
            answer = messagebox.askokcancel(title='Avertissement tout progrès sera perdu', \
                        message='Attention l\'itinéraire suivant sera écrasé: %s.json' \
                        % self.itinName.get())
        if answer == False:
            pass
        # Store all class variables into a python dictionary and write to a json
        # file.
        # Create the dictionary that stores all of the itinerary info
        self.itinInfo = {
            "Header": [_.get() for _ in self.headerInfo],
            "Footer": [_.get() for _ in self.footerInfo],
            # Check if the entries in the header need to be modified
            "HeaderMods": [_.state() for _ in self.headerEntries],
            "FooterMods": [_.state() for _ in self.footerEntries],
            "Table": [[_.get() for _ in tableRow] for tableRow in self.tableData],
            "ItineraryName": self.itinName.get()
        }
        # Save the itinerary
        saveFileName = self.itinName.get()
        with open('itinéraires en cours\\' + saveFileName + '.json', 'w') as outfile:
            json.dump(self.itinInfo, outfile)
        # Change the savestate of the program to true
        self.savestate = True
        

    def load_itinerary(self):
        # Reload the itinerary so that it is blank
        load = self.reload()
        if load:
            itinFile = filedialog.askopenfilename()
            with open(itinFile, 'r') as infile:
                self.itinInfo = json.load(infile)
            # Allows me to manipulate the entry fields in the program
            allHeaderInfo = []
            for i in range(len(self.headerEntries)):
                allHeaderInfo.append([self.headerEntries[i], self.headerInfo[i]])
            allFooterInfo = []
            for i in range(len(self.footerEntries)):
                allFooterInfo.append([self.footerEntries[i], self.footerInfo[i]])
            
            headerInfo = self.itinInfo["Header"]
            for i in range(len(self.headerEntries)):
                self.headerInfo[i].set(headerInfo[i])

            footerInfo = self.itinInfo["Footer"]
            for i in range(len(self.footerEntries)):
                self.footerInfo[i].set(footerInfo[i])
            # Check the number of rows needed to be added
            i = 0
            for tableRow in self.itinInfo["Table"]:
                self.add_row()
                for j in range(6):
                    self.tableData[i][j].set(self.itinInfo["Table"][i][j])
                i += 1
            self.itinName.set(self.itinInfo["ItineraryName"])

            # The elements of self.tableData are the rows of StringVar's that the 
            # rows of the table hold. I need to access each individual StringVar
            # and set its value to the corresponding value found in
            # self.itinInfo["Table"]

    def generate_pdf(self):
        # Initialize the document
        try:
            doc = Document(default_filepath=self.itinName.get())
        except IndexError:
            doc = Document(default_filepath="temp")
        finally:
            doc = Document(default_filepath="temp")

        # Ensure the document is made with A4 paper in mind
        doc.preamble.append(NoEscape(r'\usepackage[a4paper,margin=0.15in,includeheadfoot=True]{geometry}'))
        #------------------------------------------------------------------------------------------------------------
        # Returns the ith header entry string 
        def modifiedHString(i, string, state):
            # Check for all possible modifications to the entries
            # (italics bold underline)
            if state == '0 0 0': # No modification
                mod = r''
            if state == '1 0 0': # Italics
                mod = r'\textit{'
            if state == '1 1 0': # Italics and bold
                mod = r'\textit{\textbf{'
            if state == '1 1 1': # Italics, bold, and underline
                mod = r'\textit{\textbf{\underline{'
            if state == '0 1 1': # Bold and underline
                mod = r'\textbf{\underline{'
            if state == '0 0 1': # Underline
                mod = r'\underline{'
            if state == '0 1 0': # Bold
                mod = r'\textbf{'
            if state == '1 0 1': # Italics and underline
                mod = r'\textit{\underline{'
            # Check alignment of headers
            # right alignment
            if i == 5 or i == 7 or i == 8:
                align = r'\begin{flushright}'
                endalign = r'\end{flushright}'
            # left alignment
            if i == 6:
                align = r'\begin{flushleft}'
                endalign = r'\end{flushleft}'
            # center alignment
            if i == 0 or i == 1 or i == 2 or i == 3 or i == 4:
                align = r'\begin{center}'
                endalign = r'\end{center}'
            if string != '':
                if i == 0:
                    size = r'\LARGE{'
                    return str(align + ' ' + mod + size + string + '}' *(int(state[0])+int(state[2])+int(state[4])) + '}' + ' ' + endalign)
                return str(align + ' ' + mod + string + '}' *(int(state[0])+int(state[2])+int(state[4])) + ' ' + endalign)
            else:
                return '%'
            
        # For the header and footer we have two class variables accessible by
        # all methods:
        # 1.) headerInfo
        # 2.) footerInfo
        # Both of these variables are lists whos elements contain ModdableEntry
        # objects.
        # Accessing the 'get()' method of the ModdableEntry returns the contents
        # of the entry box, while accessing the 'state()' method of the ModdableEntry
        # returns the state of the checkboxes. The state of the textboxes determines
        # how the contents should be modified. See below for details.
        allHeaderInfo = []
        for i in range(len(self.headerEntries)):
            allHeaderInfo.append([self.headerEntries[i], self.headerInfo[i]])

        # Write all the header information to the document
        i = 0
        for entry, text in allHeaderInfo:
            string = modifiedHString(i, text.get(), entry.state())
            if string != '' and type(string) != None:
                doc.append(NoEscape(r'%s' % string))
            i += 1
        #------------------------------------------------------------------------------------------------------------
        # Write all the table information to the document
        tableWidth = "p{2.25cm}|p{7.0cm}|p{1.5cm}|p{1.5cm}|p{1.5cm}|p{3.5cm}"
        labels = ['', 'Communes', 'Routes', 'Partiel', 'Total', 'Horaire Approx']


        def modifiedTString(i, string, state):
            # Check for all possible modifications to the entries
            # (italics bold underline)
            if state == '0 0 0': # No modification
                mod = r''
            if state == '1 0 0': # Italics
                mod = r'\textit{'
            if state == '1 1 0': # Italics and bold
                mod = r'\textit{\textbf{'
            if state == '1 1 1': # Italics, bold, and underline
                mod = r'\textit{\textbf{\underline{'
            if state == '0 1 1': # Bold and underline
                mod = r'\textbf{\underline{'
            if state == '0 0 1': # Underline
                mod = r'\underline{'
            if state == '0 1 0': # Bold
                mod = r'\textbf{'
            if state == '1 0 1': # Italics and underline
                mod = r'\textit{\underline{'

            if string != '':
                return mod + string + '}' *(int(state[0])+int(state[2])+int(state[4]))
            else:
                return ' '

        with doc.create(LongTable((tableWidth))) as data_table:
            data_table.add_hline()
            data_table.add_row(labels)
            data_table.add_hline()
            data_table.end_table_header()
            data_table.end_table_footer()
            data_table.end_table_last_footer()
            i = 0
            for tableRow in self.tableData:
                try:
                    strings = [NoEscape(r'%s' % modifiedTString(i, _.get(), self.tableEntries[i].state())) for _ in tableRow]
                    # for string in strings:
                        # if string != ''
                    
                    data_table.add_row(strings)
                    i += 1
                except TclError:
                    continue
            data_table.add_hline()

        #------------------------------------------------------------------------------------------------------------
        # Write all the footer information to the document
        # Returns the ith footer string
        def modifiedFString(i, string, state):
            # Check for all possible modifications to the entries
            # (italics bold underline)
            if state == '0 0 0': # No modification
                mod = r''
            if state == '1 0 0': # Italics
                mod = r'\textit{'
            if state == '1 1 0': # Italics and bold
                mod = r'\textit{\textbf{'
            if state == '1 1 1': # Italics, bold, and underline
                mod = r'\textit{\textbf{\underline{'
            if state == '0 1 1': # Bold and underline
                mod = r'\textbf{\underline{'
            if state == '0 0 1': # Underline
                mod = r'\underline{'
            if state == '0 1 0': # Bold
                mod = r'\textbf{'
            if state == '1 0 1': # Italics and underline
                mod = r'\textit{\underline{'
            # Check alignment of headers
            # left alignment
            if i == 0:
                align = r'\begin{flushleft}'
                endalign = r'\end{flushleft}'

            # center alignment
            if i == 1 or i == 2 or i == 3:
                align = r'\begin{center}'
                endalign = r'\end{center}'

            if string != '':
                return align + ' ' + mod + string + '}' *(int(state[0])+int(state[2])+int(state[4])) + ' ' + endalign
            else:
                return '%'
        allFooterInfo = []
        for i in range(len(self.footerEntries)):
            allFooterInfo.append([self.footerEntries[i], self.footerInfo[i]])

        i = 0
        for entry, text in allFooterInfo:
            string = modifiedFString(i, text.get(), entry.state())
            if string != '' and type(string) != None:
                doc.append(NoEscape(r'%s' % string))
            i += 1
        #------------------------------------------------------------------------------------------------------------

        doc.generate_pdf('itineraires\\' + self.itinName.get(), compiler='pdflatex', clean_tex=False)

def main():
    root = Tk()
    root.title("Itinerary Creator for ACM")

    app = App(root)

    root.config(menu=app.toolBar)
    root.iconbitmap('acm.ico')
    root.geometry("1200x900+0+0") 
    root.mainloop()

main()