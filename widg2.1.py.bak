import sys
sys.path.append(r'C:\Python24\Lib')

import clr
import sys
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

from System.Drawing import Color, Point
from System.Windows.Forms import Application, Button, Form, Label,RadioButton
from System.Windows.Forms import OpenFileDialog, DialogResult, BorderStyle

#----------------------------------------------------------------
    # find last
#----------------------------------------------------------------
def find_last(field):  #returns last populated position in field
    pos = 0
    last = 0
    for e in field:
        if e != ' ':
            last = pos
            pos = pos + 1
        else:
            pos = pos + 1
    return last
    
    
#----------------------------------------------------------------
    # check for all blank field
#----------------------------------------------------------------
def all_blank(field):  #returns True if empty field
    test = True
    for e in field:
        if e != ' ':
            test = False
    return test

#----------------------------------------------------------------
    # parse layout
#----------------------------------------------------------------
def parse_layout(layout,delim,header=True): #converts layout file to list with headers
    fields = []
    new_fields = []
    pos = 0
    while pos != -1:
        pos = layout.find(',')
        if pos != -1:
            fields.append(layout[0:pos])
            layout = layout[pos + 1:]
        else:
            fields.append(layout)
    if not header:
        return fields
    headers = fields[1::2]  #strips headers out of list
    headers = delim.join(headers)
    out_fields = fields[::2]
    for e in out_fields:  #strips field lengths out of list
        new_fields.append(int(e))  #convert field lengths to int
    return new_fields,headers
    
    
#----------------------------------------------------------------
    # convert file
#----------------------------------------------------------------
def convert(FILENAME,LAYOUT,DELIM=',',HEADER=True):
    print "Hello!"
    layout = open(LAYOUT,"r")
    if HEADER:
        print "yes, header"
        for line in layout:
            fields,headers = parse_layout(line,DELIM) 
    else:
        print "no header"
        for line in layout:
            fields = parse_layout(line,DELIM,HEADER)
    print "fields = " + str(fields)
    layout.close()
    # ---------------------------------------------------------------
    # import file
    file = open(FILENAME,"r")
    # ---------------------------------------------------------------
    # create output file
    out_file_path = FILENAME + '.out'
    out_file = open(out_file_path,"w")
    # ---------------------------------------------------------------
    # process file
    fields_list = []
    first = 0
    end = len(fields)
    for f in fields:  #create list of lists:[start,end] positions for each field
        try:
            fields_list.append([first,first + f])
        except TypeError:
            return False
        first = first + f  
    if HEADER:
        out_file.write(headers + '\n')  #write headers as first line in file
    for line in file:
        count = 0    
        for field in fields_list:
            count = count + 1
            if all_blank(line[field[0]:field[1]]):  #check for empty field
                if count == end:
                    out_file.write('\n') #carrage return if last field
                else:
                    out_file.write(DELIM)  #write delimiter if empty field
            else:
                last = find_last(line[field[0]:field[1]]) + 1  #find last populated position in field
                if count == end:
                    out_file.write(line[field[0]:field[0] + last] + '\n') #carrage return if last field
                else:
                    out_file.write(line[field[0]:field[0] + last] + DELIM) #delimiter if not last field
    out_file.close()
    file.close()
    return True




class HelloWorldForm(Form):
    global FILENAME
    global LAYOUT
    global DELIM
    global HEADER
    FILENAME = ''
    LAYOUT = ''
    DELIM = ','
    HEADER = True
    
    def __init__(self):
        # main form attributes
        self.Text = 'ff2delim'
        self.BackColor = Color.LightSlateGray
        self.ForeColor = Color.White
        self.BorderStyle = BorderStyle.Fixed3D
        self.count = 0
        self.Width = 800
        self.Height = 400

#------------------------------------------------------------
        # labels
#------------------------------------------------------------
        # main description
        self.label = Label()
        self.label.Text = "Convert fixed legnth file to delimited"
        self.label.Location = Point(50, 25)
        self.label.Height = 30
        self.label.Width = 600

        # error / compleation messages
        self.label_submit = Label()
        self.label_submit.Text = ""
        self.label_submit.Location = Point(140, 305)
        self.label_submit.Height = 30
        self.label_submit.Width = 600 

        # filename label
        self.label_filename = Label()
        self.label_filename.Text = ""
        self.label_filename.Location = Point(140, 105)
        self.label_filename.Height = 30
        self.label_filename.Width = 600    
        
        #layout label
        self.label_layout = Label()
        self.label_layout.Text = ""
        self.label_layout.Location = Point(140, 135)
        self.label_layout.Height = 30
        self.label_layout.Width = 600    
#------------------------------------------------------------
        # buttons
#------------------------------------------------------------
        # filename button
        button = Button()
        button.Text = "File name"
        button.BackColor = Color.DimGray
        button.Location = Point(50, 100)
        button.Click += self.buttonPressed
        
        # layout button
        button2 = Button()
        button2.Text = "Layout"
        button2.BackColor = Color.DimGray
        button2.Location = Point(50, 130)
        button2.Click += self.button2Pressed
        
        # submit button
        button_submit = Button()
        button_submit.Text = "Convert"
        button_submit.BackColor = Color.DimGray
        button_submit.Location = Point(50, 300)
        button_submit.Click += self.button_submitPressed
        
# -----------------------------------------------------------
        # header yes or no radio buttons
#------------------------------------------------------------
        self.radioLabel2 = Label()
        self.radioLabel2.Text = "Field names in layout"
        self.radioLabel2.ForeColor = Color.White
        self.radioLabel2.Location = Point(50, 160)
        self.radioLabel2.Height = 15
        self.radioLabel2.Width = 600

        self.headerTrue = RadioButton()
        self.headerTrue.Text = "Yes"
        self.headerTrue.Location = Point(55, 175)
        self.headerTrue.Width = 50            
        self.headerTrue.Height = 15            
        self.headerTrue.CheckedChanged += self.HeaderCheckedChanged

        self.headerFalse = RadioButton()
        self.headerFalse.Text = "No"
        self.headerFalse.Location = Point(105, 175)
        self.headerFalse.Width = 50            
        self.headerFalse.Height = 15            
        self.headerFalse.CheckedChanged += self.HeaderCheckedChanged
        
        self.headerDefault = Label()
        self.headerDefault.Text = "(Defaults to yes)"
        self.headerDefault.ForeColor = Color.DarkSlateBlue
        self.headerDefault.Location = Point(50, 195)
        self.headerDefault.Height = 15
        self.headerDefault.Width = 600

#---------------------------------------------------------------
        # delimiter radio buttons
#---------------------------------------------------------------
        self.radioLabel1 = Label()
        self.radioLabel1.Text = "Delimiter"
        self.radioLabel1.Location = Point(50, 220)
        self.radioLabel1.AutoSize = True

        self.radio1 = RadioButton()
        self.radio1.Text = "Comma"
        self.radio1.Location = Point(55, 237)
        self.radio1.Width = 66            
        self.radio1.Height = 15            
        #self.radio1.Checked = True
        self.radio1.CheckedChanged += self.checkedChanged

        self.radio2 = RadioButton()
        self.radio2.Text = "Pipe"
        self.radio2.Location = Point(121, 237)
        self.radio2.Width = 50            
        self.radio2.Height = 15            
        self.radio2.CheckedChanged += self.checkedChanged

        self.radio3 = RadioButton()
        self.radio3.Text = "Tab"
        self.radio3.Location = Point(170, 237)
        self.radio3.Width = 50            
        self.radio3.Height = 15            
        self.radio3.CheckedChanged += self.checkedChanged

        self.label3 = Label()
        self.label3.Text = "(Defaults to comma delimited)"
        self.label3.ForeColor = Color.DarkSlateBlue
        self.label3.Location = Point(50, 255)
        self.label3.Height = 15
        self.label3.Width = 600
#----------------------------------------------------------------
        # add controles
#----------------------------------------------------------------
        self.Controls.Add(self.label)          #"Convert fixed legnth file to delimited"

        self.Controls.Add(self.radioLabel2)    #file names in layout
        self.Controls.Add(self.headerTrue)     #true
        self.Controls.Add(self.headerFalse)    #false
        self.Controls.Add(self.headerDefault)  #defalult message

        self.Controls.Add(self.radioLabel1)    #"delimiter"
        self.Controls.Add(self.radio1)         #comma
        self.Controls.Add(self.radio2)         #pipe
        self.Controls.Add(self.radio3)         #tab
        self.Controls.Add(self.label3)         #"(Defaults to comma delimited)"
        
        self.Controls.Add(self.label_submit)   #"",error / compleation messages
        self.Controls.Add(button)              #filename
        self.Controls.Add(button2)             #layout
        self.Controls.Add(button_submit)       #submit
        

#----------------------------------------------------------------
    # delimiter radio buttons changed
#----------------------------------------------------------------
    def checkedChanged(self, sender, args):
        global DELIM
        if not sender.Checked:
            DELIM = ','
        if sender.Text == 'Pipe':
            DELIM = '|'
        elif sender.Text == 'Tab':
            DELIM = '\t'
        else:
            if sender.Text == 'Comma':
                DELIM = ','
        
#----------------------------------------------------------------
    # filename in layout radio buttons changed
#----------------------------------------------------------------
    def HeaderCheckedChanged(self, sender, args):
        global HEADER
        if not sender.Checked:
            HEADER = True
        else:
            HEADER = False

#----------------------------------------------------------------
    # filename button handler
#----------------------------------------------------------------
    def buttonPressed(self, sender, args):
        global FILENAME
        dialogf = OpenFileDialog()
        if dialogf.ShowDialog() == DialogResult.OK:
            FILENAME = dialogf.FileName
            print "FILENAME: ",FILENAME
            self.label_filename.Text = FILENAME
            self.Controls.Add(self.label_filename) 
        else:
            print "No file selected"
            
#----------------------------------------------------------------
    # layout button handler
#----------------------------------------------------------------
    def button2Pressed(self, sender, args):
        global LAYOUT
        dialogl = OpenFileDialog()
        if dialogl.ShowDialog() == DialogResult.OK:
            LAYOUT = dialogl.FileName
            print "LAYOUT: " + dialogl.FileName
            self.label_layout.Text = LAYOUT
            self.Controls.Add(self.label_layout) 
        else:
            print "No file selected"
            
#----------------------------------------------------------------
    # submit button handler
#----------------------------------------------------------------
    def button_submitPressed(self, sender, args):
        if FILENAME and LAYOUT:
            print "HEADER = " + str(HEADER)
            if convert(FILENAME,LAYOUT,DELIM,HEADER):
                self.label_submit.ForeColor = Color.White
                self.label_submit.Text = "file has been converted!"
            else:
                self.label_submit.ForeColor = Color.DarkRed
                self.label_submit.Text = "ERROR: Layout has field names, please select Yes for field names"
        elif not FILENAME:
            self.label_submit.ForeColor = Color.DarkRed
            self.label_submit.Text = "ERROR: You must select a file name"
        else:
            self.label_submit.ForeColor = Color.DarkRed
            self.label_submit.Text = "ERROR: You must select a layout file"
            
#----------------------------------------------------------------
#----------------------------------------------------------------

form = HelloWorldForm()
Application.Run(form)