import re
from tkinter import Tk, Label, Button
from tkinter.filedialog import askopenfilename, askopenfilenames

import csv
import codecs
from datetime import datetime, time

def time_diff(start, end):
    if start <= end: # e.g., 10:33:26-11:15:49
        return end - start
    else: # end < start e.g., 23:55:00-00:25:00
        end += time(1) # +day
        assert end > start
        return end - start
def opencsv():
    filez = askopenfilenames(parent=root,title='Choose a file',filetypes=(("csv files", "*.csv"),("All files","*.*")))
    filesnames = root.tk.splitlist(filez)
    mylabel2 = Label(root, text = "Following files have been opened, you can find the newly generated files in the same folder.").pack()
    for filename in filesnames:
        mylabel3 = Label(root, text = filename).pack()
        try:
            inputfile = csv.DictReader(codecs.open(filename, 'rU', 'utf-16'), delimiter='\t')
        except:
            inputfile = csv.DictReader(codecs.open(filename, 'r'), delimiter=',')        
        #bless https://stackoverflow.com/questions/7894856/line-contains-null-byte-in-csv-reader-python
        mydict = {} 
        maxtime = "1/1/1900, 12:59:59 PM"
        leavecount = 0
        maxtime = datetime.strptime(maxtime, '%m/%d/%Y, %I:%M:%S %p')

        for row in inputfile:
            x = datetime.strptime(row["Timestamp"], '%m/%d/%Y, %I:%M:%S %p')
            if maxtime < x:
                maxtime = x    
            if row["Full Name"] not in mydict.keys():
                mydict[row["Full Name"]] = (x, x, 0)
            else:
                if row["User Action"] =="Left":
                    mydict[row["Full Name"]] = (mydict[row["Full Name"]][0], x, mydict[row["Full Name"]][2]+1)
                elif (row["User Action"] =="Joined") or (row["User Action"] =="Joined before"):
                    mydict[row["Full Name"]] = (mydict[row["Full Name"]][0], mydict[row["Full Name"]][0], mydict[row["Full Name"]][2])


        for i,j in mydict.items():
            if j[0] == j[1]:
                j = (j[0], maxtime, j[2])
            timespent = time_diff(j[0],j[1])    
                
         
        new_filename = re.sub('\.csv', '_edited.csv', filename)

        with open(new_filename, 'w', newline='') as csvfile:
            fieldnames = ['Full Name', 'Time spent', 'Times left']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for i,j in mydict.items():
                if j[0] == j[1]:
                    j = (j[0], maxtime, j[2])
                timespent = time_diff(j[0],j[1])    
                #print ("{}: {} and {} leaves".format(i, timespent, j[2]))
                writer.writerow({'Full Name': i, 'Time spent': timespent, 'Times left': j[2]})    

root = Tk()
root.title("Attendance formatter for MS Teams")

message = "This is a tiny program to open .csv attendance files downloaded from microsoft teams\n \
and create a new csv file containing the names of all meeting participants, the length of time (first in last out) they \n \
were in the meeting, and the number of times they left mid-meeting."
mylabel = Label(root, text = message).pack()
mybutton =Button(root, text ="Choose csv files to convert.", command=opencsv).pack()

root.mainloop()