import Tkinter as tk
import tkFileDialog
import os
import validate

file_names = [0,0]

#get file name & path, store in file_names and put in Entry widget
def getFileName(entryName, element):
    
    global file_names
    file_names[element -1] = tkFileDialog.askopenfilename()

    entryName.delete(0,'end')
    entryName.insert(0, os.path.basename(file_names[element -1]))

    #print file_names

#button to triger validate.py logic
def button_click(get):
    global unique_key

    get_nospace = get.replace(" ", "")
    unique_key = get.split(',')


    
    print "Unique Key: "
    print unique_key

    validate.main(file_names,unique_key)
    get_text(validate.text_container)


#return strings to be printed
def get_text(string):

	
	text_box.config(state='normal')
	text_box.insert('end', string)
	text_box.config(state='disabled')




#root window
root = tk.Tk()
root.title("Post-Load Validation")
root.geometry('330x190')




label1  = tk.Label(root, text = 'Extract File')
label2  = tk.Label(root, text = 'DCW Sheet File')
label3  = tk.Label(root, text = 'Unique Field in DCW')
entry1 = tk.Entry(root)
entry2 = tk.Entry(root)
entry3 = tk.Entry(root)
button1 = tk.Button(text = "Open ", command = (lambda: getFileName(entry1, 1)))
button2 = tk.Button(text = "Open ", command = (lambda: getFileName(entry2, 2)))
button3 = tk.Button(text = "Check", command = (lambda: button_click(entry3.get())))
text_box = tk.Text(root, wrap = 'word')
text_box.config(state='disabled', width =2 , height = 6)
S = tk.Scrollbar(root, command = text_box.yview)


entry1.insert(0, "Extract.xlsx")
entry2.insert(0, "DCW.xlsx")
entry3.insert(0, "MATNR")

#pack to position object instance
label1.grid(row = 0, column = 1, padx= 3)
entry1.grid(row = 0, column = 2)
button1.grid(row = 0, column = 3, padx= 3)
label2.grid(row = 1, column = 1, padx= 3)
entry2.grid(row = 1, column = 2)
button2.grid(row = 1, column = 3, padx= 3)
label3.grid(row = 2, column = 1, padx= 3)
entry3.grid(row = 2, column = 2)
button3.grid(row = 2, column = 3, padx= 3)
text_box.grid(row = 3, column = 1, columnspan =3 , sticky = "we", padx= 5)
S.grid (row =3, column = 4, sticky = 'nsw')
text_box['yscrollcommand'] = S.set

root.mainloop()
