from tkinter import *
from tkinter import scrolledtext
from tkinter import messagebox
import xlsxwriter 

window = Tk()

window.title("Push To excel")

window.geometry('1000x500')

# number of col1 
lbl1 = Label(window, text="Number of rows in col1")
lbl1.grid(column=0, row=0)
num1 = Entry(window,width=10)
num1.grid(column=1, row=0)
# number of col1 
lbl2 = Label(window, text="Number of rows in col2")
lbl2.grid(column=0, row=1)
num2 = Entry(window,width=10)
num2.grid(column=1, row=1)
# number of col1 
lbl3 = Label(window, text="Number of rowsin col3")
lbl3.grid(column=0, row=2)
num3 = Entry(window,width=10)
num3.grid(column=1, row=2)
column1 = scrolledtext.ScrolledText(window,width=20,height=20)
column1.grid(column=10,row=40)
     # column2
column2 = scrolledtext.ScrolledText(window,width=20,height=20)
column2.grid(column=40,row=40)
     # column3
column3 = scrolledtext.ScrolledText(window,width=40,height=20)
column3.grid(column=80,row=40)
def toExcel():

	rownum1 = int(num1.get())
	rownum2 = int(num2.get())
	rownum3 = int(num3.get())
	# messagebox.showinfo(rownum)
	couloumn1 = []
	couloumn2 = []
	couloumn3 = [[]]
	couloumn3 = []
	# array 1
	for x in range(rownum1):
		couloumn1.append(column1.get(str(x+1) + ".0",str(x+2) + ".0"))

	# for r in couloumn1:
		# print(r)
    # array 2
	for x in range(rownum2):
		couloumn2.append(column2.get(str(x+1) + ".0",str(x+2) + ".0"))

	# for r in couloumn2:
		# print(r)

    # array 3
	for x in range(rownum3):
		couloumn3.append([])
		# string  column3.get(str(x+1) + ".0",str(x+2) + ".0")
		row = column3.get(str(x+1) + ".0",str(x+2) + ".0")
		# ###########
		arr = row.split(' ')
		for z in arr: # we have array of every row
			couloumn3[x].append(z)
	for i in range(len(couloumn3)):
		for j in range(len(couloumn3[i])):
			print(couloumn3[i][j])
	# push to excell
	workbook = xlsxwriter.Workbook('Data.xlsx') 
  
# The workbook object is then used to add new  
# worksheet via the add_worksheet() method. 
	worksheet = workbook.add_worksheet() 
# push to excel 
	count = 1 
	string =""
	array = []
	count2 = 1
	for x in couloumn1:
	    for y in couloumn2:
	    	for i in range(len(couloumn3)):
	    		for j in range(len(couloumn3[i])):
	    			if j == len(couloumn3[i])-1:
	    				string+= str(couloumn3[i][j])  
	    			else: 
	    				string+= str(couloumn3[i][j]) + "+"
    				array.append(couloumn3[i][j])
	    		array.insert(0,x)
	    		array.insert(1,y)
	    		for v in array:
	    			worksheet.write('B' + str(count2), v)
	    			print(v)
	    			count2 +=1 ;
	    		worksheet.write('A' + str(count), x + "+"  + y + "+" + string)
	    		count += len(array) 
    			array = []
    			string = ""
    	
    	


   
	workbook.close() 
	messagebox.showinfo('Excel','Updated')

getButton = Button(window, text="Excel" , command=toExcel)
getButton.grid(column=5, row=50)

# function to add number of rows



  


def showCols():

  
    lbl1.configure(text="col1")
    lbl2.configure(text="col2")
    lbl3.configure(text="col3")
    num1.config(state=DISABLED)
    num2.config(state=DISABLED)
    num3.config(state=DISABLED)
    # column1
   
    

show = Button(window, text="show", command=showCols)

show.grid(column=2, row=3)

window.mainloop()
