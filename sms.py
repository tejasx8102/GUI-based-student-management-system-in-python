#import libraries
from tkinter import *
import tkinter.ttk as ttk
import tkinter.messagebox as tkMessageBox
import sqlite3
import openpyxl
import re

#function to define database
def Database():
    global conn, cursor
    #creating student database
    conn = sqlite3.connect("student.db")
    cursor = conn.cursor()
    #creating STUD_REGISTRATION table
    cursor.execute(
        "CREATE TABLE IF NOT EXISTS STUD_REGISTRATION (STU_ID INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, STU_NAME TEXT, STU_CONTACT TEXT, STU_EMAIL TEXT, STU_ROLLNO TEXT, STU_BRANCH TEXT)")



#defining function for creating GUI Layout
def DisplayForm():
    #creating window
    display_screen = Tk()
    #setting width and height for window
    display_screen.geometry("1280x720")
    #setting title for window
    display_screen.title("GUI based student management system")
    global tree
    global SEARCH
    global name,contact,email,rollno,branch
    SEARCH = StringVar()
    name = StringVar()
    contact = StringVar()
    email = StringVar()
    rollno = StringVar()
    branch = StringVar()
    #creating frames for layout
    #topview frame for heading
    TopViewForm = Frame(display_screen, width=600, bd=1, relief=SOLID)
    TopViewForm.pack(side=TOP, fill=X)
    #first left frame for registration from
    LFrom = Frame(display_screen, width="350")
    LFrom.pack(side=LEFT, fill=Y)
    #seconf left frame for search form
    LeftViewForm = Frame(display_screen, width=500,bg="purple")
    LeftViewForm.pack(side=LEFT, fill=Y)
    #mid frame for displaying students record
    MidViewForm = Frame(display_screen, width=600)
    MidViewForm.pack(side=RIGHT)
    #label for heading
    lbl_text = Label(TopViewForm, text="Student Management System", font=('verdana', 18), width=600,bg="pink",fg="violet")
    lbl_text.pack(fill=X)
    #creating registration form in first left frame
    Label(LFrom, text="Name  ", font=("Arial", 12)).pack(side=TOP)
    Entry(LFrom,font=("Arial",10,"bold"),textvariable=name).pack(side=TOP, padx=10, fill=X)
    Label(LFrom, text="Contact ", font=("Arial", 12)).pack(side=TOP)
    Entry(LFrom, font=("Arial", 10, "bold"),textvariable=contact).pack(side=TOP, padx=10, fill=X)
    Label(LFrom, text="Email ", font=("Arial", 12)).pack(side=TOP)
    Entry(LFrom, font=("Arial", 10, "bold"),textvariable=email).pack(side=TOP, padx=10, fill=X)
    Label(LFrom, text="Rollno ", font=("Arial", 12)).pack(side=TOP)
    Entry(LFrom, font=("Arial", 10, "bold"),textvariable=rollno).pack(side=TOP, padx=10, fill=X)
    Label(LFrom, text="Branch ", font=("Arial", 12)).pack(side=TOP)
    Entry(LFrom, font=("Arial", 10, "bold"),textvariable=branch).pack(side=TOP, padx=10, fill=X)
    Button(LFrom,text="Submit",font=("Arial", 10, "bold"),command=register).pack(side=TOP, padx=10,pady=5, fill=X)

    #creating search label and entry in second frame
    lbl_txtsearch = Label(LeftViewForm, text="Enter name to Search", font=('verdana', 10),bg="gray")
    lbl_txtsearch.pack()

    # Create the search entry
    search = Entry(LeftViewForm, textvariable=SEARCH, font=('verdana', 15), width=10)
    search.pack(side=TOP, padx=10, fill=X)

    # Bind the SearchByName function to the KeyRelease event of the search entry
    search.bind('<KeyRelease>', SearchByName)

    # Creating view button
    btn_view = Button(LeftViewForm, text="View All", command=DisplayData)
    btn_view.pack(side=TOP, padx=10, pady=10, fill=X)

    # Creating reset button
    btn_reset = Button(LeftViewForm, text="Reset", command=Reset)
    btn_reset.pack(side=TOP, padx=10, pady=10, fill=X)

    # Creating delete button
    btn_delete = Button(LeftViewForm, text="Delete", command=Delete)
    btn_delete.pack(side=TOP, padx=10, pady=10, fill=X)

    # Creating export button
    btn_export = Button(LeftViewForm, text="Export to Excel", command=ExportToExcel)
    btn_export.pack(side=TOP, padx=10, pady=10, fill=X)

    # Bubble sort algorithm
    btn_sort = Button(LeftViewForm, text="Sort by Name", command=SortRecords)
    btn_sort.pack(side=TOP, padx=10, pady=10, fill=X)

    # Binary Search
    btn_search_name = Button(LeftViewForm, text="Search by Name", command=SearchByName)
    btn_search_name.pack(side=TOP, padx=10, pady=10, fill=X)


    #setting scrollbar
    scrollbarx = Scrollbar(MidViewForm, orient=HORIZONTAL)
    scrollbary = Scrollbar(MidViewForm, orient=VERTICAL)
    tree = ttk.Treeview(MidViewForm,columns=("Student Id", "Name", "Contact", "Email","Rollno","Branch"),
                        selectmode="extended", height=100, yscrollcommand=scrollbary.set, xscrollcommand=scrollbarx.set)
    scrollbary.config(command=tree.yview)
    scrollbary.pack(side=RIGHT, fill=Y)
    scrollbarx.config(command=tree.xview)
    scrollbarx.pack(side=BOTTOM, fill=X)
    #setting headings for the columns
    tree.heading('Student Id', text="Student Id", anchor=W)
    tree.heading('Name', text="Name", anchor=W)
    tree.heading('Contact', text="Contact", anchor=W)
    tree.heading('Email', text="Email", anchor=W)
    tree.heading('Rollno', text="Rollno", anchor=W)
    tree.heading('Branch', text="Branch", anchor=W)
    #setting width of the columns
    tree.column('#0', stretch=NO, minwidth=0, width=0)
    tree.column('#1', stretch=NO, minwidth=0, width=100)
    tree.column('#2', stretch=NO, minwidth=0, width=150)
    tree.column('#3', stretch=NO, minwidth=0, width=80)
    tree.column('#4', stretch=NO, minwidth=0, width=120)
    tree.pack()
    DisplayData()
#function to insert data into database


def register():
    Database()
    # Getting form data
    name1 = name.get()
    con1 = contact.get()
    email1 = email.get()
    rol1 = rollno.get().upper()  # Convert to uppercase for case-insensitive comparison
    branch1 = branch.get()

    # Applying empty validation
    if name1 == '' or con1 == '' or email1 == '' or rol1 == '' or branch1 == '':
        tkMessageBox.showinfo("Warning", "Fill in all the required fields.")
    else:
        # Validating contact number
        if not con1.isdigit() or len(con1) != 10:
            tkMessageBox.showinfo("Warning", "Contact number must be 10 digits long and contain only numbers.")
        else:
            # Validating roll number format
            if not re.match(r'^(20BCE|20bce)[0-5][0-9][0-9]$', rol1):
                tkMessageBox.showinfo("Warning", "Invalid roll number format. Roll number should start with '20BCE' or '20bce' followed by a three-digit number from 000 to 599.")
            else:
                # Execute query
                conn.execute('INSERT INTO STUD_REGISTRATION (STU_NAME,STU_CONTACT,STU_EMAIL,STU_ROLLNO,STU_BRANCH) \
                              VALUES (?,?,?,?,?)', (name1, con1, email1, rol1, branch1))
                conn.commit()
                tkMessageBox.showinfo("Message", "Stored successfully")
                # Refresh table data
                DisplayData()
        conn.close()


def Reset():
    #clear current data from table
    tree.delete(*tree.get_children())
    #refresh table data
    DisplayData()
    #clear search text
    SEARCH.set("")
    name.set("")
    contact.set("")
    email.set("")
    rollno.set("")
    branch.set("")


def Delete():
    #open database
    Database()
    if not tree.selection():
        tkMessageBox.showwarning("Warning","Select data to delete")
    else:
        result = tkMessageBox.askquestion('Confirm', 'Are you sure you want to delete this record?',
                                          icon="warning")
        if result == 'yes':
            curItem = tree.focus()
            contents = (tree.item(curItem))
            selecteditem = contents['values']
            tree.delete(curItem)
            cursor=conn.execute("DELETE FROM STUD_REGISTRATION WHERE STU_ID = %d" % selecteditem[0])
            conn.commit()
            cursor.close()
            conn.close()


#function to search data
def SearchRecord():
    #open database
    Database()
    #checking search text is empty or not
    if SEARCH.get() != "":
        #clearing current display data
        tree.delete(*tree.get_children())
        #select query with where clause
        cursor=conn.execute("SELECT * FROM STUD_REGISTRATION WHERE STU_NAME LIKE ?", ('%' + str(SEARCH.get()) + '%',))
        #fetch all matching records
        fetch = cursor.fetchall()
        #loop for displaying all records into GUI
        for data in fetch:
            tree.insert('', 'end', values=(data))
        cursor.close()
        conn.close()


#defining function to access data from SQLite database
def DisplayData():
    # Clear current data
    tree.delete(*tree.get_children())

    # Get data from the database
    data = GetDataFromDatabase()

    # Loop for displaying all data in GUI
    for row in data:
        tree.insert('', 'end', values=row)


def ExportToExcel():
    # Open the Excel workbook
    wb = openpyxl.Workbook()
    ws = wb.active

    # Add column headings
    ws.append(["Student Id", "Name", "Contact", "Email", "Rollno", "Branch"])

    # Get the data from the database
    data = GetDataFromDatabase()

    # Add data to the Excel sheet
    for row in data:
        ws.append(row)

    # Save the workbook
    wb.save("student_data.xlsx")

    # Show a message box to inform the user about the successful export
    tkMessageBox.showinfo("Export to Excel", "Student data has been exported to student_data.xlsx")


def GetDataFromDatabase():
    # Open database
    Database()

    # Select all data from the database
    cursor = conn.execute("SELECT * FROM STUD_REGISTRATION")

    # Fetch all data from the database
    fetch = cursor.fetchall()

    # Close the cursor and connection
    cursor.close()
    conn.close()

    return fetch


def BubbleSort(arr):
    n = len(arr)
    for i in range(n - 1):
        for j in range(0, n - i - 1):
            if arr[j][1] > arr[j + 1][1]:
                arr[j], arr[j + 1] = arr[j + 1], arr[j]


def SortRecords():
    # Get data from the database
    data = GetDataFromDatabase()

    # Sort the data using bubble sort algorithm based on student names
    BubbleSort(data)

    # Clear current data
    tree.delete(*tree.get_children())

    # Display sorted data in GUI
    for row in data:
        tree.insert('', 'end', values=row)


def BinarySearch(arr, target):
    left, right = 0, len(arr) - 1
    while left <= right:
        mid = left + (right - left) // 2
        name_parts = arr[mid][1].split()  # Splitting the name into parts
        if any(part.startswith(target) for part in name_parts):
            return mid
        elif arr[mid][1] < target:
            left = mid + 1
        else:
            right = mid - 1
    return -1


def SearchByName():
    # Get data from the database
    data = GetDataFromDatabase()

    # Sort the data using bubble sort algorithm based on student names
    BubbleSort(data)

    # Get the search target from the Entry widget
    target = SEARCH.get().lower()  # Convert to lowercase for case-insensitive search

    # Perform binary search to find matching student records
    matching_indices = []
    index = BinarySearch(data, target)
    while index != -1:
        matching_indices.append(index)
        index = BinarySearch(data[index + 1:], target)
        if index != -1:
            index = index + matching_indices[-1] + 1

    if matching_indices:
        # Clear current data
        tree.delete(*tree.get_children())
        # Display the search results in GUI
        for index in matching_indices:
            tree.insert('', 'end', values=data[index])
    else:
        tkMessageBox.showinfo("Search Result", "No matching students found.")


#calling function
DisplayForm()

if __name__=='__main__':
#Running Application
 mainloop()
