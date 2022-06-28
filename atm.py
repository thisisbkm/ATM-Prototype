import xlrd
import xlwt
from time import sleep
from xlutils.copy import copy
from tkinter import *
from tkinter import messagebox
from tkinter.simpledialog import askinteger
from tkinter.simpledialog import askstring
from datetime import datetime

def load():
    global book
    book=xlrd.open_workbook("/media/bkm/F014230D1422D67E/MACHINE LEARNING-TECHVANTO/ATM-PROTOTYPE/customers_data.xls")
    global sheet
    sheet=book.sheet_by_index(0)  
      
def start(f):
    load()
    def check(e):
        if(accNo.get() and passwd.get()):
            try:
                acn = int(accNo.get())
                psd = int(passwd.get())
            except:
                messagebox.showwarning("Warning","Account Number and Pin must be Integers")
            for i in range(sheet.nrows):
                if(sheet.cell_value(i,0)==acn):
                    if(sheet.cell_value(i,3)==psd):
                        accNo.delete(0,END)
                        passwd.delete(0,END)
                        mainscreen(i,f)
                    else:
                        messagebox.showwarning("Incorrect Pin", "Entered Pin Is Invalid")
                    return
            messagebox.showinfo("Invalid Account","Entered Account Dosen't Exist !")
        else:
            messagebox.showinfo("Info", "All feilds are mandatory !!",parent=f)

    f.geometry("500x500")
    f.resizable(0,0)
    f.title("ATM")
    
    # mainscreen(1, Toplevel(f))
    # return
    Label(f,text="Account Number : ").place(x=114,y=150)
    accNo=Entry(f)
    accNo.place(x=250,y =150)

    Label(f,text="Pin : ").place(x=200,y =200)
    passwd=Entry(f,show="*")
    passwd.place(x=250,y =200)
    
    # accNo.insert(0,"1200001")
    # passwd.insert(0, "8989")

    Button(f,text="Submit",
        command=lambda:check(None),
        activebackground="Green",
        activeforeground="White",
        background="Green",
        foreground="White"
        ).place(x=350,y =250)
    f.bind("<Return>",check)

    Button(
        f,text="Clear",
        activebackground="Red",activeforeground="White",
        background="Red",foreground="White",
        command=lambda:accNo.delete(0,END) and passwd.delete(0,END)
        ).place(x=250,y=250)

    Button(
        f,text = "Create Account",
        background="Yellow",activebackground="Yellow",
        command=lambda:createAcc(f)
        ).place(anchor=CENTER,x=400,y=450)


def mainscreen(row,s):
    s.withdraw()
    f=Toplevel(s)

    def deposit():
        load()
        amount=askinteger("Deposit","Amount to Deposit")

        if(amount<=0):
            messagebox.showwarning("Negative Amount","Amount should be greater than 0")
            return
        with open("log.txt",'a') as fil:
            fil.write(str(datetime.now().strftime("[%d-%m-%Y %H:%M:%S] ")) + str(int(sheet.cell_value(row,0)))+ " deposited "+str(amount)+"\n")
        amount+=sheet.cell_value(row,2)
        wb=copy(book)
        wsheet = wb.get_sheet(0)
        wsheet.write(row,2,amount)
        wb.save('customers_data.xls')
        messagebox.showinfo("","Amount Deposited Successfully\n Current Balance {}".format(amount))
    
    def withdraw():
        load()
        amount=askinteger("Withdraw","Amount to Withdraw")

        if(amount<=0):
            messagebox.showwarning("Negative Amount", "Amount should be greater than 0")
            return

        balance = sheet.cell_value(row,2)
        if(amount>balance):
            messagebox.showerror("","Insufficient Balcance")
            return

        wb=copy(book)
        wsheet = wb.get_sheet(0)
        wsheet.write(row,2,balance-amount)
        wb.save('customers_data.xls')
        with open("log.txt",'a') as fil:
            fil.write(str(datetime.now().strftime("[%d-%m-%Y %H:%M:%S] ")) + str(int(sheet.cell_value(row,0)))+ " withdrawed "+str(amount)+"\n")
        messagebox.showinfo("","Amount Withdrawn Successfully\n Current Balance {}".format(balance-amount))
    
    def balance():
        load()
        messagebox.showinfo("","Account Balance : {}".format(sheet.cell_value(row,2)))

    def changepin():
        load()
        newpin=askstring("Pin Change", "Enter the New Pin",show='*')

        if(newpin==None or newpin==""):
            return
        rnewpin=askstring("Pin Change", "Re-Enter the Pin",show='*')

        if(newpin!=rnewpin):
            messagebox.showwarning("","Pin Don't Match")
            return

        if(len(newpin)!=4):
            messagebox.showinfo("Invalid Length","Pin should be of 4 Digts")
            return

        wb=copy(book)
        wsheet = wb.get_sheet(0)
        wsheet.write(row,3,int(newpin))
        wb.save('customers_data.xls')
        messagebox.showinfo("","Pin Change Successful !")
        load()

    def mtransfer():
        load()
        acno = askinteger("Transfer", "Enter Account Number To Transfer")
        if(acno==sheet.cell_value(row,0)):
            messagebox.showwarning("","You can't transfer to Your account !")
            return

        for i in range(sheet.nrows):
            if(sheet.cell_value(i,0)==acno):
                amt = askinteger("Transfer","Enter amount you wish to transfer")
                if(amt==None):
                    return
                if(amt>sheet.cell_value(row,2)):
                    messagebox.showerror("Insufficient Funds","Your account don't have enough balance")
                elif(amt>0):
                    if(messagebox.askyesno("Confirmation", "Are you sure to transfer ?")):
                        wb=copy(book)
                        wsheet=wb.get_sheet(0)
                        wsheet.write(i,2,sheet.cell_value(i,2)+amt)
                        wsheet.write(row,2,sheet.cell_value(row,2)-amt)
                        wb.save('customers_data.xls')
                        load()
                        messagebox.showinfo("Transfer","Amount Transferred Successfully !")
                        with open("log.txt",'a') as fil:
                            fil.write(str(datetime.now().strftime("[%d-%m-%Y %H:%M:%S] ")) + str(int(sheet.cell_value(row,0)))+ " transferred "+str(amt)+" to "+str(int(sheet.cell_value(i,0)))+"\n")
                        return
        if(acno!=None):
            messagebox.showwarning("Invalid Account", "No Such Account")

    f.grab_set()
    f.geometry("600x600")
    f.resizable(0,0)
    f.title("Main Menu")
    f.protocol("WM_DELETE_WINDOW",False)
    
    Label(f,text=("Welcome "+ sheet.cell_value(row,1)).upper(),font=("Times New Roman",15),fg="Blue").place(anchor=CENTER,x=300,y=20)
    Button(f,text="DEPOSIT",width=15,height=4,command=deposit).place(x=150,y=80)
    Button(f,text="WITHDRAW",width=15,height=4,command=withdraw).place(x=350,y=80)
    Button(f,text="CHANGE PIN",width=15,height=4,command=changepin).place(x=150,y=240)
    Button(f,text="BALANCE ENQUIRY",width=15,height=4,command=balance).place(x=350,y=240)
    Button(f,text="MONEY TRANSFER",width=15,height=4,command=mtransfer).place(x=150,y=400)
    Button(f,text="END TRANSACTION",width=15,height=4,
        activebackground="yellow",background="yellow",
        relief="ridge",borderwidth=4,
        command=lambda:(f.destroy(),s.deiconify())
        ).place(x=350,y=400)



def createAcc(s):
    s.withdraw()
    f=Toplevel(s)

    def addEntry(e):
        if(not(name.get() and pin.get() and rpin.get() and amount.get())):
            messagebox.showinfo("Info","All Feilds are Mandatory !!",parent=f)
            return
        

        try:
            p=int(pin.get())
            rp=int(rpin.get())
            amt = int(amount.get())

            if(p!=rp):
                messagebox.showerror("Error","Pin Don't Match",parent=f)
            elif(len(pin.get())!=4):
                messagebox.showerror("Error","Entered Pin should be 4 Digits",parent=f)
            else:
                wb=copy(book)
                wsheet = wb.get_sheet(0)
                rows = sheet.nrows
                nac=sheet.cell_value(sheet.nrows-1,0)+1
                wsheet.write(rows,0,int(nac))
                wsheet.write(rows,1,name.get())
                wsheet.write(rows,2,amt)
                wsheet.write(rows,3,p)
                wb.save('customers_data.xls')
                messagebox.showinfo("Info","Account Details Saved\nYour Account Number : {}".format(nac),parent=f)
                load()
                with open("log.txt",'a') as fil:
                    fil.write(str(datetime.now().strftime("[%d-%m-%Y %H:%M:%S]"))+ " New User "+ str(nac)+ " with balance "+str(sheet.cell_value(sheet.nrows-1,2))+"\n")
                f.destroy()
                s.deiconify()
        
        except Exception as e:
            print(e)
            messagebox.showwarning("Warning","Pin and Amount should be Integers",parent=f)
    

    f.grab_set()
    f.geometry("500x400")
    f.resizable(0,0)
    f.protocol("WM_DELETE_WINDOW",False)
    f.title("Create Account")
    f.bind("<Return>",addEntry)


    Label(f,text="Enter Your Name ",).place(x=20,y=50)
    name=Entry(f,width=30)
    name.place(x=150,y=50)

    Label(f,text="Enter Your Pin ",).place(x=20,y=80)
    pin=Entry(f,width=30,show="*")
    pin.place(x=150,y=80)

    Label(f,text="Re-Enter Pin ").place(x=20,y=110)
    rpin=Entry(f,width=30,show="*")
    rpin.place(x=150,y=110)

    Label(f,text="Amount to Deposit").place(x=20,y=140)
    amount=Entry(f,width=30)
    amount.insert(0,"0")
    amount.place(x=150,y=140)

    Button(f,text="Submit",activeforeground="White",activebackground="Green",background="Green",foreground="White",command=lambda:addEntry(None)).place(anchor=CENTER,x=300,y =200)
    
    Button(
        f,text="Clear",
        activebackground="Red",activeforeground="White",
        background="Red",foreground="White",
        command=lambda:name.delete(0,END) and pin.delete(0,END) and rpin.delete(0,END) and amount.delete(0,END)
        ).place(anchor=CENTER,x=200,y=200)
    
    Button(f,text="Cancel",activebackground="Yellow",background="Yellow",command=lambda:(f.destroy(), s.deiconify())).place(anchor=CENTER,x=430,y=350)


root = Tk()
start(root)
root.mainloop()