import os
import win32com.client as win32
import tkinter as tk
#from tkinter import *
from tkinter import Label
from tkinter import Entry
from tkinter import messagebox
from tkinter import Button 




   

root=tk.Tk()



width, height = root.winfo_screenwidth, root.winfo_screenheight

root.geometry("{0}x{1}+0+0".format(root.winfo_screenwidth(), root.winfo_screenheight()))
#root.resizable(0, 0)


root.wm_title("Graduating List and Outstanding")
#root.wm_iconbitmap('001.ico')
w= 300
l1=Label(root,text="        ")
l1.grid(row=0,column=0,padx= (w, 0)  )

l2=Label(root,text="HARVEST GRADUATION LIST AND OUTSTANDING COURSES", font='Helvetica 18 bold')
l2.grid(row=0,column=1,pady= (w-100, 0))

l3=Label(root,text="        ")
l3.grid(row=0,column=2)


l4=Label(root,text="Direct Entry:")
l4.grid(row=2,column=0,padx= (w, 0),pady= (w-250, 0))



l6=Label(root,text="        ")
l6.grid(row=1,column=2)

l7=Label(root,text="Normal Entry:")
l7.grid(row=3,column=0, padx= (w, 0), pady= (w-250, 0))



l9=Label(root,text="        ")
l9.grid(row=2,column=2)

l41=Label(root,text="   ")
l41.grid(row=6,column=1)

l51=Label(root,text="COURTSEY: Department of Mathematics")
l51.grid(row=8,column=1)


e1=Entry(root, width= 125, bg="#00FF00")

e1.grid(row=2,column=1,pady= (w-250, 0))
e1.get()  




e2=Entry(root, width= 125, bg="#00FF00")

e2.grid(row=3,column=1,pady= (w-250, 0))
e2.get()  

def hy():
    l41=Label(root,text="               ")
    
    l41=Label(root,text=e1.get())
    l41.grid(row=4,column=1)
    print(e1.get())


def run():
    RE=[]
    ReN=e1.get()
    RE.append(ReN)
    ReD=e2.get()
    RE.append(ReD)

    if RE[1] =="" or RE[0]=="":
            
        
        
        messagebox.showwarning("ERROR", "AN ENTRY IS EMPTY!")
        return 2
    

    
    xl = win32.Dispatch('Excel.Application')


    


    try:
        wb = xl.Workbooks.Open(os.path.join(os.getcwd(), 'comprehensiveresultlist.xlsx'))
                    
    except Exception:
        messagebox.showwarning("ERROR", "FILE NOT FOUND")
        messagebox.showwarning("ERROR", "FILE NAME SHOULD BE comprehensiveresultlist.xlsx")
        return "File not Found"

    try:
        ws_sheet1 = wb.Worksheets('SHEET1') 
                    
    except Exception:
        try:
            ws_sheet1 = wb.Worksheets('Sheet1')
        except Exception:
            try:
                ws_sheet1 = wb.Worksheets('comprehensiveresultlist')
            except Exception:
                messagebox.showwarning("ERROR", "SHEET NOT FOUND")
                

            return "Sheet not Found"
    
    xl.Visible = True
    ws_sheet1.Cells(1,"M").Value = "THE REAL OUTSATNDING COURSES"
    ws_sheet1.Cells(1,"N").Value = "NUMBER OF OUTSTANDING COURSES"
    ws_sheet1.Cells(1,"O").Value = "NUMBER OF PASSED COURSES"
    
    def spcomma(string):            
      
        
        list_string = string.split(',')
          
        return list_string
    
    def sp_(string):             
      
        
        list_string = string.split('-')
          
        return list_string
    
    
    
    def seperate(A):            
        
        for i in A:
            B = sp_(i)
            C1.append(B)
    
    
    
    
    def sepReg(A):           
        R=spcomma(A)
        Re2=R
        return Re2
        
              
    
    
    def notdone(A,B):          
        notdone=[]
        for i in A:
            if i not in B:
                
                notdone.append(i)
        return notdone
    
    
    
    def fail(A):                
        B=[]
        for i in A:
            if 'F' in i[1]:
                B.append(i[0])
        return B
    
    
    
    def listToString(s): 
           
        str1 = "" 
            
        for ele in s: 
            str1 += ele+','  
        
        return str1 
    
    def reEle(A):
        B=[]
          
        ele=''
        elec=spcomma(ele)
        for i in A:
            if i not in elec:
                B.append(i)
        return B
    
    def GP(R,U): 
    
        TR=0
        TUNIT=0
    
        for i in U:
            TUNIT=TUNIT+i[1]
        
        for i in R:
            for j in U:
                if i[0]==j[0]:
                    
                    
                    
                    if i[1]=='A':
                        Y=5
                    elif i[1]=='B':
                        Y=4
                    elif i[1]=='C':
                        Y=3
                    elif i[1]=='D':
                        Y=2
                    elif i[1]=='E':
                        Y=1
                    elif i[1]=='F':
                        Y=0
                    
                    x=Y*j[1]
                    TR=TR+x
                
        if TUNIT !=0:
            GP=TR/TUNIT 
            return GP
        else:
            return 0
    
    def gp4(c,reg1):  
        p=[]
        for i in c:        
            for j in reg1:
                if i[0]==j[0]:
                    p.append(j)
        return p
    
    def Su(R,U):
        TR=0
        for i in R:
            for j in U:
                if i[0]==j[0]:
                    TR=TR+j[1]
        return TR
    
    def Su1(R,U):
        TR=0
        for i in R:
            for j in U:
                if i==j[0]:
                    TR=TR+j[1]
        return TR
               
    def CCode(C1):
        
        for i in C1:
            if i[0]=="GST122":
                i[0]="GST203"
            if i[0]=="MTH142":
                i[0]="MTH102"
            if i[0]=="PHY132":
                i[0]="PHY102"
            if i[0]=="CHM132":
                i[0]="CHM102"
            
        
    
    def CPassed(passed):
        for i in passed:
            
            if i=="GST122":
                i="GST203"
            if i=="MTH142":
                i="MTH102"
            if i=="PHY132":
                i="PHY102"
            if i=="CHM132":
                i="CHM102"

    def save():
        try:
            wb.SaveAs(Filename=os.path.join(os.getcwd(), 'Graduation_List.xlsx'))
            
        except Exception:
            messagebox.showwarning("ERROR", "REMOVE Graduation_List")
            
            return 1
    
      
    n1=1
    n2=1
    n3=1
    while n1<10000:
        if  str(ws_sheet1.Cells(n2, 1))=="None":
            
            n1=10000
        else:
            n2+=1
            n3+=1
          
    
    
    
    for n in range(2,n2):
        C2=(ws_sheet1.Cells(n, 10))  
        C= str(C2)  
        C1=[]       
        
        passed=[]       
        
        
        D=(ws_sheet1.Cells(n, 5)) 
        DE=str(D)   
        J1=seperate(spcomma(C)) 
        J=spcomma(C)  

        for i in J: 
    
                sp_(i)
                passed.append(sp_(i)[0])
        if DE=="DE":
           Re=RE[0]
           
            
        else:
            Re=RE[1]
            
        Re2=sepReg(Re)  
        Elc=reEle(Re2)  
        
        passed=reEle(passed) 
        
        
                
        
        out=notdone(Elc, passed)        
        out2=listToString(out)
        ws_sheet1.Cells(n, "M").Font.color = 32768
        ws_sheet1.Cells(n,"M").Value = out2
        ws_sheet1.Cells(n,"N").Value = len(out)
        ws_sheet1.Cells(n,"O").Value = len(passed)
        
        if C1[0][0]=='None':
            ws_sheet1.Cells(n,"O").Value = 0
        
         
                
        # if Elc[len(Elc)-1] not in out:
        #     n3+=1
        #     ws_sheet1.Range(ws_sheet1.Cells(n,1),ws_sheet1.cells(n,18)).Copy(ws_sheet1.Range(ws_sheet1.Cells(n3,1),ws_sheet1.cells(n3,18))) 

        

       
    wb.Sheets.Add().Name="With Project"
    ws_sheet2 = wb.Worksheets('With Project')
    ws_sheet1.Range(ws_sheet1.Cells(1,1),ws_sheet1.cells(1,18)).Copy(ws_sheet2.Range(ws_sheet2.Cells(1,1),ws_sheet2.cells(1,18))) 
    ws_sheet1.Range(ws_sheet1.Cells(n2+1,1),ws_sheet1.cells(n3,18)).Copy(ws_sheet2.Range(ws_sheet2.Cells(2,1),ws_sheet2.cells(n3-n2+1,18))) 

        
    

    
    while save()==1:
        save()
    else:
        
        messagebox.showinfo("SUCCESS", "GRADUATION LIST CREATED")
        
           
        

    

B1=Button(root,text="GENERATE LIST", bg="skyblue", activebackground="green",command=run)
B1.grid(row=7,column=1)


root.mainloop()

