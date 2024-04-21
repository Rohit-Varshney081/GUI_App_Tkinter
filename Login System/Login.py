from tkinter import *
from tkinter import messagebox
import mysql.connector as sql

background="#06283D"
framebg="#EDEDED"
framefg="#06283D"

global trial_no
trial_no=0
def trial():
    global trial_no
    trial_no+=1
    print("trial no is ",trial_no)
    if trial_no==3:
        messagebox.showwarning("Warning","You have more than limit !!")
        root.destroy()  #program close

def loginuser():
    username=user.get()
    password=code.get()
    print("Username entered :", username)
    print("password entered :", password)
    if (username=="" or username=="UserID")or(password=="" or password=="Password"):
        messagebox.showerror("Entry error","Type username or password!!")

    else:
        try:
            conn=sql.connect(host='localhost',user='root',password='Usha@081')
            if conn.is_connected():
                print("Successfully Connected !!")
            cursorobj=conn.cursor()
            try:
                #........Create a database........
                db="create database studentregistration"
                cursorobj.execute(db)
                print("Database Created !")
            except:
                print("Database Already Exists.....")
        except:
            messagebox.showerror("Connection error")
        
        cd="use studentregistration"
        cursorobj.execute(cd)

        cd1="select * from login where Username=%s and Password=%s"
        cursorobj.execute(cd1,(username,password))
        myresult=cursorobj.fetchone()
        print(myresult)

        if  myresult==None:
            messagebox.showinfo("Invalid userID and password !!")
            # but user can try many times and crack password so lets make that user can tryu only upto 3 times
            trial()
        else:
            messagebox.showinfo("Login Successfully !!")
            root.destroy()
            import details2

def register():
    root.destroy()
    import register

    
root=Tk()
root.title("Login System")
root.geometry("1250x700+210+100")
root.config(bg=background)
root.resizable(False,False)


#icon image
image_icon=PhotoImage(file="Image/icon.png")
root.iconphoto(False,image_icon)

#background image
frame=Frame(root,bg="red")
frame.pack(fill=Y)
backgroundimage=PhotoImage(file="Image/LOGIN.png")
Label(frame,image=backgroundimage).pack()

#user entry
def user_enter(e):
    user.delete(0,"end")
def user_leave(e):
    if user.get()=='':
        user.insert(0,"UserID")
        
user=Entry(frame,width=18,fg="#fff",border=0,bg="#375174",font=("Arial Bold",24))
user.insert(0,"UserID")
user.bind("<FocusIn>",user_enter)
user.bind("<FocusOut>",user_leave)
user.place(x=500,y=315)

#password entry
def password_enter(e):
    code.delete(0,"end")
def password_leave(e):
    pwd=code.get()
    if pwd=='':
        code.insert(0,"Password")
        
code=Entry(frame,width=18,fg="#fff",border=0,bg="#375174",font=("Calibari Bold",24))
code.insert(0,"Password")
code.bind("<FocusIn>",password_enter)
code.bind("<FocusOut>",password_leave)
code.place(x=500,y=410)

#Hide and show button
button_mode=True
def hide():
    global button_mode
    if button_mode:
        eyebutton.config(image=closeeye,activebackground="white")
        code.config(show="*")
        button_mode=False
    else:
        eyebutton.config(image=openeye,activebackground="white")
        code.config(show="")
        button_mode=True
        
openeye=PhotoImage(file='Image/openeye.png')
closeeye=PhotoImage(file='Image/close eye.png')
eyebutton=Button(frame,image=openeye,bg="#375174",bd=0,command=hide)
eyebutton.place(x=780,y=410)

#Login button

loginButton=Button(root,text="LOGIN",bg="#1f5675",fg="white",width=10,height=1,font=("arial",16,"bold"),bd=0,command=loginuser)
loginButton.place(x=570,y=600)

label=Label(root,text="Don't have an account ?",fg="#fff",bg="#00264d",font=("Microsoft YaHei UI Light",9))
label.place(x=500,y=500)

registerButton=Button(root,width=10,text="add new user",border=0,bg="#00264d",cursor="hand2",fg="#57a1f8",command=register)
registerButton.place(x=650,y=500)


root.mainloop()
