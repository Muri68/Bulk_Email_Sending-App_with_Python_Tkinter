from tkinter import *
from PIL import ImageTk # pip install pillow ---- to install PILLOW
from tkinter import messagebox, filedialog
import os
import pandas as pd #pip install pandas
import email_function
import time

class Bulk_Email:
	def __init__(self, root):
		self.root = root
		self.root.title("Bulk Email Application")
		self.root.geometry("1100x600+200+50")
		self.root.resizable(False, False)
		self.root.config(bg="#ffffff")

		# Icons
		self.setting_icon = ImageTk.PhotoImage(file="images/setting.png")

		# Head Title --- image=self.setting_icon, compound=RIGHT, 
		title = Label(self.root, text="BULK EMAIL SENDING APP", padx=10, font=("Algerian",48,"bold"), bg="#2f303a", fg="#ffffff", anchor="w").place(x=0,y=0,relwidth=1)
		
		self.desc = Label(self.root, text="Use Excell file to send Bulk Email at once, Ensure the Email Column in the Excell sheet is Named 'Email'", padx=10, font=("Century Gothic",15,"bold"), bg="#ffffff", fg="#ffffff")
		self.desc.place(x=0,y=77,relwidth=1)

		# Setting Button
		btn_setting = Button(self.root, image=self.setting_icon, bd=0, activebackground="#2f303a", bg="#2f303a", cursor="hand2", command=self.setting_window).place(x=1010, y=5)

		# RADIO BTN DEFAULT SELECT
		self.var_choice = StringVar()
		self.var_choice.set("single")

		# SINGLE EMAIL BTN===========
		single = Radiobutton(self.root, text="Single", value="single", command=self.check_single_or_bilk, variable=self.var_choice, font=("Century Gothic",30,"bold"), bd=0, activebackground="#ffffff", bg="#ffffff", fg="#262626").place(x=50, y=120)
		
		# BULK EMAIL BTN===========
		bulk = Radiobutton(self.root, text="Bulk", value="bulk", command=self.check_single_or_bilk, variable=self.var_choice, font=("Century Gothic",30,"bold"), bd=0, activebackground="#ffffff", bg="#ffffff", fg="#262626").place(x=250, y=120)

		# Label of TEXT AREA =======
		to = Label(self.root, text="To (Receiver's Email Address)", font=("Century Gothic",18,"bold"),bg="#ffffff").place(x=50, y=220)
		subject = Label(self.root, text="Subject ", font=("Century Gothic",18,"bold"),bg="#ffffff").place(x=50, y=280)
		message = Label(self.root, text="Message ", font=("Century Gothic",18,"bold"),bg="#ffffff").place(x=50, y=340)

		# TEXT AREA ======
		self.txt_to = Entry(self.root, font=("Century Gothic",16), bg="lightgray")
		self.txt_to.place(x=420, y=220, width=350, height=35)

		self.txt_subject = Entry(self.root, font=("Century Gothic",16), bg="lightgray")
		self.txt_subject.place(x=420, y=280, width=500, height=35)

		self.txt_message = Text(self.root, font=("Century Gothic",16), bg="lightgray")
		self.txt_message.place(x=420, y=340, width=650, height=100)

		# BROWSE BTN FOR BULK EMAIL
		self.btn_browse = Button(self.root, command=self.browse_file,  font=("Century Gothic",18,"bold"), bd=0, cursor="", bg="#ffffff", fg="#ffffff", activebackground="#ffffff", activeforeground="#ffffff", state=DISABLED)
		self.btn_browse.place(x=800, y=220, width=130, height=35)

		# Bottom Buttons
		btn_clear = Button(self.root, text="CLEAR", command=self.clear1, font=("Century Gothic",18,"bold"), bg="#2f303a", fg="#ffffff", activebackground="#2f303a", activeforeground="#ffffff", cursor="hand2").place(x=800, y=470, width=130, height=30)
		btn_send = Button(self.root, command=self.send_email, text="SEND", font=("Century Gothic",18,"bold"), bg="#32abe6", fg="#ffffff", activebackground="#32abe6", activeforeground="#ffffff", cursor="hand2").place(x=940, y=470, width=130, height=30)
		self.checking_if_login_file_exist()

		# STATUS AREA =======
		self.total = Label(self.root, font=("Century Gothic",18,"bold"),bg="#ffffff")
		self.total.place(x=50, y=550)

		self.sent = Label(self.root, font=("Century Gothic",18,"bold"), fg="#00f23d", bg="#ffffff")
		self.sent.place(x=400, y=550)

		self.left = Label(self.root, font=("Century Gothic",18,"bold"), fg="#32abe6", bg="#ffffff")
		self.left.place(x=600, y=550)

		self.failed = Label(self.root, font=("Century Gothic",18,"bold"), fg="#ac2a2a", bg="#ffffff")
		self.failed.place(x=800, y=550)
		# message = Label(self.root, text="Message ", font=("Century Gothic",18,"bold"),bg="#ffffff").place(x=50, y=340)



	# BROWSE FILE FUNCTION
	def browse_file(self):
		op = filedialog.askopenfile(initialdir='/', title="Select Excel File", filetypes=(("All Files", "*.*"),("Excel Files", ".xlsx")))
		if op!=None:
			data = pd.read_excel(op.name)
			if 'Email' in data.columns:
				self.emails = list(data['Email'])
				c = []
				for i in self.emails:
					if pd.isnull(i)==False:
						c.append(i)

				self.emails = c
				if len(self.emails)>0:
					self.txt_to.config(state=NORMAL)
					self.txt_to.delete(0, END)
					self.txt_to.insert(0, str(op.name.split("/")[-1]))
					self.txt_to.config(state='readonly')
					self.total.config(text="Total Emails Found: " + str(len(self.emails)))
					self.sent.config(text="Sent: ")
					self.left.config(text="Left: ")
					self.failed.config(text="Failed: ")
				else:
					messagebox.showerror("Error", "This file does not have any Emails in it", parent=self.root)
		
			else:
				messagebox.showerror("Error", "Please select Excel file that have Column named Email", parent=self.root)
		




	def send_email(self):
		x = len(self.txt_message.get('1.0', END))
		
		if self.txt_to.get()=="" or self.txt_subject.get()=="" or x==1:
			messagebox.showerror("Error", "All field are required", parent=self.root)
		else:
			if self.var_choice.get()=="single":
				status = email_function.email_sent_func(self.txt_to.get(), self.txt_subject.get(), self.txt_message.get('1.0', END), self.email_,self.password_)
				if status=="s":
					messagebox.showinfo("Success", "Email sent Successfully", parent=self.root)
				if status=="f":
					messagebox.showerror("Failed", "Email Not Send pleace Try Again", parent=self.root)

			if self.var_choice.get()=="bulk":
				self.fail=[]
				self.s_count = 0
				self.f_count = 0
				for x in self.emails:
					status = email_function.email_sent_func(x, self.txt_subject.get(), self.txt_message.get('1.0', END), self.email_,self.password_)
					if status == "s":
						self.s_count+=1
					if status == "f":
						self.f_count+=1
					self.status_bar()
					#time.sleep(1)
				messagebox.showinfo("Success", "Email sent Successfully", parent=self.root)


	def status_bar(self):
		self.total.config(text="STATUS: " + str(len(self.emails))+"=>>")
		self.sent.config(text="Sent: " + str(self.s_count))
		self.left.config(text="Left: "+ str(len(self.emails)-(self.s_count + self.f_count)))
		self.failed.config(text="Failed: " + str(self.f_count))
		self.total.update()
		self.sent.update()
		self.left.update()
		self.failed.update()


	def check_single_or_bilk(self):
		if self.var_choice.get()=="single":
			self.btn_browse.config(state=DISABLED,text="", bd=0, bg="#ffffff", fg="#ffffff", activebackground="#ffffff", activeforeground="#ffffff")
			self.desc.config(bg="#ffffff", fg="#ffffff")
			self.txt_to.config(state=NORMAL)
			self.txt_to.delete(0, END)
			self.clear1()
	
		if self.var_choice.get()=="bulk":
			self.btn_browse.config(state=NORMAL, text="BROWSE", bg="#32abe6", fg="#ffffff", activebackground="#32abe6", activeforeground="#ffffff", cursor="hand2",)
			self.desc.config(bg="#32abe6", fg="#262626")
			self.txt_to.config(state='readonly')
			self.txt_to.delete(0, END)

	def clear1(self):
		self.txt_to.config(state=NORMAL)
		self.txt_to.delete(0, END)
		self.txt_subject.delete(0, END)
		self.txt_message.delete('1.0', END)
		self.var_choice.set("single")
		self.btn_browse.config(state=DISABLED,text="", bd=0, bg="#ffffff", fg="#ffffff", activebackground="#ffffff", activeforeground="#ffffff")
		self.desc.config(bg="#ffffff", fg="#ffffff")
		self.total.config(text="")
		self.sent.config(text="")
		self.left.config(text="")
		self.failed.config(text="")





	def setting_window(self):
		self.checking_if_login_file_exist()
		self.root2=Toplevel()
		self.root2.title("Setting")
		self.root2.geometry("700x450+350+90")
		self.root2.resizable(False, False)
		self.root2.config(bg="#ffffff")
		self.root2.focus_force()
		self.root2.grab_set()

		title = Label(self.root2, text="Credencials Setting", image=self.setting_icon, padx=10, compound=LEFT, font=("Algerian",28,"bold"), bg="#2f303a", fg="#ffffff", anchor="w").place(x=0,y=0,relwidth=1)
		
		desc = Label(self.root2, text="Login to your Email Address to anable you Send Mails ", padx=10, font=("Century Gothic",15,"bold"), bg="#32abe6", fg="#262626").place(x=0,y=77,relwidth=1)

		# Label of TEXT AREA =======
		Email = Label(self.root2, text="Email", font=("Century Gothic",18,"bold"),bg="#ffffff").place(x=50, y=220)
		password = Label(self.root2, text="Password ", font=("Century Gothic",18,"bold"),bg="#ffffff").place(x=50, y=280)
		
		# TEXT AREA ======
		self.txt_email = Entry(self.root2, font=("Century Gothic",16), bg="lightgray")
		self.txt_email.place(x=220, y=220, width=350, height=35)

		self.txt_password = Entry(self.root2, show="*", font=("Century Gothic",16), bg="lightgray")
		self.txt_password.place(x=220, y=280, width=350, height=35)
		
		# Bottom Buttons
		btn_clear2 = Button(self.root2, command=self.clear2, text="CLEAR", font=("Century Gothic",18,"bold"), bg="#2f303a", fg="#ffffff", activebackground="#2f303a", activeforeground="#ffffff", cursor="hand2").place(x=300, y=350, width=130, height=30)
		btn_send2 = Button(self.root2, command=self.email_login, text="SAVE", font=("Century Gothic",18,"bold"), bg="#32abe6", fg="#ffffff", activebackground="#32abe6", activeforeground="#ffffff", cursor="hand2").place(x=440, y=350, width=130, height=30)

		# Inserting already saved Email and password to the field
		self.txt_email.insert(0,self.email_)
		self.txt_password.insert(0,self.password_)
	
	# Clear Function
	def clear2(self):
		self.txt_email.delete(0, END)
		self.txt_password.delete(0, END)


	def checking_if_login_file_exist(self):
		if os.path.exists("login_credencials.txt")==False:
			f = open('login_credencials.txt', 'w')
			f.write(",")
			f.close()
		f2=open("login_credencials.txt","r")
		self.credencials=[]
		for i in f2:
			self.credencials.append( [i.split(",")[0],i.split(",")[1]] )
		self.email_=self.credencials[0][0]
		self.password_=self.credencials[0][1]


	def email_login(self):
		if self.txt_email.get()=="" or self.txt_password.get()=="":
			messagebox.showerror("Error", "All field are required", parent=self.root2)
		else:
			f = open('login_credencials.txt', 'w')
			f.write(self.txt_email.get() + "," + self.txt_password.get())
			f.close()
			messagebox.showinfo("Success", "Email save Successfully", parent=self.root)
			self.checking_if_login_file_exist()


root = Tk()
obj = Bulk_Email(root)
root.mainloop()