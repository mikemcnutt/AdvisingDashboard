import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import webbrowser
import uuid
import datetime
import html
import win32com.client

APP_TITLE = "Advising Dashboard"
HEADER_TEXT = "One Dashboard To Rule Them All"

BLUE1 = "#1e3a8a"
BLUE2 = "#3b82f6"
CARD = "#e0e7ff"

def load_json(path):
    with open(path,"r",encoding="utf-8") as f:
        return json.load(f)

def find_plan(obj,season,year):
    plans=obj.get("data",{}).get("semesterPlans",[])
    for p in plans:
        if p.get("season")==season and str(p.get("year"))==str(year):
            return p
    return None

def term_state(obj,season,year):
    plan=find_plan(obj,season,year)
    if not plan:
        return "unadvised"
    courses=plan.get("courses",[])
    if not courses:
        return "unadvised"
    if plan.get("notComplete"):
        return "partial"
    return "done"

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1300x850")

        self.folder=tk.StringVar()
        self.year=tk.StringVar(value="2026")

        self.spring=tk.BooleanVar()
        self.summer=tk.BooleanVar()
        self.fall=tk.BooleanVar(value=True)

        self.search=tk.StringVar()

        self.subject=tk.StringVar(value="Advising Appointment Needed")
        self.link=tk.StringVar()

        self.students=[]
        self.needs=[]
        self.partial=[]
        self.done=[]

        self.build_ui()

    def build_ui(self):

        header=tk.Label(self,text=HEADER_TEXT,font=("Segoe UI",22,"bold"),fg="white",bg=BLUE1)
        header.pack(fill="x")

        top=tk.Frame(self)
        top.pack(fill="x",padx=10,pady=8)

        ttk.Label(top,text="Year").pack(side="left")
        ttk.Combobox(top,textvariable=self.year,values=[str(x) for x in range(2026,2035)],width=6).pack(side="left",padx=5)

        ttk.Checkbutton(top,text="Spring",variable=self.spring).pack(side="left")
        ttk.Checkbutton(top,text="Summer",variable=self.summer).pack(side="left")
        ttk.Checkbutton(top,text="Fall",variable=self.fall).pack(side="left")

        ttk.Entry(top,textvariable=self.search,width=20).pack(side="left",padx=10)

        ttk.Entry(top,textvariable=self.folder,width=40).pack(side="left",padx=10)
        ttk.Button(top,text="Browse",command=self.browse).pack(side="left")

        ttk.Button(top,text="Scan",command=self.scan).pack(side="left",padx=10)

        email=tk.LabelFrame(self,text="Email Settings")
        email.pack(fill="x",padx=10,pady=6)

        ttk.Label(email,text="Subject").pack(anchor="w")
        ttk.Entry(email,textvariable=self.subject).pack(fill="x")

        ttk.Label(email,text="Scheduling link").pack(anchor="w")
        ttk.Entry(email,textvariable=self.link).pack(fill="x")

        ttk.Label(email,text="Message").pack(anchor="w")

        self.body=tk.Text(email,height=4)
        self.body.pack(fill="x")
        self.body.insert("1.0","Please reply to schedule an advising appointment.")

        main=tk.Frame(self)
        main.pack(fill="both",expand=True,padx=10,pady=10)

        self.col1=self.make_column(main,"Needs Advised")
        self.col2=self.make_column(main,"Advised Not Complete")
        self.col3=self.make_column(main,"Advised")

        self.col1.pack(side="left",fill="both",expand=True,padx=5)
        self.col2.pack(side="left",fill="both",expand=True,padx=5)
        self.col3.pack(side="left",fill="both",expand=True,padx=5)

    def make_column(self,parent,title):
        frame=tk.LabelFrame(parent,text=title)
        canvas=tk.Canvas(frame)
        scroll=ttk.Scrollbar(frame,orient="vertical",command=canvas.yview)
        inner=tk.Frame(canvas)

        inner.bind("<Configure>",lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0,0),window=inner,anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)

        canvas.pack(side="left",fill="both",expand=True)
        scroll.pack(side="right",fill="y")

        frame.inner=inner
        return frame

    def browse(self):
        d=filedialog.askdirectory()
        if d:
            self.folder.set(d)

    def selected_terms(self):
        terms=[]
        y=self.year.get()

        if self.spring.get():
            terms.append(("Spring",y))
        if self.summer.get():
            terms.append(("Summer",y))
        if self.fall.get():
            terms.append(("Fall",y))

        return terms

    def scan(self):

        folder=Path(self.folder.get())
        if not folder.exists():
            messagebox.showerror("Error","Folder not found")
            return

        terms=self.selected_terms()
        if not terms:
            messagebox.showerror("Error","Select a semester")
            return

        self.needs=[]
        self.partial=[]
        self.done=[]

        for f in folder.rglob("*.json"):
            try:
                obj=load_json(f)
                name=obj["student"]["firstName"]+" "+obj["student"]["lastName"]
                sid=obj["student"]["studentId"]

                bucket="done"

                for season,year in terms:
                    state=term_state(obj,season,year)

                    if state=="unadvised":
                        bucket="needs"
                        break
                    elif state=="partial":
                        bucket="partial"

                data=(name,sid,str(f),obj)

                if bucket=="needs":
                    self.needs.append(data)
                elif bucket=="partial":
                    self.partial.append(data)
                else:
                    self.done.append(data)

            except:
                pass

        self.render()

    def render(self):

        for col in [self.col1,self.col2,self.col3]:
            for w in col.inner.winfo_children():
                w.destroy()

        for name,sid,path,obj in self.needs:
            self.render_student(self.col1.inner,name,sid,path,obj,needs=True)

        for name,sid,path,obj in self.partial:
            self.render_student(self.col2.inner,name,sid,path,obj,email_button=True)

        for name,sid,path,obj in self.done:
            self.render_student(self.col3.inner,name,sid,path,obj)

    def render_student(self,parent,name,sid,path,obj,needs=False,email_button=False):

        row=tk.Frame(parent,bd=1,relief="solid",bg=CARD)
        row.pack(fill="x",pady=3)

        link=tk.Label(row,text=name,font=("Segoe UI",10,"bold"),fg="blue",cursor="hand2",bg=CARD)
        link.pack(anchor="w",padx=5)

        link.bind("<Button-1>",lambda e:self.open_editor(path))

        tk.Label(row,text=sid,bg=CARD).pack(anchor="w",padx=5)

        if email_button:
            b=tk.Button(row,text="Email",bg=BLUE2,fg="white",command=lambda:self.email_student(obj))
            b.pack(anchor="e",padx=5,pady=3)

    def open_editor(self,json_path):

        token=str(uuid.uuid4())
        url=f"http://127.0.0.1:8123/Advising10.html?token={token}&json={json_path}"
        webbrowser.open(url)

    def email_student(self,obj):

        first=obj["student"].get("firstName","")
        k=obj["student"].get("kctcsEmail","")
        p=obj["student"].get("personalEmail","")

        body=self.body.get("1.0","end").strip()
        subject=self.subject.get()

        html_body=f"""
        <html>
        <body>
        <p>Hello {html.escape(first)},</p>
        <p>{html.escape(body)}</p>
        <p><a href="{html.escape(self.link.get())}">Schedule Appointment</a></p>
        </body>
        </html>
        """

        outlook=win32com.client.Dispatch("Outlook.Application")
        mail=outlook.CreateItem(0)

        mail.To=";".join([k,p])
        mail.Subject=subject
        mail.HTMLBody=html_body
        mail.Display()

if __name__=="__main__":
    app=App()
    app.mainloop()
