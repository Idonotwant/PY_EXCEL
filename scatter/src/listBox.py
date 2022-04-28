'''
Author: Wei-Ju-Chen(Idonotwant)
ProjectName: listBox
Description: listBox
Brief_Log:
2022/04/26 Start
Note: 
'''
from cProfile import label
from tkinter import *
from tkinterdnd2 import DND_FILES,TkinterDnD
alignmode = 'nswe'




class LListbox(TkinterDnD.Tk):
    def __init__(self, *args,**aragvs):
        super().__init__()
        #basic setting 
        self.wm_title('List box practice')
        self.configure(background='white')
        self.wm_minsize(800,600)
        self.wm_maxsize(self.winfo_screenwidth(),self.winfo_screenheight())

        self.filepaths = []
        #top level grid
        self.div1 = Frame(self,bg='green')
        self.div1.grid(row=0,column=0,sticky=alignmode)
        #inputbox
        self.iB1 = Entry(self.div1)
        self.iB1.bind('<Return>',self.ib1Enter)
        self.iB1.grid(row=0,column=0,columnspan=2,sticky=alignmode)
        
        #lisgbox,with common selecting mode
        self.lB1 = Listbox(self.div1,selectmode=EXTENDED,background='#ffe0d6')
        self.lB1.bind('<Delete>',self.lb1Del)
        self.lB1.drop_target_register(DND_FILES)
        self.lB1.dnd_bind("<<Drop>>",self.dropIn)
        self.lB1.bind("<Return>",self.leave)
        self.lB1.grid(column=0,row=1,sticky=alignmode)

        #div2,input control box
        self.div2 = Frame(self.div1,bg='gray')
        self.div2.grid(column=1,row=1,sticky=alignmode)

        #control button in div2
        self.bClean = Button(self.div2,background="white",text='Clean',command=self.cleanlB1)
        self.bClean.grid(column=0,row=0,sticky=alignmode)
        self.bEnter = Button(self.div2,background='white',text='Enter',command=self.leave)
        self.bEnter.grid(column=0,row=1,sticky=alignmode)
        self.labExplain = Label(self.div2,text='select and click enter')
        self.labExplain.grid(column=0,row=2)
        #canvas
        #self.canPic = Canvas(self.div2,background='black')
        
        #align element
        self.div2.columnconfigure(0,weight=1)
        self.div2.rowconfigure(0,weight=1)
        self.div2.rowconfigure(1,weight=1)
        self.div2.rowconfigure(2,weight=1)
        self.div1.columnconfigure(0,weight=5)
        self.div1.columnconfigure(1,weight=1)
        self.div1.rowconfigure(0,weight=1)
        self.div1.rowconfigure(1,weight=5)
        
        self.rowconfigure(0,weight=1)
        self.columnconfigure(0,weight=1)

    def run(self,*args):# run the kernal
        self.mainloop()
    def leave(self,*args):# enter
        data = self.lB1.curselection()
        if not len(data):#empty
            self.filepaths = (self.lB1.get(0,self.lB1.size()-1))
        else:
            self.filepaths = (self.lB1.get(data[0],data[-1]))
        self.destroy()
    def ib1Enter(self,*args):# enter the value
        self.lB1.insert(END,self.iB1.get())
        self.iB1.delete("0",END)
    def lb1Del(self,*args):# delete the value in box
        data = self.lB1.curselection()
        if not len(data):#data empty,default to delete top one
            self.lB1.delete(0)
        else:
            self.lB1.delete(data[0],data[-1])
    def cleanlB1(self,*args):# clean box
        self.lB1.delete(0,self.lB1.size()-1)

    def dropIn(self,event):
        self.lB1.insert("end",event.data)
if __name__ == '__main__':
    a = LListbox()
    a.run()
    for i in range(len(a.filepaths)):
        print(a.filepaths[i])