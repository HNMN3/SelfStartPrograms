import os, winshell
from win32com.client import Dispatch
import os
import shutil
import subprocess
import Tkinter
from _winreg  import *
from Tkinter import tkinter
import pickle

class MainWindow:
    def __init__(self,parent):
        self.username = os.path.expandvars("%userprofile%")
        self.temp_add,self.temp_delete = {},{}
        self.username=self.username[self.username.rindex("\\")+1:]
        self.start_path = r"C:\Documents and Settings\%s\Start Menu\Programs\Startup" %self.username
        self.parent = parent
        if not os.path.exists("./Data"):
            os.mkdir("./Data",000)
        self.filename = "Data/data.aspf"
        self.allPrograms = "Data/programs.aspf"
        self.dirty = False
        self.data_all,self.data_to_open = {},{}
        #creating menubar
        menubar = tkinter.Menu(self.parent)
        self.parent["menu"] = menubar

        #File option in menubar
        fileMenu = tkinter.Menu(menubar,tearoff=0)

        fileMenu.add_command(label = "Save..",command = self.saveIt,accelerator = "Ctrl+S",underline=0)
        self.parent.bind("<Control-s>",self.saveIt)
        fileMenu.add_separator()
        fileMenu.add_command(label = "Exit",command = self.quitConfirm,accelerator = "Ctrl+Q",underline=0)
        self.parent.bind("<Control-q>",self.quitConfirm)

        menubar.add_cascade(label = "File",menu = fileMenu,underline=0)
        #Edit Options in menubar
        editMenu = tkinter.Menu(menubar,tearoff=0)
        editMenu.add_command(label="Find another software..",command = self.findSoftware,accelerator="Ctrl+F",underline=0)
        self.parent.bind("<Control-f>",self.findSoftware)
        editMenu.add_separator()
        editMenu.add_command(label = "Reset",command = self.reset,accelerator="Ctrl+R",underline=0)
        self.parent.bind("<Control-r>",self.reset)

        menubar.add_cascade(label="Edit",menu=editMenu,underline=0)

        #Main Frame below menubar
        frame = tkinter.Frame(self.parent)

        #Headings of the lists
        labelAllProgram = tkinter.Label(frame,text="All Programs",anchor=tkinter.W,height=2)
        labelBootProgram = tkinter.Label(frame,text="Programs to be start on Boot up",anchor=tkinter.E)
        labelAllProgram.grid(row=0,column=0,sticky=tkinter.W)
        labelBootProgram.grid(row=0,column=3,sticky=tkinter.W)

        #list1 with scrollbar
        scrollbar1 = tkinter.Scrollbar(frame,orient= tkinter.VERTICAL)
        self.list1 = tkinter.Listbox(frame,yscrollcommand = scrollbar1.set,selectmode='multiple',height=35,width=90)
        self.list1.grid(row=1,column=0,sticky="ns")

        scrollbar1["command"] = self.list1.yview
        scrollbar1.grid(row=1,column=1,sticky = tkinter.NS)

        #Buttons For Shifting Programms between lists
        self.lrTools = tkinter.Frame(frame)
        try:
            self.imgRight = tkinter.PhotoImage(file="Image/right.GIF")
            self.imgLeft = tkinter.PhotoImage(file="Image/left.GIF")
            self.imgRefresh = tkinter.PhotoImage(file="Image/refresh.gif")
            buttonRight = tkinter.Button(self.lrTools,image=self.imgRight,command=self.moveToRight,width=55,height=55)
            buttonRight.grid(row=0,column=0,pady=60)
            buttonLeft = tkinter.Button(self.lrTools,image=self.imgLeft,command=self.moveToLeft,width=55,height=55)
            buttonLeft.grid(row=3,column=0,pady=60)
            buttonRefresh = tkinter.Button(self.lrTools,image=self.imgRefresh,command=self.Refresh,width=55,height=55)
            buttonRefresh.grid(row=6,column=0,pady=60)
        except Exception as e:
            messagebox.showwarning("Error","Error in loading file..  " + str(e))
            self.statusBar.after(3000,self.clearStatusBar())
        self.lrTools.grid(row=1,column=2,rowspan=2,sticky=tkinter.NS)

        #list2 with scrollbar
        scrollbar2 = tkinter.Scrollbar(frame,orient= tkinter.VERTICAL)
        self.list2 = tkinter.Listbox(frame,yscrollcommand = scrollbar2.set,selectmode='multiple',height=35,width=90)
        self.list2.grid(row=1,column=3,sticky=tkinter.NSEW)
        scrollbar2["command"] = self.list2.yview
        scrollbar2.grid(row=1,column=4,sticky = tkinter.NS)

        #status Bar of Application
        self.statusBar = tkinter.Label(frame,text="Ready...",anchor=tkinter.W,height=2)
        self.statusBar.after(3000,self.clearStatusBar)
        self.statusBar.grid(row=2,column=0,sticky=tkinter.EW,columnspan=3)
        #save Button
        self.saveButton = tkinter.Button(frame,text="Save",command = self.saveIt)
        self.saveButton.grid(row=2,column=3,sticky=tkinter.E)

        #run Software Button
        self.runButton = tkinter.Button(frame,text="Run",command = self.runProgram)
        self.runButton.grid(row=2,column=3,sticky=tkinter.W)

        #loading the data
        self.loadDataList2()
        self.loadDataList1()

        #Showing frame to window and configuring weights
        frame.grid(sticky=tkinter.NSEW)

    def makeSoftwareStart(self,*ignore):
        ##print("Making")
        for item in self.temp_delete.keys():
            str = self.start_path+"\\"+item +".lnk"
            #print("deleting " + str)
            os.unlink(str)
            if os.path.exists(str):
                print("unable to remove " +str)
        try:
            shell = Dispatch('WScript.Shell')
            for item in self.temp_add.keys():
                path  = self.start_path +"\\" + item + ".lnk"
                target = self.temp_add[item]
                shortcut = shell.CreateShortCut(path)
                shortcut.Targetpath = target
                shortcut.save()
        except Exception as e:
            print("Error" + str(e))
        self.temp_delete.clear()
        self.temp_add.clear()
        return

    def Refresh(self,*ignore):
        self.updateList1()
        self.updateList2()

    def runProgram(self,*ignore):
        flag = (True if len(self.list1.curselection())>0 else False)
        if not flag and self.list2.curselection() is None:
            ##print("No Data")
            return
        ##print(flag)
        index = (self.list1.curselection() if flag else self.list2.curselection())
        ##print(index)
        for i in index:
            item = (self.list1.get(i) if flag else self.list2.get(i))
            ##print(item)
            s = (self.data_all[item] if flag else self.data_to_open[item])
            subprocess.call([s])

    def loadDataList2(self,*ignore):
        try:
            if not os.path.exists(os.curdir + "/" + self.filename):
                ##print("file not exist")
                [os.unlink(self.start_path + "\\" + f) for f in os.listdir(self.start_path)]
                return

            with open(self.filename,"rb") as fh:
                self.data_to_open = pickle.load(fh)
            fh.close()
            self.updateList2()
            self.statusBar["text"] = "{0} loaded".format(len(self.data_to_open))
            self.statusBar.after(3000,self.clearStatusBar)
        except Exception:
            print("Error occured " + str(Exception))
    def loadDataList1(self,*ignore):
        if(not os.path.exists(self.allPrograms)):
            [os.unlink(self.start_path + "\\" + f) for f in os.listdir(self.start_path)]
            aReg = ConnectRegistry(None,HKEY_LOCAL_MACHINE)
            aKey = OpenKey(aReg, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
            for i in range(1024):
                try:
                    asubkey_name=EnumKey(aKey,i)
                    asubkey=OpenKey(aKey,asubkey_name)
                    name=QueryValueEx(asubkey, "DisplayName")[0]
                    path = QueryValueEx(asubkey,"DisplayIcon")[0]
                    if(path.endswith("exe") and not name in self.data_to_open.keys()):
                        self.data_all[name]=path
                except EnvironmentError as e:
                    continue
            with open(self.allPrograms,"wb") as fh:
                pickle.dump(self.data_all,fh,pickle.HIGHEST_PROTOCOL)
            fh.close()
        else:
            ##print("Already saved")
            with open(self.allPrograms,"rb") as fh:
                self.data_all = pickle.load(fh)
            fh.close()
        self.updateList1()

    def findSoftware(self,*ignore):
        ##print("Finding")
        file = filedialog.askopenfilename(title="Auto Start Programs",initialdir=".",filetypes=[("Executable Files","*.exe")])
        if not file is None:
            sep = file.split('/')
            key = sep[len(sep)-1]
            key = str(key[:len(key)-4])
            key = key.capitalize()
            if (key not in self.data_all.keys()) and (key not in self.data_to_open.keys()):
                self.statusBar["text"] = "Application added to Programs list"
                self.statusBar.after(5000,self.clearStatusBar())
                self.data_all[key]=file
                self.updateList1()
                self.dirty = True
            else:
                self.statusBar["text"] = "File is Already in list"
                self.statusBar.after(3000,self.clearStatusBar())
        else:
            self.statusBar["text"] = "File is corrupted"
            self.statusBar.after(3000,self.clearStatusBar())
    def updateList1(self,*ignore):
        self.list1.delete(0,tkinter.END)
        for item in sorted(self.data_all.keys()):
            self.list1.insert(tkinter.END,item)
    def updateList2(self,*ignore):
        self.list2.delete(0,tkinter.END)
        for item in sorted(self.data_to_open.keys()):
            self.list2.insert(tkinter.END,item)
    def moveToRight(self,*ignore):
        ##print(self.parent.winfo_height())
        index = self.list1.curselection()
        if not index:
            return
        self.dirty = True
        #print(self.data_all)
        for item in sorted(index,reverse=True):
            data = self.list1.get(item)
            self.data_to_open[data]=str(self.data_all[data])
            self.temp_add[data] = self.data_to_open[data]
            if data in self.temp_delete.keys():
                self.temp_delete.pop(data)
                #print(self.temp_delete)
            self.data_all.pop(data)
            self.list1.delete(item)
        self.updateList2()
        self.statusBar["text"] = "Softwares moved Succesfully"
        self.statusBar.after(2000,self.clearStatusBar)
    def moveToLeft(self,*ignore):
        index = self.list2.curselection()
        if not index:
            return

        self.dirty=True
        for item in sorted(index,reverse=True):
            data = str(self.list2.get(item))
            self.data_all[data] = str(self.data_to_open[data])
            self.temp_delete[data] = self.data_all[data]
            #print(self.temp_delete)
            if data in self.temp_add.keys():
                self.temp_add.pop(data)
            self.data_to_open.pop(data)
            self.list2.delete(item)
        self.updateList1()
        self.statusBar["text"] = "Softwares moved Succesfully"
        self.statusBar.after(2000,self.clearStatusBar)
    def clearStatusBar(self,*ignore):
        self.statusBar["text"] = "Waiting for your next event.."
    def saveIt(self,*ignore):
        try:
            self.makeSoftwareStart()
            if os.path.exists(self.filename):
                os.chmod(self.filename,0o777)
                os.unlink(self.filename)
            #print(self.data_to_open)
            with open(self.filename,"wb") as fh:
                pickle.dump(self.data_to_open,fh,pickle.HIGHEST_PROTOCOL)
            fh.close()
            if os.path.exists(self.allPrograms):
                os.chmod(self.allPrograms,0o777)
                os.unlink(self.allPrograms)
            with open(self.allPrograms,"wb") as fi:
                pickle.dump(self.data_all,fi,pickle.HIGHEST_PROTOCOL)
            fi.close()
        except Exception as e:
            #print(e)
            self.statusBar["text"] = "Error in saving file"
            self.statusBar.after(3000,self.clearStatusBar)
        self.statusBar["text"] = "All Saved Succesfully"
        self.statusBar.after(3000,self.clearStatusBar)
        self.dirty=False

    def quit(self,*ignore):
        self.parent.destroy()

    def reset(self,*ignore):
        self.list2.delete(0,tkinter.END)
        ##print(len(self.data_all))
        for key in self.data_to_open.keys():
            self.data_all[key] = self.data_to_open[key]
            self.temp_delete[key] = self.data_to_open[key]
            if key in self.temp_add.keys():
                self.temp_add.pop(key)
        self.data_to_open.clear()
        #print(self.data_to_open)
        self.updateList1()
        self.updateList2()
        self.dirty=True
    def quitConfirm(self,*ignore):
        if not self.dirty:
            ##print("Closed")
            self.quit()
        else:

            saveornot = messagebox.askyesno("AutoStartSoftware","Changes are not saved..do you want to quit?")
            if saveornot==False:
                return
            self.quit()

if __name__=="__main__":
    window = Tkinter.Tk()
    window.state('normal')
    m = Tkinter.MainWindow(window)
    window.bind("<Escape>", lambda *ignore: m.quitConfirm())
    window.protocol("WM_DELETE_WINDOW",m.quitConfirm)
    window.resizable(False,False)
    window.title("Auto Start Programs")
    header = tkinter.PhotoImage(file="Image/download.ico")
    window.tk.call('wm', 'iconphoto', window._w, header)
    window.mainloop()
