import psutil
import time
from datetime import datetime
from tkinter import messagebox
import pymsteams
import os
from configparser import ConfigParser
import configparser
from tkinter import ttk
from tkinter.ttk import *
from tkinter import *
import sys
import keyboard
import threading

global stringF, radio, ii, durration, user

stringF = None
radio = 1
stringF= None
ii = 0
durration = 300
web_hook = 'WEB_HOOK_URL'
gate_1 = False

try:
    user = os.path.expanduser('~')
    cpu = os.environ['COMPUTERNAME']
except:
    user = "UNKNOWN"
    cpu = "UNKNOWN"

def start():

    config() 
    
    if stringF == None:
        listexe()
    else:

        start_thread_timer()
        timeout = 8 
        timeout_start = time.time()
    
        while time.time() < timeout_start + timeout:
            if keyboard.is_pressed('F2'):
                try:
                    ws.destroy()
                except:
                    pass
                listexe()
    
    chk()

def timer_window():
    global stringF, ws
    ws = Tk()
    ws.eval('tk::PlaceWindow . center')
    ws.title('Status Monitor')
    ws.geometry('500x230')
    string_label = ttk.Label(ws, text='The following processe(s) will be monitored:', font=('Arial',15), justify=CENTER)
    string_label.pack(pady=(25,0))
    string_labelprc = ttk.Label(ws, text='{0}\n'.format("\n\u2022 ".join(['', *stringF])), font=('Arial',11))
    string_labelprc.pack()
    string_label2 = ttk.Label(ws, text='Press the "F2" key to reconfigure.', font=('Arial',15), justify=CENTER)
    string_label2.pack()

    def timer_func():
        for i in range(0,9):
            
            if i == 8:
                
                string_label2.destroy()
                timer_label1 = ttk.Label(ws, text='Status Monitor Running...',font=('Arial',15), justify=CENTER)
                timer_label1.pack(padx=2,pady=2)
                    
            if i < 4:
                timer_label = Label(ws, text=f'{(8-i)} second(s) left', font=('Arial',25), justify=CENTER)
                timer_label.pack(padx=2,pady=2)

            if i > 3 and i <8:
                timer_label = Label(ws, text=f'{(8-i)} second(s) left', fg='red', font=('Arial',25), justify=CENTER)
                timer_label.pack(padx=2,pady=2)
                # timer_label.config(fg='red')
                
            ws.update()
            timer_label.destroy()
            time.sleep(1)
            
            

    ws.attributes('-topmost', True)
    ws.update()
    ws.attributes('-topmost', False)
    ws.after(1, timer_func)
    ws.mainloop()
    

def start_thread_timer():
    start_thread=threading.Thread(target=timer_window)
    start_thread.start() 

def start_thread_start():
    start_thread=threading.Thread(target=start)
    start_thread.start() 

def listexe():
    global user,cpu
    def Scankey(event):
        global list1, ii
        val = event.widget.get()

        val_len = len(val)

        if val == '':
            data = list1
            ii = 0

        else:
            data = []
            value = selected.get()
            check = var1.get()
            if value == 1:
                if check == 0:
                    for item in list1:
                        if val.lower() in item[0].lower()[0:val_len]:
                            data.append(item)
                if check == 1:
                    for item in list1:
                        if val.lower() in item[0].lower():
                            data.append(item)		
            if value == 2:	
                for item in list1:
                    if val in str(item[1])[0:val_len]:
                        data.append(item)			

        Update(data)
        
        if val.isnumeric() == True and selected.get() != 2 and val != '' and ii != 1:
            ch = messagebox.askyesno('Input','Numeric input detected, switch to PID search option?')
            if ch == True:
                selected.set(2)
            if ch == False:
                ii = 1
                
            
        elif val.isnumeric() == False and selected.get() != 1 and val != '' and ii != 1:
            ch = messagebox.askyesno('Input','Non-numeric input detected, switch to search by name option?')
            if ch == True:
                selected.set(1)
            if ch == False:
                ii = 1
                
    def Update(data):
        
        for item in listbox.get_children():
            listbox.delete(item)
       
        # put new data
        for item in data:
            listbox.insert('','end', value=item)
            
    
    def Update_lb2(data):
        
        listbox2.delete(0, 'end')
        # put new data
        for item in data:
            listbox2.insert('end', item)

    def genList():

        global list1
        # res = []
        list1 = []
        for process in psutil.process_iter(attrs=['name','pid','status']):
            list1.append((process.name(), process.pid, process.status()))
        # list1 = []
        # [list1.append(x) for x in res if x not in list1]
        return sorted(list1)

    def add():

        selectedval = selected.get()
        if selectedval == 2:
            for selected_item in listbox.selection():
                item = listbox.item(selected_item)
                record = item['values']
                if str(record[1]) in listbox2.get(0,END):
                    messagebox.showerror('Duplicate value','Cannot add duplicate value...')
                    return
                listbox2.insert(END,str(record[1]))

        if selectedval == 1:
            for selected_item in listbox.selection():
                item = listbox.item(selected_item)
                record = item['values']
                if record[0] in listbox2.get(0,END):
                    messagebox.showerror('Duplicate value','Cannot add duplicate value...')
                    return
                listbox2.insert(END,record[0])

    def add_manual(event):

        selectedval = selected.get()
        if selectedval == 2:
            if event.widget.get() == '':
                return
            else:
                if event.widget.get().isnumeric() == True:
                    listbox2.insert(END,event.widget.get())
                    manual_entry.delete(0, 'end')
                else:
                    messagebox.showerror('PID value error','PID must be an integer value.')
        if selectedval == 1:
            if event.widget.get() == '':
                return
            else:
                listbox2.insert(END,event.widget.get())
                manual_entry.delete(0, 'end')


    def remove():
        listbox2.delete(ANCHOR)
    
    def sendout():
        global stringF, radio, durration, web_hook
        stringF = listbox2.get(0, END)
        while '' in stringF:
            stringF = None
        
        radio = selected.get()
        durration = durration_VAR.get()
        web_hook = web_hook_url.get()
        ws.destroy()
        return stringF, radio, durration, web_hook

    def click(*args):
        if search_entry.get() == 'Search...':
            search_entry.delete(0, 'end')
            ws.focus()
        

    def click_manual(*args):
        if manual_entry.get() == 'Manual entry...':
            manual_entry.delete(0, 'end')
            ws.focus()
  
    def leave(*args):
        if len(search_entry.get()) == 0:
            search_entry.delete(0, 'end')
            search_entry.insert(0, 'Search...')
            ws.focus()

    def leave_manual(*args):
        if len(manual_entry.get()) == 0:
            manual_entry.delete(0, 'end')
            manual_entry.insert(0, 'Manual entry...')
            ws.focus()

    def exit():
        sys.exit()

    def relist(): 
        global list1
        list1 = genList()
        Update(list1)
        return list1
    
    def enchkb(event):
        checkbox.config(state='enabled')
        lb2_list = listbox2.get(0, END)
        
        res=[]
       
        for item in lb2_list:
            for value in list1:
                if item == str(value[1]):
                    res.append(value[0])
                
        resfinal = []
        [resfinal.append(x) for x in res if x not in resfinal]
     
        Update_lb2(resfinal)

    def dischkb(event):
        checkbox.config(state='disabled')
        lb2_list = listbox2.get(0, END)  

        res=[]
   
        for item in lb2_list:
            for value in list1:
                if item == value[0]:
                    res.append(str(value[1]))
          
        resfinal = []
        [resfinal.append(x) for x in res if x not in resfinal]
   
        Update_lb2(resfinal)
      
    def clear_all():
        listbox2.delete(0,END)

    def option_menu():
        newwindow3 = Toplevel()
        newwindow3.attributes('-topmost', True)
        newwindow3.title('Options')
        newwindow3.geometry()
        mainsource = LabelFrame(newwindow3, text = 'Settings')
        mainsource.pack(side=LEFT, pady=1, padx=4, fill=X, expand=True)

        durration_label = ttk.Label(mainsource, text="Scan interval (second's):", )
        durration_label.pack(side=TOP,pady=3,anchor=W)
        durration_entry = ttk.Entry(mainsource, textvariable=durration_VAR, justify='center')
        durration_entry.pack(side=TOP, expand=False, anchor=W, padx=2, fill=X,)

        teams_hook = ttk.Label(mainsource, text='Microsoft Webhook URL:', )
        teams_hook.pack(side=TOP,pady=2, anchor=W)
        teams_hook_entry = ttk.Entry(mainsource, textvariable=web_hook_url)
        teams_hook_entry.pack(side=TOP, expand=True,fill=X,anchor=W, padx=(0,2), pady=2)

        test_btn = ttk.Button(mainsource, text='TEST', command=notify)
        test_btn.pack(fill=X, expand=True,anchor=W, padx=2, pady=2)

    def notify():
    
        now = datetime.now()  
        eventTime = now.strftime("%H:%M:%S, %m/%d/%Y")
        myTeamsMessage = pymsteams.connectorcard(web_hook_url.get())
        myTeamsMessage.title("Service Monitor")
        myTeamsMessage.text('The following service(s) are being monitored as of {0}:   \n___\n{3}   \n___\nLocation-   \nUsername: {1}   \nPC Name: {2}'.format(eventTime, user, cpu,"   \n\u2022 ".join(['', *listbox2.get(0, END)])))
        myTeamsMessage.color("#35db40")
        myTeamsMessage.send()
        messagebox.showinfo("Service Monitor", 'The following service(s) are being monitored as of {0}:\n{1}'.format(eventTime,"\n\u2022 ".join(['', *listbox2.get(0, END)])))

# --------------------------------------------------------------------------------------
    ws = Tk()
    ws.title('Status Monitor')
    
    # ws.geometry('271x183+1630+800')
    style = ttk.Style()
    style.theme_use('xpnative')
    style.configure('TButton', justify='center')
    
    menubar = Menu(ws)
    filemenu = Menu(menubar, tearoff=0)
    filemenu.add_command(label="Options", command=option_menu)
    filemenu.add_command(label="Exit", command=ws.quit)
    menubar.add_cascade(label="File", menu=filemenu)
    

# --------------------------------------------------------------------------------------
    sub_master_frame = Frame(ws)
    sub_master_frame.pack(anchor=W)

    box_frame = LabelFrame(sub_master_frame, text='Running Processes', )
    box_frame.pack(side=LEFT, padx=(5,5), pady=(0,5))

    RadioFrame = LabelFrame(sub_master_frame, text = 'Add/Remove', )
    RadioFrame.pack(side=TOP, padx=(0,5), pady=(0,0), anchor=N, fill=Y, expand=True)   

    buttons = Frame(sub_master_frame,)
    buttons.pack(side=BOTTOM, padx=(0,6), pady=(1,5))

# --------------------------------------------------------------------------------------
    main_frame = Frame(box_frame)
    main_frame.pack(anchor=W,  padx=5)

    main_frame_options = Frame(main_frame,)
    main_frame_options.pack(side=RIGHT,anchor=E, padx=(4,0), pady=(0,1))
   
# --------------------------------------------------------------------------------------
    search_entry = ttk.Entry(main_frame, width=30)
    search_entry.pack(side=LEFT,expand=TRUE, pady=(0,1), anchor=W)
    search_entry.bind('<KeyRelease>', Scankey)
    search_entry.bind("<Button-1>", click)
    search_entry.insert(0, 'Search...')     
    search_entry.bind("<Leave>", leave) 
# --------------------------------------------------------------------------------------

    selected = IntVar()
    var1 = IntVar()

    durration_VAR = IntVar()
    durration_VAR.set(durration)

    web_hook_url = StringVar()
    web_hook_url.set(web_hook)

    checkbox = ttk.Checkbutton(main_frame_options, text="Contains", variable=var1, onvalue=1, offvalue=0)
    checkbox.pack(side=RIGHT,anchor=E)

    r2 = ttk.Radiobutton(main_frame_options, text='PID',value=2, variable=selected)
    r2.pack(padx=4, side=RIGHT,anchor=E)
    r1 = ttk.Radiobutton(main_frame_options, text='Name',value=1,variable=selected)
    r1.pack(side=RIGHT,anchor=E)
    r1.bind("<Button-1>", enchkb)
    r2.bind("<Button-1>", dischkb)
    selected.set(radio)

# --------------------------------------------------------------------------------------
    addremove = Frame(RadioFrame)
    addremove.pack(side=LEFT, padx=5, pady=5)

    Display = ttk.Button(addremove,text ="OK",command = sendout)
    Display.pack(side=BOTTOM)
    
    add1 = ttk.Button(addremove,text ="Add",command = add)
    add1.pack(side=TOP)

    remove1 = ttk.Button(addremove,text ="Remove",command = remove)
    remove1.pack()
# --------------------------------------------------------------------------------------
    updatelist = ttk.Button(buttons,text ="Reload",command = relist)
    updatelist.pack(side=LEFT)

    quit = ttk.Button(buttons,text ="Exit",command = exit)
    quit.pack(side=RIGHT)

    clearall = ttk.Button(buttons,text ="Clear All",command = clear_all)
    clearall.pack(side=LEFT)
# --------------------------------------------------------------------------------------
    columns = ('Process_name', 'PID', 'Status')

    listbox = ttk.Treeview(box_frame, columns=columns, show='headings')
    listbox.heading('Process_name', text='Process Name')
    listbox.heading('PID', text='PID')
    listbox.heading('Status', text='Status')
    listbox.column('Status', anchor=CENTER, width=100)
    listbox.column('PID', anchor=CENTER, width=100)
    listbox.pack(side=LEFT, pady=(0,5) , padx=(5,0))

    scrollbar = ttk.Scrollbar(box_frame, orient='vertical', command=listbox.yview)
    scrollbar.pack(side=RIGHT, pady=(0,5),padx=(0,5), fill=Y)
    listbox.configure(yscroll=scrollbar.set)
# --------------------------------------------------------------------------------------
    man_lb = Frame(RadioFrame)
    man_lb.pack(side=RIGHT,padx=(0,5))

    listbox2 = Listbox(man_lb, relief=FLAT, border=2)
    listbox2.pack(expand=True,side=BOTTOM, fill=X,)

    manual_entry = ttk.Entry(man_lb)
    manual_entry.bind("<Button-1>", click_manual)
    manual_entry.insert(0, 'Manual entry...')     
    manual_entry.bind("<Leave>", leave_manual) 
    manual_entry.bind("<Return>", add_manual) 
    manual_entry.pack(expand=True, side=TOP, anchor=CENTER)

    if not stringF == None:
        for i in stringF:
            listbox2.insert(END, f'{i}')
           
    global list1
    list1 = genList()
    Update(list1)

    ws.config(menu=menubar)
    ws.attributes('-topmost', True)
    ws.update()
    ws.attributes('-topmost', False)
    ws.mainloop()

# --------------------------------------------------------------------------------------

def config():
    global stringF, radio, durration, web_hook, user, cpu
    try:
        parser = ConfigParser()
        parser.read('config.ini')
        stringF = parser.get("string", "chklist")
        stringF = stringF.split(', ')
        radio = int(parser.get("radio", "radio"))
        durration = int(parser.get("durration", "durration"))
        web_hook = parser.get("Hook", "URL")
        
        if stringF == ['']:
            stringF = None
       
        return stringF, durration, web_hook
    except:
        
        try:
            # savestring = ', '.join(stringF)
            config_file = configparser.ConfigParser()
            config_file["string"]={"chklist": ''}
            config_file["radio"]={"radio": str(radio)}
            config_file["durration"]={"durration": durration}
            config_file["Hook"]={"URL": web_hook}

            with open("config.ini","w") as file_object:
                config_file.write(file_object)
        except Exception as e:
            print(e)
            messagebox.showerror("Load Error", f"Config file could not be loaded, please refer to the following error:\n\n{e}")
         

def save_config():
    try:
        savestring = ', '.join(stringF)
        config_file = configparser.ConfigParser()
        config_file["string"]={"chklist": savestring}
        config_file["radio"]={"radio": str(radio)}
        config_file["durration"]={"durration": durration}
        config_file["Hook"]={"URL": web_hook}

        with open("config.ini","w") as file_object:
            config_file.write(file_object)
    except Exception as e:
        messagebox.showerror("Save Error", f"Config file could not be saved, please refer to the following error:\n\n{e}")


def chk():
    global stringF, web_hook, gate_1
    # while stringF == None or '' in stringF or stringF == ():
    while stringF == ():
        messagebox.showerror("Incomplete Selection", "Please select at least one process from the list")
        listexe()
        if not stringF == ():
            break
        
    save_config()
    global list1  
    if radio == 1:
        comp = []
        
        while True:
            runing = []
            
            for service in stringF:
                try:
                    for process in psutil.process_iter(attrs=['name']):
                        if process.info['name'] == service:
                            runing.append(service)
                            break
                        else:
                            continue
                except:
                    pass

            if len([i for i in stringF if i not in runing]) > 0 and gate_1 == False:
                comp = [i for i in stringF if i not in runing]
                # print(comp)
                gate_1 = True
                # print("Not running", [i for i in stringF if i not in runing])
                # print('message sent')
        
                now = datetime.now()  
                eventTime = now.strftime("%H:%M:%S, %m/%d/%Y")
                myTeamsMessage = pymsteams.connectorcard(web_hook)
                myTeamsMessage.title("Service(s) not detected!")
                myTeamsMessage.text('The following service(s) do not appear to be running as of {0}:   \n___\n{3}   \n___\nLocation-   \nUsername: {1}   \nPC Name: {2}'.format(eventTime, user,cpu,"   \n\u2022 ".join(['', *[i for i in stringF if i not in runing]])))
                myTeamsMessage.color("#DB4035")
                myTeamsMessage.send()
                print(gate_1)

            elif [i for i in stringF if i not in runing] != comp and gate_1 == True:
                # print("Not running", [i for i in stringF if i not in runing])
                # print('comp', comp)
                # print('differences detected')
                gate_1 = False
                print(gate_1)
            
            else:
                print('standby', gate_1)
                            
            time.sleep(durration)

    if radio == 2:
        comp = []
        while True:
            runing = []
            for pid in stringF:
                if psutil.pid_exists(int(pid)) == True:
                    runing.append(pid)
                else:
                    continue
                
            if len([i for i in stringF if i not in runing]) > 0 and gate_1 == False:
                comp = [i for i in stringF if i not in runing]
                # print(comp)
                gate_1 = True
                # print("Not running", [i for i in stringF if i not in runing])
                # print('message sent')
        
                now = datetime.now()  
                eventTime = now.strftime("%H:%M:%S, %m/%d/%Y")
                myTeamsMessage = pymsteams.connectorcard(web_hook)
                myTeamsMessage.title("Service(s) not detected!")
                myTeamsMessage.text('The following service(s) do not appear to be running as of {0}:   \n___\n{3}   \n___\nLocation-   \nUsername: {1}   \nPC Name: {2}'.format(eventTime, user,cpu,"   \n\u2022 ".join(['', *[i for i in stringF if i not in runing]])))
                myTeamsMessage.color("#DB4035")
                myTeamsMessage.send()
                print(gate_1)

            elif [i for i in stringF if i not in runing] != comp and gate_1 == True:
                # print("Not running", [i for i in stringF if i not in runing])
                # print('comp', comp)
                # print('differences detected')
                gate_1 = False
                print(gate_1)
            
            else:
                print('standby', gate_1)
            
            time.sleep(durration)

start_thread_start()
