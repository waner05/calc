import tkinter as tk
from tkinter import *
from tkinter import ttk, messagebox
import xlwt
from xlwt import Workbook
import matplotlib.pyplot as plt
from tkinter import simpledialog
import subprocess
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
import sympy as sp

def about():
    new = Toplevel(gui)
    new.title("About")
    new.geometry("400x150")
    new.resizable(False, False)
    label = tk.Label(new, text="This is a basic sequential calculator that can save it's" \
    " operations into an Excel file. Pressing 'C' or Clear clears the calculator output and inputs and also moves any future operations into the next row of the Excel sheet. " \
    "Pressing 'S' or Save exports the Excel sheet. Note that order of operations do not apply, all calculations are done the instant the an operation button is " \
    "clicked. Press G to enter a function to graph, and you can save your graph into a png file if requested.", wraplength=400, justify="left")
    label.pack()

def click(number):
    current = output.get()
    output.config(state = 'normal')
    if number == '.' and '.' in current:
        return
    elif number == '-':
        if current.startswith('-'):
            output.delete(0,tk.END)
            output.insert(0,current[1:])
        else:
            output.delete(0,tk.END)
            output.insert(0,'-' + current)
    else:
        output.delete(0, tk.END)
        output.insert(0, current + str(number))
    output.config(state='readonly')

def clear():
    output.config(state = 'normal')
    output.delete(0, tk.END)
    output.config(state='readonly')

def fullclear():
    global operation, num_state, last_eq, count, row
    clear()
    operation = 0
    num_state = 0
    last_eq = False
    count = 0
    row += 1

def set_operation(opcode):
    global num_state, operation, last_eq, count
    result = 0
    new_num = output.get()
    if count == 0:
        sheet.write(row, count, new_num, style)
        count += 1    
    if last_eq:
        num_state = new_num
        last_eq = False
    elif operation != 0:
        finish(result)
    num_state = output.get()
    operation = opcode
    clear()

def finish(result):
    global num_state
    global operation
    global count
    new_num = output.get()
    result = 0
    if operation == 1:
        output.config(state='normal')
        output.delete(0, tk.END)
        result = float(num_state) + float(new_num)
        output.insert(0, trailing_zero(result))
        sheet.write(row, count, '+', style)
        output.config(state='readonly')
    elif operation == 2:
        output.config(state='normal')
        output.delete(0, tk.END)
        result = float(num_state) - float(new_num)
        output.insert(0, trailing_zero(result))
        sheet.write(row, count, '-', style)
        output.config(state='readonly')
    elif operation == 3:
        output.config(state='normal')
        output.delete(0, tk.END)
        result = float(num_state)*float(new_num)
        output.insert(0, trailing_zero(result))
        sheet.write(row, count, '*', style)
        output.config(state='readonly')
    elif operation == 4:
        output.config(state='normal')
        output.delete(0, tk.END)
        result = float(num_state)/float(new_num)
        output.insert(0, trailing_zero(result))
        sheet.write(row, count, '/', style)
        output.config(state='readonly')
    count += 1
    sheet.write(row, count, new_num, style)
    count += 1
    return result
    
def equals():
    global count, num_state, last_eq, last_operand, last_operation, operation

    new_num = output.get()
    if not last_eq:
        last_operand = new_num
        last_operation = operation
    else:
        operation = last_operation
        output.config(state='normal')
        output.delete(0, tk.END)
        output.insert(0, last_operand)
        output.config(state='readonly')
    result = finish(0)
    sheet.write(row, count, '=', style)
    count += 1
    sheet.write(row, count, result, style)
    count += 1
    num_state = result
    last_eq = True

def savefile():
    wb.save('CalculatorHistory.xls')
    messagebox.showinfo("Saved", "Calculations saved as CalculatorHistory.xls")

def oldcalc():
    subprocess.run(['python','testinterface.py'])

def trailing_zero(value):
    if float(value).is_integer():
        return str(int(value))
    else:
        return str(value)

def graph_mode():
    for widget in gui.winfo_children():
        if widget != menu and widget !=output:
            widget.place_forget()
    
    output.config(state='normal')
    output.delete(0,tk.END)
    output.insert(0,"")
    output.focus_set()

    lut = Label(gui, image=imageLook,border=0)
    lut.place(x=30,y=80)
    graph_confirm = tk.Button(gui,image=imageGraphC,bg="#11C0FD",activebackground="#11C0FD",border=0,command=lambda: plot_out(graph_confirm, cancel,lut))
    graph_confirm.place(x=225,y=90)

    cancel = tk.Button(gui,image=imageCancel,bg="#11C0FD",activebackground="#11C0FD",border=0,command=lambda: graph_cancel(graph_confirm, cancel,lut))
    cancel.place(x=225,y=130)

    
    
def graph_cancel(graph_confirm, graph_cancel,lut):
    graph_confirm.destroy()
    graph_cancel.destroy()
    lut.destroy()
    clear()
    rebuild_ui()

def plot_out(graph_confirm, graph_cancel,lut):

    equation = output.get()
    clear()
    try:
        x = np.linspace(-10,10,400)
        result = eval(equation, {"x": x, "np": np, "sin": np.sin, "cos": np.cos, "tan": np.tan,"exp": np.exp, "log": np.log, "sqrt": np.sqrt, "abs": np.abs})
        y = result if isinstance(result, np.ndarray) else np.full_like(x, result)

        graph_confirm.destroy()
        graph_cancel.destroy()
        lut.destroy()
        output.place_forget()

        fig = plt.Figure(figsize=(3.5,3), dpi=100)
        ax = fig.add_subplot(111)
        ax.plot(x,y)
        ax.set_title(f"y={equation}")
        ax.grid(True)
        graph = FigureCanvasTkAgg(fig,master=gui)
        graph.get_tk_widget().place(x=10,y=10)
        back = tk.Button(gui,image=imageBack,bg="#11C0FD",activebackground="#11C0FD",border=0,command=lambda:(back_calc(graph, back, save)))
        back.place(x=270,y=315)

        save = tk.Button(gui,image=imageSave2,bg="#11C0FD",activebackground="#11C0FD",border=0,command=lambda: save_graph(fig))
        save.place(x=100, y=315)
    except Exception as e:
        messagebox.showerror("Invalid Function","Please enter a valid function")
        output.config(state='normal')
def back_calc(graph, back,save):
    graph.get_tk_widget().destroy()
    back.destroy()
    save.destroy()
    rebuild_ui()

def rebuild_ui():
    one.place(x=20,y=85)
    two.place(x=80,y=85)
    three.place(x=140,y=85)
    four.place(x=20,y=145)
    five.place(x=80,y=145)
    six.place(x=140,y=145)
    seven.place(x=20,y=205)
    eight.place(x=80,y=205)
    nine.place(x=140,y=205)
    zero.place(x=80,y=265)
    decimal.place(x=140, y=265)
    posneg.place(x=20,y=265)
    addb.place(x=220, y=85)
    subb.place(x=220, y=145)
    multb.place(x=220, y=205)
    divb.place(x=220, y=265)
    equal.place(x=280, y=85)
    clc.place(x=280,y=145)
    saveButton.place(x=280,y=265)
    graphButton.place(x=280,y=205)
    mathButton.place(x=280,y=235)
    output.place(x=20, y=0, height=80,width=312)
    note.place(x=2,y=330)
    bar_canvas.place(x=206, y=85)


def save_graph(fig):
    fig.savefig("graph.png")
    messagebox.showinfo("Saved", "Graph saved as graph.png")

def eq_mode():
    for widget in gui.winfo_children():
        if widget != menu and widget != output:
            widget.place_forget()
    clear()
    output.config(state='normal')
    output.delete(0,tk.END)
    output.focus_set()


    simplify_btn = tk.Button(gui, image=imageSimp,bg="#11C0FD",activebackground="#11C0FD",border=0,command=lambda: sym_op("simplify"))
    simplify_btn.place(x=10, y=100)

    diff_btn = tk.Button(gui, image=imageDiff,bg="#11C0FD",activebackground="#11C0FD",border=0,command=lambda: sym_op("diff"))
    diff_btn.place(x=10, y=180)

    integrate_btn = tk.Button(gui, image=imageInt,bg="#11C0FD",activebackground="#11C0FD",border=0,command=lambda: sym_op("integrate"))
    integrate_btn.place(x=10, y=220)

    solve_btn = tk.Button(gui, image=imageSolve,bg="#11C0FD",activebackground="#11C0FD",border=0,command=lambda: sym_op("solve"))
    solve_btn.place(x=10, y=140)

    lut2 = Label(gui, image=imageEqSheet,border=0)
    lut2.place(x=145,y=85)

    back_btn = tk.Button(gui, image=imageBack,bg="#11C0FD",activebackground="#11C0FD",border=0, command=lambda: sym_back(
        [simplify_btn, diff_btn, integrate_btn, solve_btn, back_btn, lut2]))
    back_btn.place(x=10, y=300)



def sym_op(op):
    ask = output.get()
    x = sp.symbols('x')
    try:
        if op == "simplify":
            result = sp.simplify(ask)
        elif op == "diff":
            result = sp.diff(ask, x)
        elif op == "integrate":
            result = sp.integrate(ask, x)
        elif op == "solve":
            result = sp.solve(ask, x)

        output.config(state='normal')
        output.delete(0,tk.END)
        output.insert(0,result)
    except Exception as e:
        messagebox.showerror("Error", "Invalid expression")

def sym_back(buttons):
    for btn in buttons:
        btn.destroy()
    clear()
    rebuild_ui()

operation = 0 #1 = add, 2 = sub, 3 = mult, 4 = div
num_state = 0
count = 0
row = 0
last_operand = 0
last_operation = 0
last_eq = False

wb = Workbook()
sheet = wb.add_sheet('Operations')
style = xlwt.XFStyle()
alignment = xlwt.Alignment()
alignment.horz = xlwt.Alignment.HORZ_CENTER
style.alignment = alignment

gui = tk.Tk()
gui.title('Calculator v1.5')
gui.geometry("350x370")
gui.resizable(False, False)
gui.iconbitmap('images/calc.ico')

image1 = PhotoImage(file='images/1.png')
image2 = PhotoImage(file='images/2.png')
image3 = PhotoImage(file='images/3.png')
image4 = PhotoImage(file='images/4.png')
image5 = PhotoImage(file='images/5.png')
image6 = PhotoImage(file='images/6.png')
image7 = PhotoImage(file='images/7.png')
image8 = PhotoImage(file='images/8.png')
image9 = PhotoImage(file='images/9.png')
image0 = PhotoImage(file='images/0.png')
imageDot = PhotoImage(file='images/dot.png')
imagePlus = PhotoImage(file='images/plus.png')
imageMinus = PhotoImage(file='images/minus.png')
imageMult = PhotoImage(file='images/mult.png')
imageDiv = PhotoImage(file='images/div.png')
imageEq = PhotoImage(file='images/equals.png')
imageClc = PhotoImage(file='images/clear.png')
imageSave = PhotoImage(file='images/save.png')
imageGraph = PhotoImage(file = 'images/graph.png')
imageBack = PhotoImage(file = 'images/backbtn.png')
imageSave2 = PhotoImage(file = 'images/save2.png')
imageGraphC = PhotoImage(file = 'images/graph1.png')
imageCancel = PhotoImage(file = 'images/Cancel.png')
imageLook = PhotoImage(file = "images/lut.png")
imagePN = PhotoImage(file = 'images/posneg.png')
imageSolve = PhotoImage(file = 'images/solve.png')
imageSimp = PhotoImage(file = 'images/simp.png')
imageDiff = PhotoImage(file = "images/diff.png")
imageInt = PhotoImage(file = 'images/int.png')
imageEqSheet = PhotoImage(file = 'images/eqsheet.png')
imageMath = PhotoImage(file = 'images/mathmode.png')

one = tk.Button(gui, image=image1,bg="#11C0FD",activebackground="#11C0FD",border=0, command=lambda: click(1))
two = tk.Button(gui, image=image2, bg="#11C0FD",activebackground="#11C0FD",border=0, command=lambda: click(2))
three = tk.Button(gui, image=image3, bg="#11C0FD",activebackground="#11C0FD",border=0, command=lambda: click(3))
four = tk.Button(gui, image=image4, bg="#11C0FD",activebackground="#11C0FD",border=0, command=lambda: click(4))
five = tk.Button(gui,image=image5, bg="#11C0FD",activebackground="#11C0FD",border=0, command=lambda: click(5))
six = tk.Button(gui, image=image6,bg="#11C0FD",activebackground="#11C0FD",border=0, command=lambda: click(6))
seven = tk.Button(gui, image=image7,bg="#11C0FD",activebackground="#11C0FD",border=0, command=lambda: click(7))
eight = tk.Button(gui, image=image8,bg="#11C0FD",activebackground="#11C0FD",border=0, command=lambda: click(8))
nine = tk.Button(gui, image=image9,bg="#11C0FD",activebackground="#11C0FD",border=0, command=lambda: click(9))
zero = tk.Button(gui, image=image0,bg="#11C0FD",activebackground="#11C0FD",border=0, command=lambda: click(0))
decimal = tk.Button(gui, image=imageDot,bg="#11C0FD",activebackground="#11C0FD",border=0, command=lambda: click('.'))
posneg = tk.Button(gui, image=imagePN,bg="#11C0FD",activebackground="#11C0FD",border=0, command=lambda: click('-'))
addb = tk.Button(gui, image=imagePlus,bg="#11C0FD",activebackground="#11C0FD",border=0, command=lambda: set_operation(1))
subb = tk.Button(gui, image=imageMinus,bg="#11C0FD",activebackground="#11C0FD",border=0, command=lambda: set_operation(2))
multb = tk.Button(gui, image=imageMult,bg="#11C0FD",activebackground="#11C0FD",border=0, command=lambda: set_operation(3))
divb = tk.Button(gui, image=imageDiv,bg="#11C0FD",activebackground="#11C0FD",border=0, command=lambda: set_operation(4))
equal = tk.Button(gui,image=imageEq,bg="#11C0FD",activebackground="#11C0FD",border=0, command=equals)
clc = tk.Button(gui,image=imageClc,bg="#11C0FD",activebackground="#11C0FD",border=0,command=fullclear)
saveButton = tk.Button(gui, image=imageSave,bg="#11C0FD",activebackground="#11C0FD",border=0,command=savefile)
graphButton = tk.Button(gui, image=imageGraph,bg="#11C0FD",activebackground="#11C0FD",border=0,command=graph_mode)
mathButton = tk.Button(gui, image=imageMath,bg="#11C0FD",activebackground="#11C0FD",border=0,command=eq_mode)

output = tk.Entry(gui, state='readonly', readonlybackground="#D5E5DA",bg="#D5E5DA", font=('Consolas',24,'bold'), justify='right')

note = tk.Label(gui, text='Note: OOP does not apply, operations are done instantaneously',bg="#11C0FD")
menu = Menu(gui)
gui.config(menu=menu)
gui.config(bg="#11C0FD")
filemenu = Menu(menu, tearoff="off")
menu.add_cascade(label='File',menu=filemenu)
filemenu.add_command(label='Graphing Mode', command=graph_mode)
filemenu.add_command(label='Symbolic Solver',command=eq_mode)
filemenu.add_command(label='Open Old Calculator', command=oldcalc)
filemenu.add_command(label='Exit',command=gui.destroy)


exportmenu = Menu(menu,tearoff="off")
menu.add_cascade(label='Export', menu=exportmenu)
exportmenu.add_command(label='to .xls', command=savefile)

helpmenu = Menu(menu,tearoff="off")
menu.add_cascade(label='Help', menu=helpmenu)
helpmenu.add_command(label='About', command=about)

bar_canvas = tk.Canvas(gui, width=2, height=230, bg="#1889B3", highlightthickness=0)
bar_canvas.place(x=206, y=85)
rebuild_ui()



gui.mainloop()