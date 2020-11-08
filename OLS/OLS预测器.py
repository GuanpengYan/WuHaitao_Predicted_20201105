from tkinter.filedialog import askopenfilename, asksaveasfilename
from warnings import filterwarnings
from pandas import DataFrame, read_excel, ExcelWriter
from tkinter import Tk, Label, Button, StringVar, Listbox, MULTIPLE, messagebox, SW, SE
from statsmodels.api import OLS, add_constant
import sklearn.neighbors.typedefs
import statsmodels.tsa.statespace._filters
import statsmodels.tsa.statespace._filters._conventional
import statsmodels.tsa.statespace._filters._univariate
import statsmodels.tsa.statespace._filters._univariate_diffuse
import statsmodels.tsa.statespace._filters._inversions
import statsmodels.tsa.statespace._smoothers
import statsmodels.tsa.statespace._smoothers._conventional
import statsmodels.tsa.statespace._smoothers._univariate
import statsmodels.tsa.statespace._smoothers._univariate_diffuse
import statsmodels.tsa.statespace._smoothers._classical
import statsmodels.tsa.statespace._smoothers._alternative

def fileOpen():
    readFile = askopenfilename(filetypes=[('Excel文件', '*.xlsx')])
    label1['text'] = readFile
    global data
    data = read_excel(readFile, index_col = 0, header = 0)
    global Y
    Y = data.loc[:, data.isnull().any()]
    listVarY.set(Y.columns.tolist())
    listboxY.select_set(0, len(Y))
    global X
    X = data.loc[:, ~data.isnull().any()]
    listVarX.set(X.columns.tolist())
    listboxX.select_set(0, len(X))
def fileSave():
    global saveFile
    saveFile = asksaveasfilename(filetypes=[('Excel文件', '*.xlsx')], defaultextension = ".xlsx")
    label2['text'] = saveFile
def run():
    varsY = [x for x in Y.columns.tolist() if Y.columns.tolist().index(x) in listboxY.curselection()]
    varsX = [x for x in X.columns.tolist() if X.columns.tolist().index(x) in listboxX.curselection()]
    global trainY
    global trainX
    trainY = Y[~ data.isnull().T.any().T]
    trainX = X[~ data.isnull().T.any().T]
    trainX = add_constant(trainX[varsX])
    testX = X[data.isnull().T.any().T]
    testX = add_constant(testX[varsX])
    result0 = DataFrame(columns = varsY)
    if(len(varsY) == 0):
        messagebox.showinfo('提示', '至少选中一个结果变量！')
        return;
    if(len(varsX) == 0):
        messagebox.showinfo('提示', '至少选中一个预测变量！')
        return;
    with ExcelWriter(saveFile, engine = "openpyxl") as writer:
        for id, varY in enumerate(varsY):
            fit = OLS(trainY.iloc[:, id], trainX).fit()
            print(fit.summary2().tables)
            result0[varY] = fit.predict(testX)
            result0.to_excel(writer, sheet_name= "SUMMARY", header = True, index = True)
            global result1
            result1 = fit.get_prediction(testX).summary_frame()
            result1.to_excel(writer, sheet_name= varY, header = True, index = True)
            global result2
            result2 = fit.summary2().tables
            result2[0].iloc[:,[0,1]].to_excel(writer, sheet_name= varY, header = False, index = False,
                                startrow = result1.shape[0] + 2, startcol = 0)
            result2[0].iloc[:,[2,3]].to_excel(writer, sheet_name= varY, header = False, index = False,
                                startrow = result1.shape[0] + 2, startcol = 5)
            result2[1].to_excel(writer, sheet_name= varY, header = True, index = True,
                                startrow = result1.shape[0] + result2[0].shape[0] + 3)
    writer.save()
    writer.close()
    messagebox.showinfo('提示', '执行完成！')

filterwarnings('ignore')
UI = Tk()
UI.title("OLS预测器")
UI.resizable(False, False)
sw = UI.winfo_screenwidth()
sh = UI.winfo_screenheight()
ww = 320
wh = 530
x = (sw-ww) / 10
y = (sh-wh) / 10
UI.geometry("%dx%d+%d+%d" %(ww,wh,x,y))
label1 = Label(UI, text="")
label2 = Label(UI, text="")
label3 = Label(UI, text="")
label4 = Label(UI, text="结果变量：", anchor=SW)
label5 = Label(UI, text="预测变量：", anchor=SW)
label6 = Label(UI, text="version 1.0.1", anchor = SE)
button1 = Button(UI, text = '打开文件', bd = 1, width = 10, command = fileOpen)
button2 = Button(UI, text = '保存文件', bd = 1, width = 10, command = fileSave)
button3 = Button(UI, text = '执行', bd = 1, width = 10, command = run)
listVarY = StringVar()
listboxY = Listbox(UI, selectmode = MULTIPLE, listvariable = listVarY, exportselection = False)
listVarX = StringVar()
listboxX = Listbox(UI, selectmode = MULTIPLE, listvariable = listVarX, exportselection = False)
label1.pack()
label1.place(x=10, y=10, width=300, height=40)
button1.pack()
button1.place(x=10, y=50, width=300, height=40)
label4.pack()
label4.place(x=10, y=100, width=140, height=40)
label5.pack()
label5.place(x=170, y=100, width=140, height=40)
listboxY.pack()
listboxY.place(x=10, y=150, width=140, height=200)
listboxX.pack()
listboxX.place(x=170, y=150, width=140, height=200)
label2.pack()
label2.place(x=10, y=360, width=300, height=40)
button2.pack()
button2.place(x=10, y=400, width=300, height=40)
button3.pack()
button3.place(x=10, y=450, width=300, height=40)
label6.pack()
label6.place(x=10, y=500, width=300, height=20)
UI.mainloop()