from tkinter.filedialog import askopenfilename, asksaveasfilename
from statsmodels.tsa.arima_model import ARIMA
from pmdarima.arima import auto_arima
from warnings import filterwarnings
from pandas import DataFrame, Series, read_excel, ExcelWriter
from tkinter import Tk, Label, Button, StringVar, Listbox, MULTIPLE, messagebox, SE
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

def autoArima(data):
    fit = auto_arima(data,
                 test='adf',
                 trace=True,
                 m = 1,
                 error_action='ignore',
                 suppress_warnings=True,
                 stepwise=True,
                 information_criterion='bic')
    model = ARIMA(data, order = fit.order)
    return model.fit()
def forcast(var,fit, predTime):
    lags = len(predTime)
    result = DataFrame(columns=['Var','Pred.Val.', 'Std.Err.','[0.025','0.975]'])
    templist = fit.forecast(lags)
    for i in range(0,lags):
        temp = Series({
                'Var': var,
                'Pred.Val.': templist[0][i],
                'Std.Err.': templist[1][i],
                '[0.025': templist[0][i] - 1.96*templist[1][i],
                '0.975]': templist[0][i] + 1.96*templist[1][i]
                },name = predTime[i])
        result = result.append(temp)
    return result
def fileOpen():
    readFile = askopenfilename(filetypes=[('Excel文件', '*.xlsx')])
    global data 
    data = read_excel(readFile, index_col = 0, header = 0)
    label1['text'] = readFile
    listVar.set(data.columns.tolist())
    listbox.select_set(0, len(data.columns))
def fileSave():
    global saveFile
    saveFile = asksaveasfilename(filetypes=[('Excel文件', '*.xlsx')], defaultextension = ".xlsx")
    label2['text'] = saveFile
def run():
    vars = [x for x in data.columns.tolist() if data.columns.tolist().index(x) in listbox.curselection()]
    result0 = DataFrame(columns = vars)
    if(len(vars) == 0):
        messagebox.showinfo('提示', '至少选中一个变量！')
        return;
    with ExcelWriter(saveFile, engine = "openpyxl") as writer:
        for id, var in enumerate(vars):
            temp = data.iloc[:,id]
            predTime = temp.index[temp.isnull()]
            temp = temp.dropna()
            fit = autoArima(temp)
            result2 = fit.summary2().tables
            result1 = forcast(var, fit, predTime)
            result0[var] = result1['Pred.Val.']
            result0.index = result1.index
            result0.to_excel(writer, sheet_name= "SUMMARY", header = True, index = True)
            result1.to_excel(writer, sheet_name= var, header = True, index = True, index_label = True, 
                                    startrow = 0)
            result2[0].iloc[:,[0,1]].to_excel(writer, sheet_name= var, header = False, index = False, 
                                startrow = result1.shape[0] + 2, startcol = 0)
            result2[0].iloc[:,[2,3]].to_excel(writer, sheet_name= var, header = False, index = False, 
                                startrow = result1.shape[0] + 2, startcol = 5)            
            result2[1].to_excel(writer, sheet_name= var, header = True, index = True, 
                                startrow = result1.shape[0] + result2[0].shape[0] + 3)
            # if(len(result2)>2):
            #     result2[2].to_excel(writer, sheet_name= var, header = True, index = True, index_label = True, 
            #                         startrow = result2[0].shape[0] + result2[1].shape[0] + 40)
        # sheet.write(0, 0, temp2)
        # print(model.fit().summary())
        # result = forcast(model, predTime)
        # print(result)
    writer.save()
    writer.close()
    messagebox.showinfo('提示','执行完成！')
filterwarnings('ignore')
UI = Tk()
UI.title("ARIMA预测器")
UI.resizable(False, False)
sw = UI.winfo_screenwidth()
sh = UI.winfo_screenheight()
ww = 320
wh = 480
x = (sw-ww) / 10
y = (sh-wh) / 10
UI.geometry("%dx%d+%d+%d" %(ww,wh,x,y))
label1 = Label(UI, text="")
label2 = Label(UI, text="")
label3 = Label(UI, text="")
label4 = Label(UI, text="version-1.0.1", anchor = SE)
button1 = Button(UI, text = '打开文件', bd = 1, width = 10, command = fileOpen)
button2 = Button(UI, text = '保存文件', bd = 1, width = 10, command = fileSave)
button3 = Button(UI, text = '执行', bd = 1, width = 10, command = run)
listVar = StringVar()
listbox = Listbox(UI, selectmode = MULTIPLE, listvariable = listVar)

label1.pack()
label1.place(x=10, y=10, width=300, height=40)
button1.pack()
button1.place(x=10, y=50, width=300, height=40)
listbox.pack()
listbox.place(x=10, y=100, width=300, height=200)
label2.pack()
label2.place(x=10, y=310, width=300, height=40)
button2.pack()
button2.place(x=10, y=350, width=300, height=40)
button3.pack()
button3.place(x=10, y=400, width=300, height=40)
label4.pack()
label4.place(x=10, y=450, width=300, height=20)
UI.mainloop()

