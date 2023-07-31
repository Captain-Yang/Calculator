from tkinter import *
import datetime
from tkcalendar import Calendar
import pandas as pd
import time

# 一个月默认天数
Days = 30
# 利息
interest = 0

# 用来记录每一笔本金及其时间
principals = []
dates = []
# 记录最后要输入到excel的数据
dataList = []
# index是指每一条记录
index = 1
# index_borrow用来显示第几笔借款
index_borrow = 1
# index_repay用来显示第几笔还款
index_repay = 1
dataList.append(['序号','借款时间','新增借款金额（元）','还款时间','还款金额（元）','计息期间','计息天数（天）','月息率（%）','应还利息（元）','尚欠本金（元）'])

# 计算利息
def getInterest(end_date):
    # 利息置零
    count = 0.0
    # 从输入框获取月利率,除于100获得小数
    MonRate = (round(float(entry_MonRate.get()),2)) /100
    print("计算利息：月利率是:",MonRate)
    for i in range(len(principals)):
        # 计算每一笔借款的计息天数及利息
        days = (end_date - dates[i]).days
        interest = (MonRate / Days * days * principals[i])
        count = count + interest
        print("利息：", interest)
        print("计息天数:", days)
        print("计息期间：", dates[i], end_date - datetime.timedelta(days=1))
        # 这里写入excel

    print("需要偿还利息为：", count)
    return count


# 计算总共本金
def getPrincipal():
    # 初始值为0
    count = 0.0
    for i in range(len(principals)):
        # 计算总共的本金
        count = count + principals[i]
    print("总共的本金是:",count)
    return count


# 借款
def borrow():
    # 声明全局变量
    global index
    global index_borrow
    global principals
    global dates
    global dataList
    # 获取月利率百分比
    try:
        MonRate = float(entry_MonRate.get())
        if abs(MonRate) < 1e-9 or MonRate < 0:  # 使用一个很小的阈值进行比较,认为等于零，不需要在写入本金
            label_result.config(text=f"月利率必须大于0，请重试！", font=("Arial", 14))
            return
    except:
        label_result.config(text=f"请输入正确的月利率！", font=("Arial", 14))
        return
    print("借款：月利率是:", MonRate)
    # 类型是str
    print("借款的金额是:", entry_principal.get())
    # principal是借款的金额
    try:
        principal = round(float(entry_principal.get()),2)
        if abs(principal) < 1e-9 or principal < 0:  # 使用一个很小的阈值进行比较,认为等于零，不需要在写入本金
            label_result.config(text=f"借款金额必须大于0，请重试！", font=("Arial", 14))
            return
    except:
        label_result.config(text=f"请输入正确的借款金额！", font=("Arial", 14))
        return

    # 将日期转化为datetime对象
    selected_date = cal_begin.get_date()
    begin_date = datetime.datetime.strptime(selected_date, "%Y-%m-%d").date()

    # dates[-1]是上一次借款的时间
    if len(dates) != 0:
        days = (begin_date - dates[-1]).days
        if days <= 0:
            label_result.config(text=f"本次时间应该晚于上一次借款/还款时间，请重试！", font=("Arial", 14))
            return

    principals.append(principal)
    dates.append(begin_date)
    entry_principal.delete(0, END)  # 清空 entry_principal
    # 获取尚欠本金
    count = round(getPrincipal(),2)
    # 需要对selected_date的格式进行处理，将2018-01-02 变成 2018.01.02
    format_selected_date = selected_date.replace("-", ".")
    data = [index,format_selected_date,principal,'','','','',MonRate,'',count]
    dataList.append(data)
    index = index + 1

    label_result.config(text=f"这是第{index_borrow}笔借款，本次借款的金额是: {principal},借款时间是{format_selected_date}",font=("Arial", 14))
    index_borrow = index_borrow + 1

# 还款
def repay():
    # 声明全局变量
    global index
    global index_repay
    global principals
    global dates
    global dataList

    # 获取月利率百分比
    try :
        MonRate = float(entry_MonRate.get())
        if abs(MonRate) < 1e-9 or MonRate < 0:  # 使用一个很小的阈值进行比较,认为等于零，不需要在写入本金
            label_result.config(text=f"月利率必须大于0，请重试！", font=("Arial", 14))
            return
    except :
        label_result.config(text=f"请输入正确的月利率！", font=("Arial", 14))
        return
    print("还款：月利率是:", MonRate)
    # 类型是str
    print("还款的金额是:", entry_repayment.get())
    try:
        repayment = round(float(entry_repayment.get()),2)
        if abs(repayment) < 1e-9 or repayment < 0:  # 使用一个很小的阈值进行比较,认为等于零，不需要在写入本金
            label_result.config(text=f"还款金额必须大于0，请重试！", font=("Arial", 14))
            return
    except:
        label_result.config(text=f"请输入正确的借款金额！", font=("Arial", 14))
        return
    if len(dates) <= 0:
        label_result.config(text=f"目前不需要还款，请重试！",font=("Arial", 14))
        return
    selected_date = cal_end.get_date()
    end_date = datetime.datetime.strptime(selected_date, "%Y-%m-%d").date()
    interest = round(getInterest(end_date),2)

    # 尚欠本金 = 总共的本金 + 利息 - 还款金额
    count = getPrincipal()
    principal = round((count + interest - repayment),2)
    if principal < 0:
        label_result.config(text=f"还款金额超过尚欠金额，请重试！",font=("Arial", 14))
        return
    print("尚欠本金（元）：", principal)

    # dates[0]是上一次还款的时间
    days = (end_date - dates[0]).days
    if days <= 0:
        label_result.config(text=f"本次时间应该晚于上一次借款/还款时间，请重试！", font=("Arial", 14))
        return
    format_selected_date = selected_date.replace("-", ".")
    data = [index, '', '', format_selected_date, repayment, str(dates[0]).replace("-", ".") + '-' + str(end_date - datetime.timedelta(days=1)).replace("-", "."),
            days, MonRate, interest, principal]
    dataList.append(data)
    index = index + 1

    principals = []
    dates = []
    if abs(principal) < 1e-9:  # 使用一个很小的阈值进行比较,认为等于零，不需要在写入本金
        return
    principals.append(principal)
    dates.append(end_date)
    entry_repayment.delete(0, END)  # 清空 entry_repayment

    label_result.config(text=f"这是第{index_repay}笔还款，本次还款的金额是: {repayment},还款时间是{format_selected_date}", font=("Arial", 14))
    index_repay = index_repay + 1

# 保存文件到excel中
def write_to_excel():
   try:
       global dataList
       df = pd.DataFrame(dataList)
       df.to_excel('data.xlsx', index=False, header=False)
       label_result.config(text=f"已保存到Excel文件中", font=("Arial", 14))
   except  Exception as e:
       label_result.config(text=f"无法保存到Excel文件中，请将data.xlsx关闭后再尝试！", font=("Arial", 14))


# def select_date_begin():
#     # selected_date的类型是str
#     selected_date = cal_begin.get_date()
#     print("借款的日期是:", selected_date)
#     # begin_date = datetime.datetime.strptime(selected_date, "%Y-%m-%d").date()
#     # print("借款的日期是:", begin_date)
#
# def select_date_end():
#     selected_date = cal_end.get_date()
#     print("还款的日期是:", selected_date)

root = Tk()
root.title("利息计算器")

# 设置窗口大小和位置
window_width = 600
window_height = 500
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = int((screen_width / 2) - (window_width / 2))
y = int((screen_height / 2) - (window_height / 2))
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

frame_main = Frame(root)
frame_1 = Frame(frame_main)
frame_2 = Frame(frame_main)
frame_3 = Frame(frame_main)
# frame_4用来显示操作信息
frame_4 = Frame(frame_main)

fontsize = 12

label_MonRate_text = Label(frame_1, text="月利率：", font=("Arial", fontsize))

entry_MonRate = Entry(frame_1, width=3,font=("Arial", fontsize))

label_MonRate = Label(frame_1, text="%",font=("Arial", fontsize))

entry_principal = Entry(frame_1, width=7,font=("Arial", fontsize))

label_money1 = Label(frame_1, text="元",font=("Arial", fontsize))

button_principal = Button(frame_1, text="借款", command=borrow,font=("Arial", fontsize))

entry_repayment = Entry(frame_1, width=7,font=("Arial", fontsize))

label_money2 = Label(frame_1, text="元",font=("Arial", fontsize))

button_repayment = Button(frame_1, text="还款", command=repay,font=("Arial", fontsize))

# 创建两个 Calendar 小部件
cal_begin = Calendar(frame_2, selectmode="day", date_pattern="yyyy-mm-dd")
cal_begin.configure(selectbackground="red", selectforeground="white", normalbackground="white", normalforeground="black")
# ,headersforeground='black'
cal_end = Calendar(frame_2, selectmode="day", date_pattern="yyyy-mm-dd")
cal_end.configure(selectbackground="blue", selectforeground="white", normalbackground="white", normalforeground="black")

button_count = Button(frame_3, text="导出为Excel文件", command=write_to_excel,font=("Arial", 12))

label_result = Label(frame_4, text="")


# # 创建一个按钮，用于选择日期
# button_begin = Button(frame_entry, text="选择日期", command=select_date_begin)
# button_end = Button(frame_entry, text="选择日期", command=select_date_end)

# 创建空白的 Frame 组件
frame_space_1 = Frame(frame_main, height=40)
frame_space_2 = Frame(frame_main, height=20)
frame_space_3 = Frame(frame_main, height=20)
frame_space_4 = Frame(frame_main, height=20)

frame1_space_1 = Frame(frame_1, width=10)
frame1_space_2 = Frame(frame_1, width=120)

frame2_space_1 = Frame(frame_2, width=60)

frame3_space_1 = Frame(frame_3, width=350)


# 显示所有组件
frame_main.pack()
frame_space_1.pack(side=TOP)
frame_1.pack(side=TOP)
frame_space_2.pack(side=TOP)
frame_2.pack(side=TOP)
frame_space_3.pack(side=TOP)
frame_3.pack(side=TOP)
frame_space_4.pack(side=TOP)
frame_4.pack(side=LEFT)

#月利率的显示
label_MonRate_text.pack(side=LEFT)
entry_MonRate.pack(side=LEFT)
label_MonRate.pack(side=LEFT)


# 借款的显示
frame1_space_1.pack(side=LEFT)
entry_principal.pack(side=LEFT)
label_money1.pack(side=LEFT)
button_principal.pack(side=LEFT)
frame1_space_2.pack(side=LEFT)  # 空白的 Frame，好像只能用一次

# 还款
entry_repayment.pack(side=LEFT)
label_money2.pack(side=LEFT)
button_repayment.pack(side=LEFT)

# 日期选择组件
cal_begin.pack(side=LEFT)
frame2_space_1.pack(side=LEFT)
# button_begin.pack(side=LEFT)
cal_end.pack(side=LEFT)
# button_end.pack(side=LEFT)

# 导出为excel文件
frame3_space_1.pack(side=LEFT)
button_count.pack(side=LEFT)
# 输出操作信息
label_result.pack(side=LEFT)

root.mainloop()

time.sleep(5)