#!/usr/bin/env python3
# -*- coding:utf-8 -*-
import cx_Oracle, os, openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, colors, Alignment, Border
import tkinter as tk
from tkinter import ttk
import datetime
g_font = ('Monaco', 12)

kb = []
result = []
teachercount = []
# 1、功能函数
def connDB():
    global kb, max_jsz, sub_control2_qszbox, sub_control2_jszbox, sub_control3_enq
    conn = cx_Oracle.connect('zfxfzb/zfsoft_hqwy@orcl')
    cur = conn.cursor()
    sql_kb = 'SELECT * from kb_allinfor'
    cur.execute(sql_kb)
    for row in cur:
        kb.append(row)
    cur.close()

    # 当学期最大结束周
    cur = conn.cursor()
    sql_kb = 'SELECT max(jsz) from kb_allinfor'
    cur.execute(sql_kb)
    max_jsz = [i for i in cur][0][0]
    cur.close()

    sub_control2_qszbox['values'] = list(range(1, max_jsz + 1))
    sub_control2_jszbox['values'] = list(range(1, max_jsz + 1))

    # print('连接成功', end='\n')
    sub_control1.configure(text='连接成功')  # 设置button显示的内容
    sub_control1.configure(state='disabled')  # 将按钮设置为灰色状态，不可使用状态
    sub_control3_enq.configure(state='normal')


def getnum():
    global kb, result, teachercount, sub_control3_enqtext, sub_control3_enqop
    input = [sub_control2_xqbox.get(), sub_control2_kjbox.get(), sub_control2_qszbox.get(), sub_control2_jszbox.get(), sub_control2_dszbox.get()]
    prar = {}
    if input[0] == '1-5':
        prar['xqj'] = [1, 2, 3, 4, 5]
    elif input[0] == '1-6' or input[0] == '':
        prar['xqj'] = [1, 2, 3, 4, 5, 6]
    else:
        prar['xqj'] = [int(input[0])]
    if input[1] == '白天' or input[1] == '':
        prar['sjd'] = ['1', '3', '5', '7']
    elif input[1] == '上午':
        prar['sjd'] = ['1', '3']
    elif input[1] == '下午':
        prar['sjd'] = ['5', '7']
    else:
        prar['sjd'] = [input[1][0]]


    if input[2] == '':
        prar['qsz'] = 20
    else:
        prar['qsz'] = int(input[2])
    # prar['jsz'] = int(input[3])
    if input[3] == '':
        prar['jsz'] = 1
    else:
        prar['jsz'] = int(input[3])


    if input[4] == '单周':
        prar['dsz'] = [None, '单']
    elif input[4] == '双周':
        prar['dsz'] = [None, '双']
    else:
        prar['dsz'] = [None, '单', '双']

    result = []
    teachercount = []

    for i in kb:
        if i[3] in prar['xqj'] and str(i[4]) in prar['sjd'] and i[5] in prar['dsz'] and i[6] <= prar['qsz'] and i[7] >= \
                prar['jsz']:
            result.append(i)
            teachercount.append(i[2])

    sub_control3_enqtext.configure(state='normal')
    sub_control3_enqtext.configure(text='该时间段有课教师有: ' + str(len(set(teachercount))) + '人', font=g_font)
    sub_control3_enqop.configure(state='normal')


def outtoexcel():
    global result
    wb = openpyxl.Workbook()  # 打开新的工作簿
    table = wb.worksheets[0]
    table.title = '有课教师信息'

    # 写入表头
    col = 1
    for tit in ('序号','教师','课程名称','星期几','第几节','地点','单双周','起始结束周','课程性质','开课学院'):
        table[get_column_letter(col) + '1'] = tit
        col += 1

    # 设置格式
    table.column_dimensions['A'].width = 7
    table.column_dimensions['C'].width = 36
    table.column_dimensions['F'].width = 10
    table.column_dimensions['H'].width = 12
    table.column_dimensions['I'].width = 14
    table.column_dimensions['J'].width = 18

    alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)  # 上下左右居中,自动换行
    for column in range(1,11):
        table[get_column_letter(column) + '1'].alignment = alignment


    # 写入数据
    # ('2018-2019', '2', '13022', 2, 3, None, 10, 14, 2, '(2018-2019-2)-00014718-13022-1', '2002005014', 'qt', '中心-303',
    #  '姜丽媛', 'ERP业务实训2', '综合技能训练', '国际经济贸易学院')
    rownum = 2
    for row in result:
        table['A' + str(rownum)] = rownum-1
        table['B' + str(rownum)] = row[13]
        table['C' + str(rownum)] = row[14]
        table['D' + str(rownum)] = '星期' + str(row[3])
        if row[8] == 1:
            table['E' + str(rownum)] = str(row[4]) + '节'
        else:
            table['E' + str(rownum)] = str(row[4]) + '、' + str(row[4]+1) + '节'
        table['F' + str(rownum)] = row[12]
        table['G' + str(rownum)] = row[5]
        table['H' + str(rownum)] = str(row[6]) + '-' + str(row[7]) + '周'
        table['I' + str(rownum)] = row[-2]
        table['J' + str(rownum)] = row[-1]
        rownum += 1

     # 保存excel
    if ('有课教师信息.xlsx') not in os.listdir('.'):  # 如果本地目录下有重名文件，则在文件名后面加"(n)"，n顺次加1
        fname = '有课教师信息.xlsx'
    else:
        fnum = 1
        while True:
            fnamee = '有课教师信息' + '(' + str(fnum) + ')' + '.xlsx'
            if fnamee not in os.listdir('.'):
                fname = fnamee
                break
            else:
                fnum += 1
    wb.save(filename=fname)

    out_end_text = tk.Label(tch_box3, text='输出完成，请查看: \n' + '《' + fname + '》', font=g_font)
    out_end_text.pack(fill='both', expand=1, padx=2, pady=2, side=tk.BOTTOM)


gra_info = []
def grainfo_connDB():
    global gra_info, grainfo_crl1, grainfo_yeart
    conn = cx_Oracle.connect('zfxfzb/zfsoft_hqwy@orcl')
    cur = conn.cursor()
    sql_kb = 'SELECT XH,XM,XZB,RXRQ,SFZH,YWXM,YWCSD,XB,XY,ZYMC,DQSZJ,BDH FROM GRA_INFO'
    cur.execute(sql_kb)
    for row in cur:
        gra_info.append(row)
    cur.close()
    conn.close()

    grainfo_yeart.insert(-1, datetime.datetime.now().year)  # 设置Entry控件的默认值
    grainfo_crl1.configure(text='连接成功')  # 设置button显示的内容
    grainfo_crl1.configure(state='disabled')  # 将按钮设置为灰色状态，不可使用状态
    grainfo_bu.configure(state='normal')
    grainfo_but.configure(state='normal')

def gra_search(ev=None):
    global grainfo_termsc, grainfo_termse, gra_info, grainfo_box3, resultForms
    grainfo_box3.pack(fill="both", expand=1, padx=2, side=tk.BOTTOM)
    s_result = []
    if grainfo_termsc.get() == '学号':
        for stu in gra_info:
            if stu[0] == grainfo_termse.get():
                s_result.append(stu[:7])
    elif grainfo_termsc.get() == '姓名':
        for stu in gra_info:
            if stu[1] == grainfo_termse.get():
                s_result.append(stu[:7])
    elif grainfo_termsc.get() == '身份证号':
        for stu in gra_info:
            if stu[8] == grainfo_termse.get():
                s_result.append(stu[:7])

    # 清空Treeview中的内容
    x = resultForms.get_children()
    for item in x:
        resultForms.delete(item)

    # 将新查询的数据放入Treeview
    for i in s_result:
        resultForms.insert('', i[0], text=i[0], values=i[0:])

def create_grainfo():
    global grainfo_rulesc, grainfo_yesrt
    '''
    将grainfo_yesrt中选定毕业时间的毕业生插入table gra_info中，并将新数据按grainfo_rulesc中的规则生成毕业证和学位证
    '''
    pass



# 2、切换界面函数，选择功能后隐藏其他界面，显示自己界面
def btn_click_0(event=None):
    global frm_content_1
    btn_text = event.widget['text']
    if btn_text == '毕业资格审核':
        tch_box0.pack_forget()
        signupinfo_box0.pack_forget()
        grainfo_box0.pack_forget()
        frm_content_gra.pack(fill="both", expand=1, padx=2, side=tk.BOTTOM)

    elif btn_text == '有课教师查询':
        frm_content_gra.pack_forget()
        signupinfo_box0.pack_forget()
        grainfo_box0.pack_forget()
        tch_box0.pack(fill="both", expand=1, padx=2, side=tk.BOTTOM)

    elif btn_text == '报名信息查询':
        frm_content_gra.pack_forget()
        grainfo_box0.pack_forget()
        tch_box0.pack_forget()
        signupinfo_box0.pack(fill="both", expand=1, padx=2, side=tk.BOTTOM)


    elif btn_text == '毕业生信息查询':
        frm_content_gra.pack_forget()
        signupinfo_box0.pack_forget()
        tch_box0.pack_forget()
        grainfo_box0.pack(fill="both", expand=1, padx=2, side=tk.BOTTOM)




# 3、基本界面布局
root = tk.Tk()
root.title('吉林外国语大学——教务管理系统补充程序')
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

# 4、底层控件frm
frm = tk.Frame(root)
frm.pack(fill="both", expand=1)

# 5、菜单控件布局
frm_menu = tk.LabelFrame(frm)
frm_menu.pack(fill="both", expand=1, padx=2, side=tk.TOP)

menu_list = ['毕业资格审核', '有课教师查询', '报名信息查询', '毕业生信息查询']
for index, value in enumerate(menu_list):
    menu_button = tk.Button(frm_menu, anchor="w", text=value, font=g_font)
    menu_button.bind('<Button-1>', btn_click_0)
    menu_button.pack(fill="both", expand=1, padx=2, pady=2, side=tk.LEFT)


# 6、功能界面布局
frm_content_gra = tk.LabelFrame(frm)  # 功能界面底层容器-毕业资格审核
frm_content_gra.pack(fill="both", expand=1, padx=2, side=tk.BOTTOM)
sub_graduate1 = tk.Button(frm_content_gra, text='1、点我连接数据课', font=g_font)
sub_graduate1.pack(fill='both', expand=1, padx=2, pady=2)
frm_content_gra.pack_forget()


tch_box0 = tk.LabelFrame(frm)  # 功能界面底层容器-有课教师查询
tch_box0.pack(fill="both", expand=1, padx=2, side=tk.BOTTOM)
tch_text = tk.Label(tch_box0, text='本功能可按时间查询有课教师\n请按照1、2、3步完成操作。', font=g_font)
tch_text.pack(fill='both', expand=1, padx=2, pady=2, side=tk.TOP)

tch_box1 = tk.LabelFrame(tch_box0)
tch_box1.pack(fill="both", expand=1, padx=2)
sub_control1 = tk.Button(tch_box1, text='1、点我连接数据库', font=g_font, command=connDB)
sub_control1.pack(fill='both', expand=1, padx=2, pady=2)

tch_box2 = tk.LabelFrame(tch_box0)
tch_box2.pack(fill="both", expand=1, padx=2)
sub_control2 = tk.Label(tch_box2, text='2、选择时间段', font=g_font)
sub_control2.pack(fill="both", expand=1, padx=2)

tch_box21 = tk.LabelFrame(tch_box2)
tch_box21.pack(fill="both", expand=1, padx=2)
sub_control2_xq = tk.Label(tch_box21, text='星期：', font=g_font)
sub_control2_xq.pack(padx=10, pady=2, side=tk.LEFT, expand='no')
number = tk.StringVar()
sub_control2_xqbox = ttk.Combobox(tch_box21, textvariable=number)
sub_control2_xqbox['values'] = (1, 2, 3, 4, 5, 6, '1-5', '1-6')
sub_control2_xqbox.pack(padx=2, pady=2, side=tk.LEFT)

sub_control2_kj = tk.Label(tch_box21, text='课节：', font=g_font)
sub_control2_kj.pack(padx=10, pady=2, side=tk.LEFT)
number = tk.StringVar()
sub_control2_kjbox = ttk.Combobox(tch_box21, textvariable=number)
sub_control2_kjbox['values'] = ('1、2节', '3、4节', '5、6节', '7、8节', '9、10节', '上午', '下午', '白天')
sub_control2_kjbox.pack(padx=2, pady=2, side=tk.LEFT)


tch_box22 = tk.LabelFrame(tch_box2)
tch_box22.pack(fill="both", expand=1, padx=2)
sub_control2_qzz = tk.Label(tch_box22, text='周次：', font=g_font)
sub_control2_qzz.pack(fill='both', padx=10, pady=3, side=tk.LEFT)
number = tk.StringVar()
sub_control2_qszbox = ttk.Combobox(tch_box22, textvariable=number)
sub_control2_qszbox.pack(padx=2, pady=3, side=tk.LEFT)
sub_control2_qzz = tk.Label(tch_box22, text=' ——  ', font=g_font)
sub_control2_qzz.pack(padx=10, pady=3, side=tk.LEFT)
number = tk.StringVar()
sub_control2_jszbox = ttk.Combobox(tch_box22, textvariable=number)
sub_control2_jszbox.pack(padx=0, pady=3, side=tk.LEFT)


tch_box23 = tk.LabelFrame(tch_box2)
tch_box23.pack(fill="both", expand=1, padx=2)
sub_control2_dsz = tk.Label(tch_box23, text='单双周：', font=g_font)
sub_control2_dsz.pack(fill='both', padx=2, pady=3, side=tk.LEFT)
sub_control2_dszbox = ttk.Combobox(tch_box23)
sub_control2_dszbox['values'] = ('', '单周', '双周')
sub_control2_dszbox.pack(fill='both', padx=2, pady=3, side=tk.LEFT)


tch_box3 = tk.LabelFrame(tch_box0)  # 查询结果控件的容器
tch_box3.pack(fill="both", expand=1, padx=2)
sub_control3_enq = tk.Button(tch_box3, text='3、点击查询', font=g_font, command=getnum)
sub_control3_enq.pack(fill='both', expand=1, padx=2, pady=3, side=tk.TOP)
sub_control3_enq.configure(state='disable')  # 暂时设为灰色，查询成功后显示
sub_control3_enqtext = tk.Label(tch_box3, text='查询结果', font=g_font)
sub_control3_enqtext.pack(fill='both', expand=1, padx=2, pady=2)
sub_control3_enqtext.configure(state='disable')  # 暂时设为灰色，查询成功后显示
sub_control3_enqop = tk.Button(tch_box3, text='详细名单输出到Excel', font=g_font, command=outtoexcel)
sub_control3_enqop.pack(fill='both', expand=1, padx=2, pady=3)
sub_control3_enqop.configure(state='disable')  # 暂时设为灰色，查询成功后显示

tch_box0.pack_forget()


signupinfo_box0 = tk.LabelFrame(frm)   # 功能界面底层容器-报名信息查询
signupinfo_box0.pack(fill="both", expand=1, padx=2, side=tk.BOTTOM)
signup_crl1 = tk.Button(signupinfo_box0, text='1、点我连接数据库', font=g_font, command=connDB)
signup_crl1.pack(fill='both', expand=1, padx=2, pady=2)

signupinfo_box0.pack_forget()



grainfo_box0 = tk.LabelFrame(frm)   # 功能界面底层容器-毕业生信息查询
grainfo_box0.pack(fill="both", expand=1, padx=2, side=tk.BOTTOM)
grainfo_crl1 = tk.Button(grainfo_box0, text='点我连接数据库', font=g_font, command=grainfo_connDB)
grainfo_crl1.pack(fill='both', expand=1, padx=2, pady=2)

grainfo_box1 = tk.LabelFrame(grainfo_box0, text='生成毕业生信息', font=g_font)
grainfo_box1.pack(fill="both", expand=1, padx=2, pady=15)

grainfo_rules =tk.Label(grainfo_box1, text='生成规则：', font=g_font)
number = tk.StringVar()
grainfo_rulesc = ttk.Combobox(grainfo_box1, textvariable=number)
grainfo_rulesc['values'] = ('学号正序', '学号倒叙','报到号正序', '报到号倒叙', '身份证号正序', '身份证号倒叙')
grainfo_year = tk.Label(grainfo_box1, text=' 毕业年份：', font=g_font)
grainfo_yeart = tk.Entry(grainfo_box1, bd=2)  # bd为边框大小，缺省值为1
grainfo_bu = tk.Button(grainfo_box1, text='生成', font=g_font, command=create_grainfo)
grainfo_rules.pack(expand=1, padx=2, pady=2, side=tk.LEFT)
grainfo_rulesc.pack(expand=1, padx=2, pady=2, side=tk.LEFT)
grainfo_year.pack(expand=1, padx=2, pady=2, side=tk.LEFT)
grainfo_yeart.pack(expand=1, padx=2, pady=2, side=tk.LEFT)
grainfo_bu.pack(expand=1, padx=10, pady=2, side=tk.LEFT)
grainfo_bu.configure(state='disable')

grainfo_box2 = tk.LabelFrame(grainfo_box0, text='毕业生信息查询', font=g_font)
grainfo_box2.pack(fill="both", expand=1, padx=2, pady=15)
grainfo_terms = tk.Label(grainfo_box2, text='查询条件：', font=g_font)
number = tk.StringVar()
grainfo_termsc = ttk.Combobox(grainfo_box2, textvariable=number)
grainfo_termsc['values'] = ('学号', '姓名', '身份证号')
var = tk.StringVar()
grainfo_in = tk.Label(grainfo_box2, text=' 输入内容：', font=g_font)
grainfo_termse = tk.Entry(grainfo_box2, bd=2, textvariable=var)
grainfo_termse.bind('<Return>', gra_search)
grainfo_but = tk.Button(grainfo_box2, text='查询', font=g_font, command=gra_search)
grainfo_terms.pack(expand=1, padx=2, pady=2, side=tk.LEFT)
grainfo_termsc.pack(expand=1, padx=2, pady=2, side=tk.LEFT)
grainfo_in.pack(expand=1, padx=2, pady=2, side=tk.LEFT)
grainfo_termse.pack(expand=1, padx=2, pady=2, side=tk.LEFT)
grainfo_but.pack(expand=1, padx=10, pady=2, side=tk.LEFT)
grainfo_but.configure(state='disable')

grainfo_box3 = tk.LabelFrame(grainfo_box0, text='查询结果', font=g_font)

# 创建Treeview来展示查询到的数据
resultForms = ttk.Treeview(grainfo_box3, height=5, show="headings",
                           columns=('xh', 'xm', 'xzb', 'rxrq', 'sfzh', 'byzh', 'xwzh'))  # height控件高度，show=隐藏Treeview中的首列 column设置字段
resultForms.pack(side=tk.LEFT)
resultForms.heading('xh', text='学号')  # 设置字段的显示名称
resultForms.heading('xm', text='姓名')
resultForms.heading('xzb', text='班级')
resultForms.heading('rxrq', text='入学日期')
resultForms.heading('sfzh', text='身份证号')
resultForms.heading('byzh', text='毕业证号')
resultForms.heading('xwzh', text='学位证号')
resultForms.column('xh', width=80)  # 设置字段宽度
resultForms.column('xm', width=60)
resultForms.column('xzb', width=80)
resultForms.column('rxrq', width=60)
resultForms.column('sfzh', width=120)
resultForms.column('byzh', width=120)
resultForms.column('xwzh', width=120)
vbar = ttk.Scrollbar(grainfo_box3, orient=tk.VERTICAL, command=resultForms.yview)  # 设置滚动条控件
vbar.pack(side=tk.LEFT, fill='y')  # 滚动条控件与被控制的Treeview同在一个容器中，并列放置，纵向填充
resultForms.configure(yscrollcommand=vbar.set)  # 此处被控控件的 yscrollcommand=vbar.set 与上面滚动条控件的 command=resultForms.yview 为相互作用设置
grainfo_box0.pack_forget()






root.mainloop()
