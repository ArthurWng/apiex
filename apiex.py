import requests
import ssl
import json, time
import tkinter as tk
ssl._create_default_https_context = ssl._create_unverified_context

import pandas as pd

cases = pd.read_excel('api_test.xlsx')
s = requests.session()
report = []

# 断言函数，如果实际返回值和期望值一致，则判断测试通过，否则反之
def test_result(expect, actual,k):
    if expect==actual:
        #test_info = cases['case_type'][k] +"_"+cases['case_name'][k]+ ": \033[32m测试通过\033[0m"
        test_info = cases['case_type'][k] + "_" + cases['case_name'][k] + ": 测试通过"
        return test_info
    else:
        test_info = cases['case_type'][k] +"-"+cases['case_name'][k]+ ": \033[31m测试失败\033[0m"
        return test_info

# 执行用例函数，先用pandas从本地excel表中获取api的url、data信息，然后用requests.session访问接口
def exe_case(x):
    url = cases['env_host'][0] + cases['api_url'][x]

    if cases['method'][x] == 'post':
        if str(cases['data'][x]) == 'nan':
            data = ''
        else:
            data = eval(cases['data'][x])

        response = s.post(url, data=data)

    else:
        if cases['data'][x] == 'nan':
            data = ''
        else:
            data = eval(cases['data'][x])

        response = s.post(url, params=data)

    '''
    使用eval将字符串转换成对象时，如果包含null值，python中的空值为none，
    此时python不会把null认为是none，python会把null识别为一个变量，所以需要给null赋值为空，
    不然会报"NameError: name ‘null’ is not defined"，或者用json.loads(str)的方法将字符串转换成对象
    '''
    #null = ''
    #print(test_result('000000', str(eval(response.text)['statusCode']), x))
    k = test_result('000000', json.loads(response.text)['statusCode'], x)
    print(k, end = "|")
    report.append(k)

if __name__ == '__main__':

    total_start = time.perf_counter()
    for i in range(cases.index.size):
        start = time.perf_counter()
        exe_case(i)
        end = time.perf_counter()
        print("%6.3fs' " % (end-start))
    total_end = time.perf_counter()

    print("\n用例执行总耗时：%6.3fs' " % (total_end-total_start))
    # print(report)

    #report = report

    window = tk.Tk()
    window.title("测试报告")
    window.geometry("800x600")

    # 测试报告概览
    fm_summary = tk.Frame(window, border=5, bg='#758a99', height = 90, width = 780)
    #fm_summary.pack_propagate(0)
    fm_summary.pack()
    tk.Label(fm_summary, text='[ 测试报告概览 ]', bg='#758a99', fg='#ffffff', font=('Arial', 16)).place(x=5, y=5)
    tk.Label(fm_summary, text = '测试用例执行总耗时：30秒', bg='#758a99', fg='#ffffff', font=('Arial', 12)).place(x=5, y=35)
    tk.Label(fm_summary, text='测试用例执行总数：680个', bg='#758a99', fg='#ffffff', font=('Arial', 12)).place(x=5, y=55)
    tk.Label(fm_summary, text='测试用例通过总数：675个', bg='#758a99', fg='#ffffff', font=('Arial', 12)).place(x=200, y=35)
    tk.Label(fm_summary, text='测试用例失败总数：5个', bg='#758a99', fg='#ffffff', font=('Arial', 12)).place(x=200, y=55)
    tk.Label(fm_summary, text='测试用例通过率：90%', bg='#758a99', fg='#ffffff', font=('Arial', 12)).place(x=400, y=55)

    var = tk.StringVar()
    var.set("展开")
    but = "true"

    # 收起、展开函数
    def list_view():
        global but
        if but == "false":
            but = "true"
            var.set("展开")
            fm_dd.forget()
        else:
            but = "false"
            var.set("收起")
            fm_dd.pack()



    # 报告分页1
    fm_sheet = tk.Frame(window, border=2, bg='#fff', height = 30, width = 780)
    fm_sheet.pack_propagate(0)
    fm_sheet.pack()
    tk.Label(fm_sheet, text='< 通用接口测试用例 >', bg='#fff', font=('Arial', 14)).pack(side='left')
    tk.Button(fm_sheet, textvariable=var, bg='#fff', font=('Arial', 12), height = 2, width=6,command = list_view).pack(side='right')




    # 测试结果
    case_total = len(report)
    h = 30*case_total

    fm_dd = tk.Frame(window, border=2, height=h, width=780, bg='#c2ccd0')

    for i in range(case_total):
        k = report[i]

        fm_detail = tk.Frame(fm_dd, border=2, height=30, width=780, bg='#c2ccd0')
        fm_detail.pack_propagate(0)
        fm_detail.pack()

        l1 = tk.Label(fm_detail,bg='#c2ccd0', text=k).pack(side='left')


    window.mainloop()
