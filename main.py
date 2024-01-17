# -*- coding: utf8 -*-
import json
import time

import pandas
from aliyunsdkcore.acs_exception.exceptions import ClientException
from aliyunsdkcore.acs_exception.exceptions import ServerException
from aliyunsdkcore.client import AcsClient
from aliyunsdkcore.request import CommonRequest
import csv

import tkinter as tk
from tkinter import filedialog


def filetrans(file_path, appkey, filelink, word_output):
    # 地域ID，固定值。
    region_id = "cn-shanghai"
    product = "nls-filetrans"
    domain = "filetrans.cn-shanghai.aliyuncs.com"
    api_version = "2018-08-17"
    # 状态值
    with open(file_path, 'r') as f:
        file_content = csv.reader(f)
        data = list(file_content)
        accesskeyid = data[1][0]
        accesskeysecret = data[1][1]
    # 创建AcsClient实例
    client = AcsClient(ak=accesskeyid, secret=accesskeysecret, region_id=region_id)
    # 提交录音文件识别请求
    postrequst = CommonRequest()
    postrequst.set_domain(domain=domain)
    postrequst.set_version(version=api_version)
    postrequst.set_product(product=product)
    postrequst.set_action_name(action_name="SubmitTask")
    postrequst.set_method(method='POST')
    # 新接入请使用4.0版本，已接入（默认2.0）如需维持现状，请注释掉该参数设置。
    # 设置是否输出词信息，默认为false，开启时需要设置version为4.0。
    task = {"appkey": appkey, "file_link": filelink, "version": "4.0", "enable_words": True, "enable_sample_rate_adaptive": True}
    # 开启智能分轨，如果开启智能分轨，task中设置KEY_AUTO_SPLIT为True。
    task = json.dumps(obj=task)
    postrequst.add_body_params(k="Task", v=task)
    taskid = ""
    try:
        postresponse = client.do_action_with_exception(acs_request=postrequst)
        postresponse = json.loads(postresponse)
        statustext = postresponse["StatusText"]
        if statustext == "SUCCESS":
            taskid = postresponse["TaskId"]
        else:
            print("录音文件识别请求失败！")
            return
    except ServerException as e:
        print(e)
    except ClientException as e:
        print(e)
    # 创建CommonRequest，设置任务ID。
    getrequest = CommonRequest()
    getrequest.set_domain(domain=domain)
    getrequest.set_version(version=api_version)
    getrequest.set_product(product=product)
    getrequest.set_action_name(action_name="GetTaskResult")
    getrequest.set_method(method='GET')
    getrequest.add_query_param(k="TaskId", v=taskid)
    # 提交录音文件识别结果查询请求
    # 以轮询的方式进行识别结果的查询，直到服务端返回的状态描述符为"SUCCESS"、"SUCCESS_WITH_NO_VALID_FRAGMENT"，
    # 或者为错误描述，则结束轮询。
    statustext = ""
    while True:
        try:
            getresponse = client.do_action_with_exception(acs_request=getrequest)
            getresponse = json.loads(getresponse)
            statustext = getresponse["StatusText"]
            if statustext == "RUNNING" or statustext == "QUEUEING":
                # 继续轮询
                time.sleep(10)
            else:
                # 退出轮询
                break
        except ServerException as e:
            print(e)
        except ClientException as e:
            print(e)
    if statustext == "SUCCESS":
        words_data = getresponse['Result']['Words']
        df = pandas.DataFrame(words_data)
        df.to_excel(excel_writer=word_output, sheet_name="识别结果", index=True)
    else:
        print("录音文件识别失败！")
    return


def speechrecognize():
    filepath = accesskey_path.get()
    appkey = appkey_entry.get()
    filelink = filelink_entry.get()
    word_output = word_output_path.get()
    filetrans(file_path=filepath, appkey=appkey, filelink=filelink, word_output=word_output)


def accesskey_path_selector():
    filepath = filedialog.askopenfilename()
    if filepath:
        accesskey_path.delete(0, tk.END)
        accesskey_path.insert(0, filepath)


def word_output_path_selector():
    output_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",  # 默认文件扩展名
        filetypes=[("Excel 结果", "*.xlsx"), ("所有文件类型", "*.*")],
    )
    if output_path:
        word_output_path.delete(0, tk.END)
        word_output_path.insert(0, output_path)


window = tk.Tk()
window.title('鬼畜语音识别')
window.geometry('500x300')

accesskey_path = tk.Entry(window)
accesskey_path.insert(0, "请选择Accesskey路径")
accesskey_path.place(x=110, y=20,  width=210, height=30)

accesskey_path_btn = tk.Button(window, text="选择路径", command=accesskey_path_selector)
accesskey_path_btn.place(x=260, y=20,  width=80, height=30)

appkey_entry = tk.Entry(window)
appkey_entry.insert(0, "请输入AppKey")
appkey_entry.place(x=110, y=60,  width=210, height=30)

filelink_entry = tk.Entry(window)
filelink_entry.insert(0, "请输入文件链接")
filelink_entry.place(x=110, y=100,  width=210, height=30)

word_output_path = tk.Entry(window)
word_output_path.insert(0, "请选择词信息输出路径")
word_output_path.place(x=110, y=140,  width=210, height=30)

word_output_path_btn = tk.Button(window, text="选择路径", command=word_output_path_selector)
word_output_path_btn.place(x=260, y=140,  width=80, height=30)

btn = tk.Button(window, text='开始', command=speechrecognize)
btn.place(x=110, y=180,  width=210, height=30)

window.mainloop()
