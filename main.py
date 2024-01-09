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


def fileTrans(file_path, appKey, fileLink, word_output):
    # 地域ID，固定值。
    REGION_ID = "cn-shanghai"
    PRODUCT = "nls-filetrans"
    DOMAIN = "filetrans.cn-shanghai.aliyuncs.com"
    API_VERSION = "2018-08-17"
    # 状态值
    with open(file_path, 'r') as f:
        file_content = csv.reader(f)
        data = list(file_content)
        accesskeyid = data[1][0]
        accesskeysecret = data[1][1]
    # 创建AcsClient实例
    client = AcsClient(accesskeyid, accesskeysecret, REGION_ID)
    # 提交录音文件识别请求
    postRequest = CommonRequest()
    postRequest.set_domain(DOMAIN)
    postRequest.set_version(API_VERSION)
    postRequest.set_product(PRODUCT)
    postRequest.set_action_name("SubmitTask")
    postRequest.set_method('POST')
    # 新接入请使用4.0版本，已接入（默认2.0）如需维持现状，请注释掉该参数设置。
    # 设置是否输出词信息，默认为false，开启时需要设置version为4.0。
    task = {"appkey": appKey, "file_link": fileLink, "version": "4.0", "enable_words": True, "enable_sample_rate_adaptive": True}
    # 开启智能分轨，如果开启智能分轨，task中设置KEY_AUTO_SPLIT为True。
    task = json.dumps(task)
    postRequest.add_body_params("Task", task)
    taskId = ""
    try:
        postResponse = client.do_action_with_exception(postRequest)
        postResponse = json.loads(postResponse)
        statusText = postResponse["StatusText"]
        if statusText == "SUCCESS":
            taskId = postResponse["TaskId"]
        else:
            print("录音文件识别请求失败！")
            return
    except ServerException as e:
        print(e)
    except ClientException as e:
        print(e)
    # 创建CommonRequest，设置任务ID。
    getRequest = CommonRequest()
    getRequest.set_domain(DOMAIN)
    getRequest.set_version(API_VERSION)
    getRequest.set_product(PRODUCT)
    getRequest.set_action_name("GetTaskResult")
    getRequest.set_method('GET')
    getRequest.add_query_param("TaskId", taskId)
    # 提交录音文件识别结果查询请求
    # 以轮询的方式进行识别结果的查询，直到服务端返回的状态描述符为"SUCCESS"、"SUCCESS_WITH_NO_VALID_FRAGMENT"，
    # 或者为错误描述，则结束轮询。
    statusText = ""
    while True:
        try:
            getResponse = client.do_action_with_exception(getRequest)
            getResponse = json.loads(getResponse)
            statusText = getResponse["StatusText"]
            if statusText == "RUNNING" or statusText == "QUEUEING":
                # 继续轮询
                time.sleep(10)
            else:
                # 退出轮询
                break
        except ServerException as e:
            print(e)
        except ClientException as e:
            print(e)
    if statusText == "SUCCESS":
        words_data = getResponse['Result']['Words']
        df = pandas.DataFrame(words_data)
        df.to_excel(word_output, index=False)
    else:
        print("录音文件识别失败！")
    return


def speechrecognize():
    filepath = accesskey_path.get()
    appkey = appkey_entry.get()
    filelink = filelink_entry.get()
    word_output = word_output_path.get()
    fileTrans(file_path=filepath, appKey=appkey, fileLink=filelink, word_output=word_output)


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
