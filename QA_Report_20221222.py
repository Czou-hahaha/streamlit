# -*- coding: utf-8 -*-
"""
@author: Corrine
"""

import streamlit as st
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px
import xlrd
import datetime
from datetime import date
from st_aggrid import AgGrid
import altair as alt
from PIL import Image
from openpyxl import load_workbook
# use the following function to use saveattachments
# from zydsaveattachments import saveattachments
import os, zipfile,shutil
import win32com.client as win32
import csv
from streamlit_autorefresh import st_autorefresh
import xlwings as xw

# using to run the saveattachments
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import smtplib
import win32com.client as win32
import pythoncom
import zipfile
import glob,shutil
# from streamlit.ScriptRunner import RerunException
# cope with the problem'CoInitialize has not been called', but why need to use nruns = 0?
import pythoncom
pythoncom.CoInitialize()

# The reason why use nruns: in this website, we actually run 2 separate process, if we set nruns = 0, we force the two process run together.
nruns = 0


# define function
# 保存邮箱里面的文件，在一个文件夹里面，并把所有的pd_change和risk factor文件更改名字之后提取出来。
# 首先
# 这里面需要改变的var就是date。需要写一个for loop一个一个传入
def saveattachments(subject, QA_filename, path, output_folder, copy_folder, date, type = 'zip'):
	# 判断path是否为目录，如果不是创建path
	if not os.path.isdir(path):
		os.mkdir(path)
	# 创建一个对象通过MAPI协议连接outlook
	outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
	# 获取outlook中的收件箱位置
	inbox = outlook.GetDefaultFolder(6)
	# 获取收件箱文件夹中的邮件对象
	messages = inbox.Items
	# 设置附件变量，循环读取邮件中的附件
	attachment = ''
	# 对于在收件箱中的邮件
	for message in messages:
		# 如果邮件名称为subject
		if message.Subject == subject:
			# body_content = message.body
			# 通过attachments读取附件
			attachments = message.Attachments
			# 通过attachment读取附件中的内容
			attachment = attachments.Item(1)
			# 对于每一个在邮件中的附件，保存在path中名字为附件名字。
			for attachment in message.Attachments:
				# os.path.join 连接两个或更多的路径名组件，如果组件中没有 /， 这个函数会自动加上。
				attachment.SaveAsFile(os.path.join(path, str(attachment)))
				break
	# 如果附件为空，显示没有检测到文件。
	if len(str(attachment))==0:
		print("didn't detect the file")
	# os.path.expanduser主要功能把 ~ 转化为user目录
	directory_to_extract_to = os.path.expanduser(f"{path}/{output_folder}")
	# 把转化之后的目录打印出来
	print(directory_to_extract_to)
	# 如果目录已经存在，则递归删除所有子文件夹和子文件，否则新建一个目录
	if os.path.isdir(directory_to_extract_to):
		shutil.rmtree(directory_to_extract_to)
	else:
		os.mkdir(directory_to_extract_to)
	# 如果type是压缩包，打印包含QA_filename的路径，mode = 'r' 指以读取模式打开文件，然后传入zip_ref变量中，并把它解压在上面的目录中。
	#https://zhuanlan.zhihu.com/p/480329034#:~:text=Python%20%E7%9A%84%20zipfile%20%E6%98%AF%E4%B8%80%E4%B8%AA,%E5%A4%A7%E5%B0%8F%E5%B9%B6%E8%8A%82%E7%9C%81%E7%A3%81%E7%9B%98%E7%A9%BA%E9%97%B4%E3%80%82
	if type == 'zip':
		print(path+f'/{QA_filename}')
		with zipfile.ZipFile(path+f'/{QA_filename}', mode = 'r') as zip_ref:
			zip_ref.extractall(directory_to_extract_to)

# 把pd_change和PD_Change_MDF&RF中的risk factor更改名字，然后保存在新的文件夹中。
#创建新的copy路径，并检查是否存在。
	directory_to_copy_to = os.path.expanduser(f"{path}/{copy_folder}")
	print(directory_to_copy_to)
	if os.path.exists(directory_to_copy_to):
		pass
	else:
		os.mkdir(directory_to_copy_to)
    # copy每一天的看板文件，并粘贴到新的copy file
	source_pd = os.path.expanduser(f"{directory_to_extract_to}\PD_Change.csv")
	dest_pd = os.path.expanduser(f"{directory_to_copy_to}\PD_Change_{date}.csv")
	source_rf = os.path.expanduser(f"{directory_to_extract_to}\PD_Change_MDF&RF\Risk_Factors_Change.csv")
	dest_rf = os.path.expanduser(f"{directory_to_copy_to}\Risk_Factors_Change_{date}.csv")
	shutil.copyfile(source_pd,dest_pd)
	shutil.copyfile(source_rf,dest_rf)


# Use Streamlit to form a web
# set wide mode (set_page_config must be called first in the script)
PAGE_CONFIG = {"page_title": "Hello",

               "page_icon": ":smiley:", "layout": "wide"}
st.set_page_config(**PAGE_CONFIG)

# set a new folder path to save attachments
if os.path.exists(os.path.expanduser("~/Desktop/Attachment")):
    pass
else:
    os.mkdir(os.path.expanduser("~/Desktop/Attachment"))

if os.path.exists(os.path.expanduser("~/Desktop/Attachment/0.summary")):
    pass
else:
    os.mkdir(os.path.expanduser("~/Desktop/Attachment/0.summary"))

# 在第一次启动时重置文件，需要把今天的数据下载之后解压，并作为我们的初始数据进行读取
# 如果这个文件已经存在，则读取另外一个合并文件。
# 把今天的数据保存下来
# 如果今天是周一，首先尝试今天是否有数据，否则上周五的数据。如果今天不是周一，首先尝试今天是否有数据，否则用前一天的数据。
# today = date.today()
# yesterday = (today - datetime.timedelta(days=1)).strftime('%Y%m%d')

if os.path.exists(os.path.expanduser("~/Desktop/Attachment/0.start")):
    pass
else:
    today = date.today()
    yesterday = (today - datetime.timedelta(days=1)).strftime('%Y%m%d')
    subject = f'IRAP_QA_{yesterday}'
    QA_filename = f'IRAP_QA_Report_{yesterday}.zip'
    path = os.path.expanduser("~/Desktop/Attachment")
    if today.weekday() == 1:
        try:
            saveattachments(subject, QA_filename,path,output_folder = 'IRAP',copy_folder = '0.start',date = yesterday, type = 'zip')
        except:
            yesterday = (today - datetime.timedelta(days=3)).strftime('%Y%m%d')
            subject = f'IRAP_QA_{yesterday}'
            QA_filename = f'IRAP_QA_Report_{yesterday}.zip'
            path = os.path.expanduser("~/Desktop/Attachment")
            saveattachments(subject, QA_filename,path,output_folder = 'IRAP',copy_folder = '0.start',date = yesterday, type = 'zip')
    else:
        try:
            saveattachments(subject, QA_filename,path,output_folder = 'IRAP',copy_folder = '0.start',date = yesterday, type = 'zip')
        except:
            yesterday = (today - datetime.timedelta(days=2)).strftime('%Y%m%d')
            subject = f'IRAP_QA_{yesterday}'
            QA_filename = f'IRAP_QA_Report_{yesterday}.zip'
            path = os.path.expanduser("~/Desktop/Attachment")
            saveattachments(subject, QA_filename,path,output_folder = 'IRAP',copy_folder = '0.start',date = yesterday, type = 'zip')

# set a summary data to save the conclusion and and summary: 来储存每一次新增的conclusion
# 同时需要update risk_factor
# 下面需要判断是否需要进行初始化，如果第一次使用需要下载最新的QAreport，然后读取数据
# 如果没有今天的数据，直接
if os.path.exists(os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_pd.csv")):
    pd_start = pd.read_csv(os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_pd.csv"))
else:
    today = date.today() 
    yesterday = (today - datetime.timedelta(days=1)).strftime('%Y%m%d')
    a_columns = ['ImportantFlag','CompanyCode','CompanyName','DataDate','Top_MDF','PD','ytdPD','ChangeInNumber','ChangeInPct','PDir_change','REGION_Name','INDUSTRY_LEVEL_1_Name','FS_type','Conclusion','Summary']
    a = pd.DataFrame(columns = a_columns)
    a.to_csv(os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_pd.csv"),index = False)
    if os.path.exists(os.path.expanduser(f"~/Desktop/Attachment/0.start/PD_Change_{yesterday}.csv")):
        read_pd = pd.read_csv(os.path.expanduser(f"~/Desktop/Attachment/0.start/PD_Change_{yesterday}.csv"))
        read_pd.to_csv(os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_pd.csv"),mode = 'a',index = False,header = False)
    else:
        yesterday = (today - datetime.timedelta(days=2)).strftime('%Y%m%d')
        if os.path.exists(os.path.expanduser(f"~/Desktop/Attachment/0.start/PD_Change_{yesterday}.csv")):
            read_pd = pd.read_csv(os.path.expanduser(f"~/Desktop/Attachment/0.start/PD_Change_{yesterday}.csv"))
            read_pd.to_csv(os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_pd.csv"),mode = 'a',index = False,header = False)
        else:
            yesterday = (today - datetime.timedelta(days=3)).strftime('%Y%m%d')
            read_pd = pd.read_csv(os.path.expanduser(f"~/Desktop/Attachment/0.start/PD_Change_{yesterday}.csv"))
            read_pd.to_csv(os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_pd.csv"),mode = 'a',index = False,header = False)
    # 把新下载的内容写入summary_data, 并作为start文件进行读取。
    # excelwrite =  pd.ExcelWriter(os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_pd.csv"), mode = 'a',if_sheet_exists='overlay')
    # read_pd.to_csv(os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_pd.csv"),mode = 'a',index = False,header = False)
    # pd_start = pd.read_csv(os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_pd.csv"))

# 新建rf_summary
# 把新下载的内容写入summary_data, 并作为start文件进行读取。
if os.path.exists(os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_rf.csv")):
    rf_start = pd.read_csv(os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_rf.csv"))
else:
    today = date.today()
    yesterday = (today - datetime.timedelta(days=1)).strftime('%Y%m%d')
    b_columns = ['CompanyCode','CompanyName','DataDate','RFID','RFValue','ytdRFValue','Change Value']
    b = pd.DataFrame(columns = b_columns)
    b.to_csv(os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_rf.csv"),index = False)
    if os.path.exists(os.path.expanduser(f"~/Desktop/Attachment/0.start/Risk_Factors_Change_{yesterday}.csv")):
        read_rf = pd.read_csv(os.path.expanduser(f"~/Desktop/Attachment/0.start/Risk_Factors_Change_{yesterday}.csv"))   
        read_rf.to_csv(os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_rf.csv"),mode = 'a',index = False,header = False)
    else:
        yesterday = (today - datetime.timedelta(days=2)).strftime('%Y%m%d')
        if os.path.exists(os.path.expanduser(f"~/Desktop/Attachment/0.start/Risk_Factors_Change_{yesterday}.csv")):
            read_rf = pd.read_csv(os.path.expanduser(f"~/Desktop/Attachment/0.start/Risk_Factors_Change_{yesterday}.csv"))   
            read_rf.to_csv(os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_rf.csv"),mode = 'a',index = False,header = False)
        else:
            yesterday = (today - datetime.timedelta(days=2)).strftime('%Y%m%d')
            read_rf = pd.read_csv(os.path.expanduser(f"~/Desktop/Attachment/0.start/Risk_Factors_Change_{yesterday}.csv"))   
            read_rf.to_csv(os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_rf.csv"),mode = 'a',index = False,header = False)
    # rf_start = pd.read_csv(os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_rf.csv"))
 

#set header and date
today = date.today() - datetime.timedelta(days=1)
yesterday = today - datetime.timedelta(days=1)
title = str(today) + ' QA_report'
st.header(title)

start_date = st.sidebar.date_input(
    "Start date:",value = yesterday
)
end_date = st.sidebar.date_input(
    "End date:",value = yesterday
)

run = st.sidebar.button(label="Run")

    # if run:
    #first need to check out whether the date is in the sheet?
    #If not in the sheet, we need to download the QA report and combine the date into one sheet (using for loop)
if run:
    # 因为每一天的csv都是前一天的数据，需要把date设置为前一天的数据。
    # read data 文件名字需要更新。因为在合并数据之后需要读取新的file
    csv_path_1 = os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_pd.csv")
    csv_path_3 = os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_rf.csv")

    df_qa = pd.read_csv(csv_path_1)
    df_rf = pd.read_csv(csv_path_3)
    
    # start_period = start_date - datetime.timedelta(days=1)
    # end_period = end_date - datetime.timedelta(days=1)
    start_period = start_date
    end_period = end_date 
    #right now the data type is datetime64  ??? whether it can be used in for loop?
    date_range = pd.date_range(start_period,end_period)
    # check date in the original file
    date_check = df_qa['DataDate'].unique()
    #compare which one of date_range in not in date_check
    date_add = date_range.difference(date_check).strftime('%Y%m%d')
    # don't need to check whether the new datefile in the target folder, because it will download automatically and cover the previous one
    # use date_add create a for loop to download the file and save the using data into a new file
    for date_n in date_add:
        subject = f'IRAP_QA_{date_n}'
        QA_filename = f'IRAP_QA_Report_{date_n}.zip'
        path = os.path.expanduser("~/Desktop/Attachment")
        # use saveattachments download and extract the file
        saveattachments(subject, QA_filename,path,output_folder = 'IRAP',copy_folder = 'datacopy',date = date_n, type = 'zip')
    # 需要把每个file合并在一起，并且传递参数到原来的文件
    # 需要把data_add的数据从大到小排序，并合并到一起。
    # here only append data and doesn't change the previous one.
    #需要合并两个表格，一个是pd_change，另一个是risk_factor_change
    # pd_write = []
    for date_m in date_add:
        pd_file_name = path + f'/datacopy/PD_Change_{date_m}.csv'
        pd_file_add = pd.read_csv(pd_file_name)
        pd_file_add.to_csv(os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_pd.csv"), header = False,index = False, mode = 'a')

    for date_m in date_add:
        rf_file_name = path + f'/datacopy/Risk_Factors_Change_{date_m}.csv'
        rf_file_add = pd.read_csv(rf_file_name)
        rf_file_add.to_csv(os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_rf.csv"), header = False,index = False,mode = 'a')


tab1, tab2 = st.tabs(["Overall","Individual"])

with tab1:

    csv_path_1 = os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_pd.csv")
    df_qa_o = pd.read_csv(csv_path_1)
    company_name = pd.DataFrame(df_qa_o["CompanyName"].drop_duplicates())
    st.header("Overall Report")
    df_show_all = df_qa_o[(df_qa_o["ImportantFlag"] != 0) & (df_qa_o['DataDate'].apply(pd.to_datetime) <= pd.to_datetime(end_date)) & (df_qa_o['DataDate'].apply(pd.to_datetime) >= pd.to_datetime(start_date))].drop(['FS_type'],axis = 1)    
    st.dataframe(df_show_all.sort_values('DataDate',ascending= False)) 
    # AgGrid(df_show_all.sort_values('DataDate',ascending= False))     

with tab2:
    csv_path_1 = os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_pd.csv")
    csv_path_3 = os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_rf.csv")

    df_qa_i = pd.read_csv(csv_path_1)
    df_rf_i = pd.read_csv(csv_path_3)

    df = df_qa_i.reset_index()
    companyCode = df['CompanyName'].drop_duplicates().to_list()
    flagType = df['ImportantFlag'].drop_duplicates().to_list()
    
# set sidebar buttons in a line: using CSS coding 

# !!! how to set the previous and next button: use session_state to update the previous state of new variables.
    
    if 'key' not in st.session_state:
        st.session_state.count = 0
        st.session_state.key = df['CompanyName'].loc[df['index'] ==0].to_list()[0]
        st.session_state.previous_company = df['CompanyName'].loc[df['index'] ==st.session_state.count].to_list()[0]
        st.session_state.next_company = df['CompanyName'].loc[df['index'] ==st.session_state.count].to_list()[0]
        st.session_state.company = df['CompanyName'].loc[df['index'] ==st.session_state.count].to_list()[0]

    next = st.sidebar.button('Next')
    if next:
        st.session_state.count += 1
        if st.session_state.count == len(df):
            st.session_state.count = 0
            st.session_state.next_company = df['CompanyName'].loc[df['index'] ==st.session_state.count].to_list()[0]
        else:
            st.session_state.next_company = df['CompanyName'].loc[df['index'] ==st.session_state.count].to_list()[0]
        st.session_state.company = st.session_state.next_company

    previous = st.sidebar.button('Previous')
    if previous:
        st.session_state.count -= 1
        if st.session_state.count == -1:
            st.session_state.count = len(df) - 1
            st.session_state.previous_company = df['CompanyName'].loc[df['index'] == st.session_state.count].to_list()[0]
        else:
            st.session_state.previous_company = df['CompanyName'].loc[df['index'] ==st.session_state.count].to_list()[0]
        st.session_state.company = st.session_state.previous_company

# set company selectbox canbe changed as the next and previous
    
    if previous | next:
        flag_type = st.sidebar.selectbox("Important Flag:", df['ImportantFlag'][df['CompanyName']==st.session_state.company])
        company = st.sidebar.selectbox("Company", [st.session_state.company])
        
    else:
        flag_type = st.sidebar.selectbox("Important Flag:", flagType)
        company = st.sidebar.selectbox("Company", companyCode)


# show the filtered company
    # if the same company showed in different days: this selectbox will help
    # flag_type = 0
    # company = 'Mobilum Technologies Inc'
    part_df_qa = df_qa_i[ (df_qa_i["ImportantFlag"] == flag_type) & (df_qa_i['CompanyName'] == company) ].sort_values('DataDate',ascending= False)
    part_df_qa = part_df_qa.iloc[:,2:]
    st.dataframe(part_df_qa,width=1000)
    st.write(f"According to your filter, there are {len(part_df_qa)} rows")

    df_rf_new = df_rf_i
    col_names = ['Company_code','Company_name','Data_date','Risk Factor','Today Value','Yest Value','Change Value']
    df_rf_new.columns = col_names
    # reshape df_rf
    df_rf_new['Change Value'] = df_rf_new['Today Value'] - df_rf_new['Yest Value']
    df_rf_new= df_rf_new[['Company_name','Data_date','Risk Factor','Today Value','Yest Value','Change Value']]
    Company_date = st.selectbox("Risk Date:", pd.to_datetime(part_df_qa['DataDate'].to_list()).strftime('%Y-%m-%d'))
    print('Company_date: ', Company_date)
    # company = 'Mobilum Technologies Inc'
    # Company_date = '2022-12-16'
    df_rf_n = df_rf_new[ (df_rf_new['Company_name'] == company) & (df_rf_new['Data_date'] == Company_date)]
    print('df_rf_n: ', df_rf_n)

# show the RF contribution picture and RF table separately
    col1, col2 = st.columns(2)
    # here only need to change the df-rf

    
    # draw a bar chart
    with col1:
        # set if when PD increase or decrease separately
        if df_qa_i[df_qa_i['CompanyName'] == company]['PD'].values[0] > df_qa_i[df_qa_i['CompanyName'] == company]['ytdPD'].values[0]:
            barchart = alt.Chart(df_rf_n).mark_bar().encode(
                x = alt.X("Change Value",scale = alt.Scale(domain = [-1,1])),
                y = alt.Y('Risk Factor',sort = "-x")
                
            ).properties(
                height = 690)
        else: 
            barchart = alt.Chart(df_rf_n).mark_bar().encode(
                x = alt.X("Change Value",scale = alt.Scale(domain = [-1,1])),
                y = alt.Y('Risk Factor',sort = "-x")
            ).properties(
                height = 690)
        
        st.altair_chart(barchart,use_container_width=True)

    with col2:
    # # show the filtered company risk factors and set the start_date is yeasterday by default
        # don't need this apply_color
        def apply_color(value):
            if value < 0:
                color = '#008000'
            elif value > 0:
                color = '#FF0000'
            else: 
                color = '#000000'
            return 'color: %s' % color
        # highlight the first three marginal contribution and change the negative or positive value and sort them
        if df_qa_i[df_qa_i['CompanyName'] == company]['PD'].values[0] > df_qa_i[df_qa_i['CompanyName'] == company]['ytdPD'].values[0]:
            df_rf_n = df_rf_n.sort_values(by = 'Change Value',ascending= False).reset_index(drop = True)
            df_rf_n = df_rf_n.style.highlight_quantile(axis = 0,color = 'yellow',q_left = 0.85,subset = ['Change Value']).set_properties(**{'font-size':'9pt'})
        else: 
            df_rf_n = df_rf_n.sort_values(by = 'Change Value',ascending= True).reset_index(drop = True)
            df_rf_n = df_rf_n.style.highlight_quantile(axis = 0,color = 'yellow',q_right = 0.15,subset = ['Change Value']).set_properties(**{'font-size':'9pt'})
        st.table(df_rf_n)
# set two text input and submit button
    st.write("Reminder: Your non-empty text will be recorded in the QA report. Only English written can be accepted.")

    col3, col4 = st.columns(2)
    write_form = st.form(key = 'write_form')
    with col3:
        conclusion = st.text_area("Conclusion: Please write Reasonable / Suspension in the following part")
    with col4:
        summary = st.text_area("Summary: Please write the firm's summary in the following part")
    # write_form = st.form(key = 'write_form')
    # # st.write("")
    # with col3:
    #     conclusion = write_form.text_input("Conclusion: Please write Reasonable / Suspension in the following part")
    # # st.write("Please write the firm's summary in the following part")
    # with col4:
    #     summary = write_form.text_area("Summary: Please write the firm's summary in the following part")

# transfer data into excel.
    submitted = write_form.form_submit_button(label = 'Submit')
# change the specific value of the working sheet
    if submitted: 
        index_i = df_qa_i[ (df_qa_i['CompanyName']==st.session_state.company) ].index.to_list()[0]
# & (df_qa_i['DataDate'] == Company_date)
        df_qa_i.loc[index_i,['Conclusion']] = conclusion
        df_qa_i.loc[index_i,['Summary']] = summary
        df_qa_i.to_csv(os.path.expanduser("~/Desktop/Attachment/0.summary/0.summary_data_pd.csv"),index = False)
        st.success("The modification has been recorded!")
