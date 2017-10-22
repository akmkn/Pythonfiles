# -*- coding: utf-8 -*-
from __future__ import unicode_literals
from django.shortcuts import render
import os,ctypes,time,datetime
import xlwt,xlrd,subprocess
from django.conf import settings
from django.http import HttpResponse
from django.http import request
from django.http import HttpRequest
import win32com,win32com.client
from xlutils.copy import copy
from win32com.client import Dispatch
# Create your views here.

global WorkBookName,AutomationPortal_QCWorkflow,EmailTriggerScriptPath
WorkBookName = os.path.join(settings.STATIC_ROOT, 'data/RMData.xls')

AutomationPortal_QCWorkflow = os.path.join(settings.STATIC_ROOT, 'vbs/AutomationPortal_QCWorkflow.vbs')
EmailTriggerScriptWorkflow = os.path.join(settings.STATIC_ROOT, 'vbs/AP_AutoEmailTrigger.vbs')

#def EmailTrigger():
    #subprocess.Popen(["wscript.exe", EmailTriggerScriptWorkflow], stdout=subprocess.PIPE)
    #os.system("wscript.exe C:\\Users\\automationqateam\\Project\\venv2\\static_cdn\\vbs\\AP_AutoEmailTrigger.vbs")

def index(request):
    global RMenvironment1, RMscriptname1, RMmachine1, RMexecute1
    global RMenvironment2, RMscriptname2, RMmachine2, RMexecute2
    global RMenvironment3, RMscriptname3, RMmachine3, RMexecute3
    global RMenvironment4, RMscriptname4, RMmachine4, RMexecute4
    global RMSanity,Email,clientipaddress,datetimestamp

    if (request.method =='POST') and (request.POST.get('Execute')=="Execute"):

        #################################obtaining client ip address############################################
        for i in request.META:
            if i == "REMOTE_ADDR":
                clientipaddress = request.META[i]
                break

        #################################obtaining datetimestamp############################################
        timetostr = str(datetime.datetime.now())
        timewithnomillisec = timetostr.split(".")
        datetimestamp = timewithnomillisec[0].replace(":", "")

        ###################################storing the data obtained#########################################
        RMSanity = request.POST.get('RMSanity')
        Email = request.POST.get('email')
       #ctypes.windll.user32.MessageBoxW(0, RMSanity, "Your Title", 1)
        if RMSanity == "RMSanity":
            #ctypes.windll.user32.MessageBoxW(0, Email, "Your Title", 1)
            RMenvironment = request.POST.get('RMenvironment')
           #RMenvironment1 = request.POST.get('RMenvironment1')
            RMenvironment1 = RMenvironment
            RMscriptname1 = request.POST.get('RMscriptname1')
            RMmachine1 = request.POST.get('RMmachine1')
            RMexecute1 = request.POST.get('RMexecute1')
            if RMexecute1 == "on":
                RMexecute1 = "Yes"

            RMenvironment2 = RMenvironment
            RMscriptname2 = request.POST.get('RMscriptname2')
            RMmachine2 = request.POST.get('RMmachine2')
            RMexecute2 = request.POST.get('RMexecute2')
            if RMexecute2 == "on":
                RMexecute2 = "Yes"

            RMenvironment3 = RMenvironment
            RMscriptname3 = request.POST.get('RMscriptname3')
            RMmachine3 = request.POST.get('RMmachine3')
            RMexecute3 = request.POST.get('RMexecute3')
            if RMexecute3 == "on":
                RMexecute3 = "Yes"

            RMenvironment4 = RMenvironment
            RMscriptname4 = request.POST.get('RMscriptname4')
            RMmachine4 = request.POST.get('RMmachine4')
            RMexecute4 = request.POST.get('RMexecute4')
            if RMexecute4 == "on":
               RMexecute4 = "Yes"

            ##########################################################################################
            #creation of excel sheet with data and saving with the concatenation of clientipaddress and datetime
            writeRMData(clientipaddress,datetimestamp)     #----------------write to newly created workbook----------------#
            ##########################################################################################

               #vbscriptcall()    #----------------vbscript to connect QC and invoke script ---------------#
            contexttopost = {'RMenvironment':RMenvironment}
            contextRMData = readRMData()
            context = {}
            for i in [contexttopost,contextRMData]:
                context.update(i)

            return render(request, 'automationUI/indexafterupdate.html', context)
               #EmailTrigger()
    else:
        return render(request, 'automationUI/index.html')

def readRMData():
    wkbook = xlrd.open_workbook(WorkBookName,'r')
    sheetname = wkbook.sheet_by_index(0)
    RM1statusval = sheetname.cell_value(1, 8)
    RM2statusval = sheetname.cell_value(2, 8)
    RM3statusval = sheetname.cell_value(3, 8)
    RM4statusval = sheetname.cell_value(4, 8)
    context = {'RM1statusval': RM1statusval,'RM2statusval':RM2statusval,'RM3statusval':RM3statusval,'RM4statusval':RM4statusval}
    return context

def vbscriptcall():
    wkbook = xlrd.open_workbook(WorkBookName)
    sheetname = wkbook.sheet_by_index(0)
    #colcount = sheetname.ncols
    rowcount = int(sheetname.nrows)
    #for i in range(rowcount-1):
    for i in range(4):
             subprocess.Popen(["wscript.exe",AutomationPortal_QCWorkflow],stdout=subprocess.PIPE)
             if i==1:
                 time.sleep(90)
             else:
                 time.sleep(10)
             #ctypes.windll.user32.MessageBoxW(0,"Middle","Your Title",1)

def writeRMData(clientipaddress,datetimestamp):
    wb = xlwt.Workbook()
    ws = wb.add_sheet('data')

    TestPlanFolderPath = "[QualityCenter]Subject\Automation\ECO\Products\AutomationPortal"
    TestSetFolderPath = "Root\ECO Automation\Product-Execution\AutomationPortal"
    SanityName = "ReleaseManagement"
    rt_IsRunning = "No"
    rt_IsStop = "No"
    Email = "ashiskumar@ap.equinix.com"
    ExecutorId = "All"

    if RMSanity == "RMSanity":
        ws.write(1, 0, RMexecute1)
        ws.write(2, 0, RMexecute2)
        ws.write(3, 0, RMexecute3)
        ws.write(4, 0, RMexecute4)

        ws.write(1, 1, TestPlanFolderPath)
        ws.write(2, 1, TestPlanFolderPath)
        ws.write(3, 1, TestPlanFolderPath)
        ws.write(4, 1, TestPlanFolderPath)

        ws.write(1, 2, RMscriptname1)
        ws.write(2, 2, RMscriptname2)
        ws.write(3, 2, RMscriptname3)
        ws.write(4, 2, RMscriptname4)

        ws.write(1, 3, TestSetFolderPath)
        ws.write(2, 3, TestSetFolderPath)
        ws.write(3, 3, TestSetFolderPath)
        ws.write(4, 3, TestSetFolderPath)

        ws.write(1, 4, SanityName)
        ws.write(2, 4, SanityName)
        ws.write(3, 4, SanityName)
        ws.write(4, 4, SanityName)

        ws.write(1, 5, RMmachine1)
        ws.write(2, 5, RMmachine2)
        ws.write(3, 5, RMmachine3)
        ws.write(4, 5, RMmachine4)

        ws.write(1, 6, rt_IsRunning)
        ws.write(2, 6, rt_IsRunning)
        ws.write(3, 6, rt_IsRunning)
        ws.write(4, 6, rt_IsRunning)

        ws.write(1, 7, rt_IsStop)
        ws.write(2, 7, rt_IsStop)
        ws.write(3, 7, rt_IsStop)
        ws.write(4, 7, rt_IsStop)

        ws.write(1, 10, RMenvironment1)
        ws.write(2, 10, RMenvironment2)
        ws.write(3, 10, RMenvironment3)
        ws.write(4, 10, RMenvironment4)

        ws.write(0, 11, "MailTo")
        ws.write(1, 11, Email)
        ws.write(0, 12, "ExecutorId")
        ws.write(1, 12, ExecutorId)

        #RMwbname = 'data/RMData_'+clientipaddress+'_'+datetimestamp+'.xls'
        #wb.save(os.path.join(settings.STATIC_ROOT,RMwbname))

        wb.save(os.path.join(settings.STATIC_ROOT, 'data/RMData.xls'))