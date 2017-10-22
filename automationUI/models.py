# -*- coding: utf-8 -*-
from __future__ import unicode_literals
# Create your models here.

from django.db import models

class ATE_ConnecttoQC(models.Model):
    qcServer = models.CharField(max_length=250)
    qcUser = models.CharField(max_length=250)
    qcPassword = models.CharField(max_length=250)
    qcDomain = models.CharField(max_length=250)
    qcProject = models.CharField(max_length=250)


class ATE_RMData(models.Model):
    RowID = models.CharField(max_length=5,default="1")
    Execute = models.CharField(max_length=250,default="Yes")
    TestPlanFolderPath = models.CharField(max_length=250)
    TestScriptNameToExecute = models.CharField(max_length=250)
    TestSetFolderPath = models.CharField(max_length=250)
    SanityName = models.CharField(max_length=250)
    RemoteMachineName = models.CharField(max_length=250)
    rt_IsScriptRunning = models.CharField(max_length=250, blank= True)
    rt_IsMachineAvailable = models.CharField(max_length=250, blank= True)
    RunningStatus = models.CharField(max_length=250, blank= True,default="No Run")
    Trigger = models.CharField(max_length=250, blank= True)
    Environment = models.CharField(max_length=250)
    IsTestSetCreated = models.CharField(max_length=250, blank= True)
    ALM_NewTestSetName = models.CharField(max_length=250, blank= True)
    Email = models.CharField(max_length=250)
    ExecutorId = models.CharField(max_length=4,default="All")
    ErrorMessage = models.CharField(max_length=250, blank= True)
    IsEmailSent = models.CharField(max_length=250, blank= True)

    def __str__(self):
        return self.TestScriptNameToExecute + self.RunningStatus + self.Trigger + self.RemoteMachineName

class ATE_QuoteApprovals(models.Model):
    RowID = models.CharField(max_length=5, default="1")
    Execute = models.CharField(max_length=250,default="Yes")
    TestPlanFolderPath = models.CharField(max_length=250)
    TestScriptNameToExecute = models.CharField(max_length=250)
    TestSetFolderPath = models.CharField(max_length=250)
    SanityName = models.CharField(max_length=250)
    RemoteMachineName = models.CharField(max_length=250)
    rt_IsScriptRunning = models.CharField(max_length=250, blank= True)
    rt_IsMachineAvailable = models.CharField(max_length=250, blank= True)
    RunningStatus = models.CharField(max_length=250, blank= True,default="No Run")
    Trigger = models.CharField(max_length=250, blank= True)
    Environment = models.CharField(max_length=250)
    IsTestSetCreated = models.CharField(max_length=250, blank= True)
    ALM_NewTestSetName = models.CharField(max_length=250, blank= True)
    Email = models.CharField(max_length=250)
    ExecutorId = models.CharField(max_length=4,default="All")
    ErrorMessage = models.CharField(max_length=250, blank= True)
    IsEmailSent = models.CharField(max_length=250, blank= True)


class ATE_ProceedToSubmitFlag(models.Model):
    rowid = models.IntegerField(default=1)
    proceedtosubmitflag = models.CharField(max_length=5,default=True)
    RequestTime = models.CharField(max_length=250,blank=True)
    ClientRequestIP = models.CharField(max_length=250,blank=True)
    RMenvironment = models.CharField(max_length=250,blank=True)
    Email = models.CharField(max_length=250,blank=True)
    ExecuteCounter = models.IntegerField(blank=True,default=0)
    def __str__(self):
        return str(self.rowid) +self.proceedtosubmitflag + self.RequestTime+ self.ClientRequestIP+ self.RMenvironment + self.Email + str(self.ExecuteCounter)

class ATE_PortalProcessedRequests_Records(models.Model):
    RequestCounter = models.CharField(max_length=250)
    Client_IP_Address = models.CharField(max_length=250)
    TestSetFolderPath = models.CharField(max_length=250)
    ALM_NewTestSetName = models.CharField(max_length=250)
    MailTo = models.CharField(max_length=250)
    sDateTime = models.DateField()

class ATE_QC_TestplanFolder_Scripts(models.Model):
    QCTestPlanPath_FolderName = models.CharField(max_length=250)
    QCTestPlanPath_ScriptNames = models.CharField(max_length=250, blank=True)