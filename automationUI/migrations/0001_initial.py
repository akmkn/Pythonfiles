# -*- coding: utf-8 -*-
# Generated by Django 1.11.1 on 2017-08-14 10:30
from __future__ import unicode_literals

from django.db import migrations, models


class Migration(migrations.Migration):

    initial = True

    dependencies = [
    ]

    operations = [
        migrations.CreateModel(
            name='ATE_ConnecttoQC',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('qcServer', models.CharField(max_length=250)),
                ('qcUser', models.CharField(max_length=250)),
                ('qcPassword', models.CharField(max_length=250)),
                ('qcDomain', models.CharField(max_length=250)),
                ('qcProject', models.CharField(max_length=250)),
            ],
        ),
        migrations.CreateModel(
            name='ATE_PortalProcessedRequests_Records',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('RequestCounter', models.CharField(max_length=250)),
                ('Client_IP_Address', models.CharField(max_length=250)),
                ('TestSetFolderPath', models.CharField(max_length=250)),
                ('ALM_NewTestSetName', models.CharField(max_length=250)),
                ('MailTo', models.CharField(max_length=250)),
                ('sDateTime', models.DateField()),
            ],
        ),
        migrations.CreateModel(
            name='ATE_ProceedToSubmitFlag',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('rowid', models.IntegerField(default=1)),
                ('proceedtosubmitflag', models.CharField(default=True, max_length=5)),
                ('RequestTime', models.CharField(blank=True, default='', max_length=250, null=True)),
                ('ClientRequestIP', models.CharField(blank=True, max_length=250)),
                ('RMenvironment', models.CharField(blank=True, max_length=250)),
                ('Email', models.CharField(blank=True, max_length=250)),
                ('ExecuteCounter', models.IntegerField(blank=True, null=True)),
            ],
        ),
        migrations.CreateModel(
            name='ATE_QC_TestplanFolder_Scripts',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('QCTestPlanPath_FolderName', models.CharField(max_length=250)),
                ('QCTestPlanPath_ScriptNames', models.CharField(blank=True, max_length=250)),
            ],
        ),
        migrations.CreateModel(
            name='ATE_QuoteApprovals',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('Execute', models.CharField(default='Yes', max_length=250)),
                ('TestPlanFolderPath', models.CharField(max_length=250)),
                ('TestScriptNameToExecute', models.CharField(max_length=250)),
                ('TestSetFolderPath', models.CharField(max_length=250)),
                ('SanityName', models.CharField(max_length=250)),
                ('RemoteMachineName', models.CharField(max_length=250)),
                ('rt_IsScriptRunning', models.CharField(blank=True, max_length=250)),
                ('rt_IsMachineAvailable', models.CharField(blank=True, max_length=250)),
                ('RunningStatus', models.CharField(blank=True, default='No Run', max_length=250)),
                ('Trigger', models.CharField(blank=True, max_length=250)),
                ('Environment', models.CharField(max_length=250)),
                ('IsTestSetCreated', models.CharField(blank=True, max_length=250)),
                ('ALM_NewTestSetName', models.CharField(blank=True, max_length=250)),
                ('MailTo', models.CharField(max_length=250)),
                ('ExecutorId', models.CharField(default='All', max_length=4)),
                ('ErrorMessage', models.CharField(blank=True, max_length=250)),
                ('IsEmailSent', models.CharField(blank=True, max_length=250)),
            ],
        ),
        migrations.CreateModel(
            name='ATE_RMData',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('Execute', models.CharField(default='Yes', max_length=250)),
                ('TestPlanFolderPath', models.CharField(max_length=250)),
                ('TestScriptNameToExecute', models.CharField(max_length=250)),
                ('TestSetFolderPath', models.CharField(max_length=250)),
                ('SanityName', models.CharField(max_length=250)),
                ('RemoteMachineName', models.CharField(max_length=250)),
                ('rt_IsScriptRunning', models.CharField(blank=True, max_length=250)),
                ('rt_IsMachineAvailable', models.CharField(blank=True, max_length=250)),
                ('RunningStatus', models.CharField(blank=True, default='No Run', max_length=250)),
                ('Trigger', models.CharField(blank=True, max_length=250)),
                ('Environment', models.CharField(max_length=250)),
                ('IsTestSetCreated', models.CharField(blank=True, max_length=250)),
                ('ALM_NewTestSetName', models.CharField(blank=True, max_length=250)),
                ('MailTo', models.CharField(max_length=250)),
                ('ExecutorId', models.CharField(default='All', max_length=4)),
                ('ErrorMessage', models.CharField(blank=True, max_length=250)),
                ('IsEmailSent', models.CharField(blank=True, max_length=250)),
            ],
        ),
    ]