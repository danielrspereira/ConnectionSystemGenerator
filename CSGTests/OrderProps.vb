Imports System.ComponentModel
Imports System.Globalization
Imports System.IO
Imports CSGCore

Public Class OrderProps
    Public authusers As New List(Of String)

    Private Sub RBAPO_CheckedChanged(sender As Object, e As EventArgs) Handles RBAPO.CheckedChanged
        RBEU.Checked = Not RBAPO.Checked

        If RBAPO.Checked Then
            General.currentjob.Plant = "Beji"
            General.domainname = "asia"
            General.SetInfosDir()
        End If
    End Sub

    Private Sub RBEU_CheckedChanged(sender As Object, e As EventArgs) Handles RBEU.CheckedChanged
        RBAPO.Checked = Not RBEU.Checked

        If RBEU.Checked Then
            General.currentjob.Plant = Order.GetEUPlant
            General.domainname = "europe"
            General.SetInfosDir()
        End If
    End Sub

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        General.decsym = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator
        General.username = SystemInformation.UserName
        General.domainname = SystemInformation.UserDomainName
        General.userdomain = SystemInformation.UserDomainName
        If General.domainname.ToLower = "asia" Then
            RBAPO.Checked = True
        Else
            RBEU.Checked = True
        End If
        General.InitDefaultSettings(Application.StartupPath)
        'check for dll
        If CheckCoreDLL(General.username) Then
            Close()
        End If
        InitAuthUsers()

        If General.username = "mlewin" Or General.username = "csgen" Or General.username = "admin-ml" Then
            CheckTest.Visible = True
            CBCSGMode.Items.Add("Continue saved Order")
            CBCSGMode.Items.Add("ETO Batch")
        ElseIf authusers.IndexOf(General.username) > -1 Then
        Else
            CBUnit.Items.Remove("GADC - Family")
            CBUnit.Items.Remove("Evaporator (2 Coils)")
        End If
    End Sub

    Private Sub BGO_Click(sender As Object, e As EventArgs) Handles BGo.Click

        Try

            Do
                General.seapp = SEUtils.ReConnect()
            Loop Until General.seapp IsNot Nothing

            If TBOrder.Text = "" Or TBPos.Text = "" Or TBNr.Text = "" Or CBCSGMode.Text = "" Then
                MsgBox("Missing entry for order details.")
            Else
                If Controllength("TBOrder", TBOrder.Text) AndAlso Controllength("TBPos", TBPos.Text) AndAlso Controllength("TBNr", TBNr.Text) Then
                    With General.currentjob
                        .OrderNumber = TBOrder.Text
                        .OrderPosition = TBPos.Text
                        .ProjectNumber = TBNr.Text
                    End With
                    Order.CreateOrderDir(General.currentjob)
                    If CBCSGMode.Text.Contains("Continue") Then
                        If CBCSGMode.Text.Contains("local") Then
                            General.OverrideData(ImportJsonFile())

                            If General.currentjob.Uid > 0 And General.currentjob.Status < 100 And General.currentjob.Status > -1 Then
                                'start the wsm with existing workspace
                                WSM.StartWSM(General.currentjob.Workspace.Substring(General.currentjob.Workspace.LastIndexOf("\") + 1), General.username, False)

                                UnitProps.Show()
                                UnitProps.Activate()
                                BGo.Enabled = False
                                General.seapp.DisplayAlerts = False
                                General.seapp.SetGlobalParameter(SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalStartEnvironmentSyncOrOrdered, SolidEdgePart.ModelingModeConstants.seModelingModeOrdered)
                            End If
                        Else
                            'find order in database
                            Dim orderjobs As List(Of JobInfo) = GetContinueJobs()
                            If orderjobs.Count = 1 Then
                                General.currentjob = Database.GetJobInfo(orderjobs(0).Uid)
                                ContinueSaved()
                            ElseIf orderjobs.Count > 1 Then
                                'show new window, load all jobs in the datagridview and select job
                                JobSelection.Show()
                                JobSelection.FillGrid(orderjobs)
                            End If
                        End If
                    Else
                        General.currentjob.IsERPTest = CheckTest.Checked
                        General.currentjob.IsPDMTest = False
                        If CBCSGMode.Text = "Batch" Then
                            Order.SearchXML(General.currentjob)
                            If General.currentjob.Path <> "" Then
                                Dim oldjobs As List(Of JobInfo) = Database.CheckNewBatch(General.currentjob)
                                If oldjobs.Count = 0 Then
                                    Database.AddJob(General.currentjob)
                                    BGo.Text = "Queued"
                                    BGo.BackColor = Color.Green
                                Else
                                    ShowMessageJob("Batch mode for this job not possible, because the unit was done at least once already!", oldjobs, 1)
                                    BGo.Text = "Failed"
                                    BGo.BackColor = Color.Red
                                End If
                            Else
                                MsgBox("Couldn't find the xml file!")
                                BGo.Text = "Failed"
                                BGo.BackColor = Color.Red
                            End If
                        Else
                            If CBUnit.Text.Contains("Family") Then
                                Order.SearchXML(General.currentjob)
                                If General.localpcffile <> "" Then
                                    Order.ConstructDataModel()
                                    General.isProdavit = True
                                Else
                                    GCOData(CBUnit.Text)
                                End If
                            Else
                                'create the current unit
                                GCOData(CBUnit.Text)
                            End If

                            If Not CBCSGMode.Text = "ETO Batch" Then
                                'start WSM
                                Order.CreateWS(General.currentjob, General.username, False)

                                General.currentjob.Uid = Database.AddJob(General.currentjob)
                            End If

                            'try to get the BOM from production order (LN Webrequest)
                            General.currentunit.BOMList = Order.ConstructBOM(Order.LNProdOrder("web_bom", "dxv?0cc5xot7jcz-1gAQ", TBOrder.Text, TBPos.Text, General.currentjob.Plant))

                            General.seapp.DisplayAlerts = False
                            'load next window
                            If CBCSGMode.Text = "Desktop" Or CBCSGMode.Text = "ETO Batch" Then
                                UnitProps.Show()
                                UnitProps.Activate()
                                General.seapp.SetGlobalParameter(SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalStartEnvironmentSyncOrOrdered, SolidEdgePart.ModelingModeConstants.seModelingModeOrdered)
                            Else
                                CopyProps.Activate()
                                CopyProps.Show()
                            End If
                            BGo.Enabled = False
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub ContinueSaved()
        Dim payload As String

        'start WSM
        Order.CreateWS(General.currentjob, General.username, False)

        'load masterassembly
        WSM.CheckoutPart(Database.GetValue("batch_csg", "PDMID", "uid", General.currentjob.Uid), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile, "asm")

        UnitProps.Show()
        UnitProps.Activate()
        OrderProps.BGo.Enabled = False
        General.seapp.DisplayAlerts = False
        General.seapp.SetGlobalParameter(SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalStartEnvironmentSyncOrOrdered, SolidEdgePart.ModelingModeConstants.seModelingModeOrdered)

        'import json
        payload = ImportJsonFile(General.currentjob.PDMID)
        If payload = "" Then
            UnitProps.BSave.Enabled = False
            UnitProps.BImport.Visible = True
        Else
            General.OverrideData(payload)
        End If

    End Sub

    Shared Function GetContinueJobs() As List(Of JobInfo)
        Dim alljobs As List(Of JobInfo) = Database.CheckNewBatch(General.currentjob)
        Dim p1jobs = From ajobs In alljobs Where ajobs.Prio = 1

        Return p1jobs.ToList
    End Function

    Shared Function GetUnits(unittype As String) As String()
        Dim unitlist As String()

        If unittype = "Condenser" Then
            unitlist = {"GCH C/V - Family", "GCD C/V - Family", "Others"}
        Else
            unitlist = {"GACV - Family", "Others"}
        End If

        Return unitlist
    End Function

    Shared Sub GCOData(unittype As String)
        General.isProdavit = False
        With General.currentunit
            If unittype.Contains("Condenser") Or unittype.Contains("C/V") Then
                .ApplicationType = "Condenser"
            ElseIf unittype = "GCO" Then
                .ApplicationType = "GCO"
            Else
                .ApplicationType = "Evaporator"
            End If
            With .OrderData
                .OrderNumber = OrderProps.TBOrder.Text
                .OrderPosition = OrderProps.TBPos.Text
                .OrderDir = General.currentjob.OrderDir
            End With
            .ModelRangeName = "NNNN"
            .ModelRangeSuffix = "NN"
            If unittype.Contains("VShape") Or unittype.Contains("GxD") Then
                .UnitDescription = "VShape"
            ElseIf unittype.Contains("2 Coil") Or unittype.Contains("GADC") Then
                .UnitDescription = "Dual"
            End If
            .SELanguageID = General.userlangID
        End With
    End Sub

    Private Sub OrderProps_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

    End Sub

    Private Sub BJob_Click(sender As Object, e As EventArgs) Handles BJob.Click
        DBInfo.Show()
        DBInfo.Activate()
    End Sub

    Private Sub CBCSGMode_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBCSGMode.SelectedIndexChanged
        If CBCSGMode.Text <> "" Then
            With General.currentjob
                .Users = General.username
                .Prio = 1
                If CBCSGMode.Text = "Batch" Then
                    .Prio = 3
                ElseIf CBCSGMode.Text = "Copy Order" Then
                    .Prio = 4
                ElseIf CBCSGMode.Text = "ETO Batch" Then
                    .Prio = 5
                End If
            End With
        End If
    End Sub

    Shared Function ImportJsonFile(Optional reffilename As String = "") As String
        Dim ofdialog As New OpenFileDialog
        Dim payload As String = ""
        Dim jsonname As String

        Try
            If reffilename <> "" Then
                'check, if json for masterassembly exists
                jsonname = General.GetFullFilename(General.currentjob.Workspace, reffilename, "json")
                If jsonname <> "" Then
                    Try
                        Dim sr As New StreamReader(jsonname)
                        payload = sr.ReadToEnd
                    Catch ex As Exception

                    End Try
                End If
            End If

            If payload = "" Then
                ofdialog.InitialDirectory = "C:\Import"
                ofdialog.Filter = "JSON files (*.json)|*.json"
                ofdialog.RestoreDirectory = True
                ofdialog.FilterIndex = 3

                If ofdialog.ShowDialog() = DialogResult.OK Then
                    Dim sr As New StreamReader(ofdialog.FileName)
                    payload = sr.ReadToEnd
                End If
            End If

        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try

        Return payload
    End Function

    Shared Sub ShowMessageJob(message As String, jobs As List(Of JobInfo), messagenumber As Integer)
        Dim fullmessage, jobmessage, prio As String
        Dim job As JobInfo = jobs.First

        jobmessage = "Total Request Count: " + jobs.Count.ToString + vbNewLine +
            "Request Time: " + job.RequestTime.ToString + vbNewLine +
            "User: " + job.Users + vbNewLine +
            "Status: " + job.Status.ToString

        Select Case messagenumber
            Case 1
                If job.Status = 100 Then
                    jobmessage += vbNewLine + "PDMID: " + job.PDMID
                End If
            Case 2
                jobmessage = ""
            Case 5
                jobmessage += vbNewLine + "PDMID: " + job.PDMID
            Case Else
                If job.Prio = 1 Then
                    prio = "Desktop"
                Else
                    prio = "Batch"
                End If
                jobmessage += vbNewLine + "System: " + prio
        End Select

        fullmessage = message + vbNewLine + jobmessage
        MsgBox(fullmessage)
    End Sub

    Shared Function CheckCoreDLL(username As String) As Boolean
        Dim updatefile As Boolean = False
        Dim infostime, localtime As Date

        If Application.StartupPath = "C:\Import" Or username = "mlewin" Then
            If File.Exists(Application.StartupPath + "\CSGCore.dll") Then
                localtime = File.GetLastWriteTimeUtc(Application.StartupPath + "\CSGCore.dll")
                'compare create time
                infostime = File.GetLastWriteTimeUtc(General.infosdir + "\CSGCoreFiles\CSGCore.dll")
                General.dlltime = localtime
                If infostime > localtime Then
                    updatefile = True
                Else
                    localtime = File.GetLastWriteTimeUtc(Application.ExecutablePath)
                    infostime = File.GetLastWriteTimeUtc(General.infosdir + "\CSGCoreFiles\CSG.exe")
                    General.apptime = localtime
                    If infostime > localtime Then
                        updatefile = True
                    End If
                End If
            End If
            If updatefile Then
                MsgBox("A new version of the CSG is available, please copy the CSG.exe and dll.")
            End If
        ElseIf username <> "mlewin" And username <> "csgen" Then
            MsgBox("Start this application from C:\Import")
            updatefile = True
        End If
        Return updatefile
    End Function

    Private Sub InitAuthUsers()
        'authusers.AddRange({"lgortva"})
    End Sub

    Shared Function Controllength(tbname As String, tbcontent As String) As Boolean
        Dim maxlength As Integer
        Dim lengthok As Boolean = True

        maxlength = MaxChars(tbname)

        If tbcontent.Length > maxlength Then
            lengthok = False
            LengthErrorMsg(tbname)
        End If

        Return lengthok
    End Function

    Shared Function MaxChars(tbname As String) As Integer
        Dim maxlength As Integer
        Select Case tbname
            Case "TBOrder"
                maxlength = 9
            Case "TBPos"
                maxlength = 4
            Case "TBNr"
                maxlength = 9
        End Select
        Return maxlength
    End Function

    Shared Sub LengthErrorMsg(tbname As String)
        Select Case tbname
            Case "TBOrder"
                MsgBox("Order number too long, max length is 9 characters!")
            Case "TBPos"
                MsgBox("Position too long, max length is 4 characters!")
            Case "TBNr"
                MsgBox("Project number too long, max length is 9 characters!")
        End Select
    End Sub

    Private Sub OrderProps_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Try
            If Directory.Exists(General.currentjob.OrderDir) And General.currentjob.Uid > 0 Then
                With General.currentunit
                    .DecSeperator = General.decsym
                    .OrderData = General.currentjob
                    .IsProdavit = General.isProdavit
                    .APPVersion = General.apptime.ToString
                    .DLLVersion = General.dlltime.ToString
                End With
                General.WriteDatatoFile(General.currentjob.OrderDir + "\LogClose")
                General.SendDataToWebservice("http://deffbswap14.europe.guentner-corp.com:800/api/values/save-log", "POST")
            End If
            If General.seapp IsNot Nothing Then
                General.seapp.DisplayAlerts = True
            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub
End Class
