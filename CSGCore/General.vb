Imports System.IO
Imports System.Management
Imports System.Net.Mail
Imports Newtonsoft.Json

Public Class General
    Public Shared decsym, username, userdomain, domainname, gpcdir, csgdir, infosdir, batfile, localpcffile, errorlogfile, actionlogfile As String
    Public Shared dlltime, apptime As Date
    Public Shared userlangID As Integer
    Public Shared isProdavit As Boolean
    Public Shared currentjob As New JobInfo
    Public Shared currentunit As New UnitData
    Public Shared coillist As New List(Of CoilData)
    Public Shared circuitlist As New List(Of CircuitData)
    Public Shared consyslist As New List(Of ConSysData)
    Public Shared seapp As SolidEdgeFramework.Application = Nothing
    Public Shared Buttonlist, CBlist As New List(Of String)

    Shared Sub InitDefaultSettings(apppath As String)
        csgdir = "C:\Import"
        batfile = csgdir + "\cdbwscall_local.bat"
        localpcffile = ""
        isProdavit = True
        StartSE()
        'search for mandatory .dll files
        If apppath = "C:\Import" Then
            WSM.CheckWSFiles(csgdir, infosdir)
        ElseIf username <> "mlewin" Then
            MsgBox("Please start this application from the directory C:\Import")
        End If
    End Sub

    Shared Sub StartSE()
        seapp = SEUtils.Connect
    End Sub

    Shared Sub SetInfosDir()
        If domainname.ToLower = "asia" Then
            infosdir = "\\10.11.11.5\id\PAS\Others\CSG_Files"
        Else
            infosdir = "\\fs01.europe.guentner-corp.com\infos\CSG_Files"
        End If
    End Sub

    Shared Sub SetGPCDir(ordernumber As String)
        If domainname.ToLower = "asia" Then
            If ordernumber.Substring(0, 1) = "C" Then
                gpcdir = "\\10.51.105.14\log_global_cn\log_gpc"
            Else
                gpcdir = "\\10.51.105.14\log_global\log_gpc"
            End If
        Else
            If Directory.Exists("\\deffbslap14.europe.guentner-corp.com\log_global\log_gpc") Then
                gpcdir = "\\deffbslap14.europe.guentner-corp.com"
            Else
                gpcdir = "\\10.1.105.14"
                If Not Directory.Exists(gpcdir + "\log_global") Then
                    gpcdir = "\\deffbslap14"
                    If Not Directory.Exists(gpcdir + "\log_global") Then
                        Debug.Print("no gpc dir found")
                    End If
                End If
            End If
            gpcdir += "\log_global\log_gpc"
            If currentjob.IsERPTest Then
                gpcdir = gpcdir.Replace("14", "14-104")
            End If
        End If
    End Sub

    Shared Function GetPCFFileEnding() As String
        Dim fileending As String
        If domainname = "europe" Then
            fileending = ".pcfgeux"
        Else
            fileending = ".pcfgapx"
        End If
        Return fileending
    End Function

    Shared Function StringToLists(rawstring As String, seperator() As String) As List(Of String)()
        Dim namelist, valuelist As New List(Of String)
        Dim proplists() As List(Of String)
        Dim stext(), props() As String

        Try
            stext = rawstring.Split(seperator, 0)

            For Each entry In stext
                If entry.Contains("value") Then
                    entry = entry.Replace(Chr(34), "")
                    'Search the property name and the value
                    props = GetXMLInfo(entry)
                    If props(0) <> "" And props(1) <> "" Then
                        namelist.Add(props(0))
                        valuelist.Add(props(1))
                    End If
                End If
            Next
        Catch ex As Exception

        End Try
        proplists = {namelist, valuelist}
        Return proplists

    End Function

    Shared Function GetXMLInfo(xmlline As String) As String()
        Dim name, value, subtext, infostring() As String
        Dim startpos, charlength As Integer

        Try
            startpos = xmlline.IndexOf("name=") + 5
            subtext = xmlline.Substring(startpos)
            charlength = subtext.IndexOf(" ")
            name = subtext.Substring(0, charlength)
            startpos = subtext.IndexOf("value=") + 6
            subtext = subtext.Substring(startpos)
            charlength = subtext.IndexOf(" ")
            value = subtext.Substring(0, charlength)

        Catch ex As Exception

        End Try

        infostring = {name, value}
        Return infostring
    End Function

    Shared Function TextToDouble(text As String) As Double
        Dim doublevalue As String
        If text = "" Then
            doublevalue = 0
        Else
            If decsym = "," Then
                doublevalue = CDbl(text.Replace(".", ","))
            Else
                doublevalue = CDbl(text.Replace(",", "."))
            End If
        End If

        Return doublevalue
    End Function

    Shared Sub CreateLogEntry(exmsg As String, <Runtime.CompilerServices.CallerMemberName> Optional memberName As String = Nothing)
        Dim loopcount As Integer = 0


        Try
            My.Computer.FileSystem.WriteAllText(errorlogfile, "Error in " + memberName + vbNewLine, True)

            If username = "csgen" AndAlso (exmsg.Contains("RPC") Or exmsg.Contains("disconnect") Or exmsg.ToLower.Contains("vault")) Then
                'kill SE, reset job
                seapp.Documents.Close()
                seapp.DoIdle()
                WSM.KillTask("Edge")

                If memberName = "RenameFiles" Then
                    My.Computer.FileSystem.WriteAllText(errorlogfile, exmsg + vbNewLine, True)
                End If

                'Do
                '    ContinueBatch()
                '    loopcount += 1
                '    currentjob.Status = Database.GetValue("CSG.batch_csg", "status", "uid", currentjob.Uid, "int")
                'Loop Until currentjob.Status = 100 Or loopcount = 5

                'If loopcount = 5 And currentjob.Status < 100 Then
                '    SendErrorMail("Error in new CSG Batch", "Connection to SE lost, restarting the job" + currentjob.Uid.ToString + "." + vbNewLine + "Sub: " + memberName)
                '    Database.UpdateJob(currentjob, {"Status"}, {"0"})
                '    Environment.Exit(0)
                'End If

                SendErrorMail("Error in new CSG Batch", "Connection to SE lost, restarting the job" + currentjob.Uid.ToString + "." + vbNewLine + "Sub: " + memberName)
                Database.UpdateJob(currentjob, {"Status"}, {"0"})

            ElseIf memberName = "ConstructBOM" Then
                My.Computer.FileSystem.WriteAllText(errorlogfile, exmsg + vbNewLine, True)
            End If
        Catch ex As Exception
            Console.WriteLine("Error in log entry" + vbNewLine + ex.ToString)
            'Console.ReadLine()
        End Try
    End Sub

    Shared Sub ContinueBatch()
        Dim objProcess As New Process
        Try
            objProcess.StartInfo.FileName = "C:\Import\ContinueBatch.exe"
            objProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal
            objProcess.StartInfo.Arguments = currentjob.Uid
            objProcess.Start()

            'Wait until the process passes back an exit code 
            objProcess.WaitForExit()

            'Free resources associated with this process
            objProcess.Close()
        Catch
            'Logfile entry
        End Try
    End Sub

    Shared Sub CreateActionLogEntry(formname As String, sendername As String, actionname As String, Optional changedvalue As String = "")
        Dim currenthour, currentmin, currentsec As Integer
        Dim currentdate As Date = Date.Today
        Dim timestp As String
        currenthour = Date.Now.Hour
        currentmin = Date.Now.Minute
        currentsec = Date.Now.Second

        Try
            timestp = currenthour.ToString + "_" + currentmin.ToString + "_" + currentsec.ToString
            If actionname = "pressed" Or changedvalue <> "" Then
                My.Computer.FileSystem.WriteAllText(actionlogfile, timestp + ": " + formname + " - " + sendername + " - " + actionname, True)
                If changedvalue <> "" Then
                    My.Computer.FileSystem.WriteAllText(actionlogfile, " to " + changedvalue, True)
                    CBlist.Add(formname + "_" + sendername)
                ElseIf actionname = "pressed" Then
                    Buttonlist.Add(formname + "_" + sendername)
                End If
                My.Computer.FileSystem.WriteAllText(actionlogfile, vbNewLine, True)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Shared Function KillTask(procname As String) As Boolean
        Dim success As Boolean = False

        Try
            Dim proc As Process = FindTask(procname)

            If proc IsNot Nothing Then
                If procname.Contains("Workspaces") Then
                    proc.CloseMainWindow()
                Else
                    proc.Kill()
                End If
                Threading.Thread.Sleep(5000)
                proc = FindTask(procname)
                If proc Is Nothing Then
                    success = True
                Else
                    proc.Kill()
                    Threading.Thread.Sleep(5000)
                End If
            End If

        Catch ex As Exception

        End Try

        Return success
    End Function

    Shared Function FindTask(procname As String) As Process
        Dim processes() As Process = Process.GetProcesses
        Dim owner, progID As String
        Dim targetproc As Process = Nothing

        Try
            progID = ""
            For Each proc In processes
                If proc.ProcessName = procname Then
                    owner = GetProcessOwner(proc.Id)
                    If owner = Environment.UserDomainName + "\" + Environment.UserName Then
                        targetproc = proc
                        Exit For
                    End If
                End If
            Next
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try

        Return targetproc
    End Function

    Shared Function GetProcessOwner(ByVal processId As Integer) As String
        Dim query As String = "Select * From Win32_Process Where ProcessID = " & processId
        Dim searcher As New ManagementObjectSearcher(query)
        Dim processList As ManagementObjectCollection = searcher.[Get]()

        For Each obj As ManagementObject In processList
            Dim argList As String() = New String() {String.Empty, String.Empty}
            Dim returnVal As Integer = Convert.ToInt32(obj.InvokeMethod("GetOwner", argList))

            If returnVal = 0 Then
                Return argList(1) & "\" & argList(0)
            End If
        Next

        Return "NO OWNER"
    End Function

    Shared Function WaitForFile(dir As String, filename As String, fileending As String, Optional maxloops As Integer = -1) As Boolean
        Dim filefound As Boolean = False
        Dim loopcount As Integer = 0

        While filefound = False
            For Each singlefile As String In Directory.GetFiles(dir)
                If singlefile.Contains(filename) And singlefile.Contains(fileending) And Not singlefile.Contains(fileending + ".") Then
                    filefound = True
                    Exit For
                End If
            Next
            Threading.Thread.Sleep(100)
            loopcount += 1
            If maxloops > -1 Then
                If loopcount > maxloops Then
                    Exit While
                End If
            End If
        End While

        Return filefound
    End Function

    Shared Function FindProcess() As Process
        Dim totaltasks() As Process
        Dim currenttask As String
        Dim wsmprocess As New Process

        totaltasks = Process.GetProcesses()

        For Each task As Process In totaltasks
            currenttask = task.ToString()
            If currenttask.Contains("Workspace") Then
                wsmprocess = task
            End If
        Next

        Return wsmprocess
    End Function

    Shared Function GetFullFilename(workspace As String, partialname As String, fileending As String) As String
        Dim fullfile As String = ""
        Dim falseending As String = fileending + "."
        If fileending.Substring(0, 1) <> "." Then
            fileending = "." + fileending
        End If

        If partialname.Contains("774198") And partialname.Length = 10 Then
            partialname = "7741980001-"
        End If
        If partialname.Contains("808394") And partialname.Length = 10 Then
            partialname = "8083940001-"
        End If

        If partialname.Contains("309896") And File.Exists(workspace + "\00003098960001-.par") = False Then
            'MsgBox("Please add the model for *309896 manually to the workspace before you continue!")
        End If

        For Each file As String In Directory.GetFiles(workspace)
            If file.Contains(partialname) And file.Contains(fileending) And file.Contains(falseending) = False Then
                fullfile = file
            End If
        Next

        Return fullfile
    End Function

    Shared Sub ReleaseObject(obj As Object)

        Try
            Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception

        End Try

    End Sub

    Shared Function IntegerRem(div As Double, pitch As Double) As Integer
        Dim multiplierdiv As Integer = 1
        Dim multiplierp As Integer = 1
        Dim mp As Integer
        Dim divm As Double = div
        Dim pitchm As Double = pitch
        Dim remainder As Integer
        Dim loopexit As Boolean = False
        Dim modxy As Integer
        Dim counter As Integer = 0
        Dim value1, value2 As Double

        Do
            If Math.Abs(divm) Mod 1 > 0 Then
                multiplierdiv *= 10
                divm = Math.Round(div * multiplierdiv, 3)
                loopexit = False
                counter += 1
            ElseIf divm Mod 1 = 0 Or counter > 7 Then
                loopexit = True
            End If
        Loop Until loopexit
        loopexit = False
        counter = 0

        Do
            If pitchm Mod 1 > 0 Then
                multiplierp *= 10
                pitchm = Math.Round(pitch * multiplierp, 3)
                loopexit = False
                counter += 1
            ElseIf pitchm Mod 1 = 0 Or counter > 7 Then
                loopexit = True
            End If
        Loop Until loopexit

        mp = Math.Max(multiplierp, multiplierdiv)
        value1 = Math.Round(div * mp, 1)
        value2 = Math.Round(pitch * mp, 1)
        modxy = Math.DivRem(CInt(value1), CInt(value2), remainder)

        Return modxy
    End Function

    Shared Function IntegerMod(div As Double, pitch As Double) As Double
        Dim multiplierdiv As Integer = 1
        Dim multiplierp As Integer = 1
        Dim mp As Integer
        Dim divm As Double = div
        Dim pitchm As Double = pitch
        Dim loopexit As Boolean = False
        Dim modxy As Double
        Dim counter As Integer = 0
        Dim value1, value2 As Double

        Do
            If Math.Abs(divm) Mod 1 > 0 Then
                multiplierdiv *= 10
                divm = Math.Round(div * multiplierdiv, 3)
                loopexit = False
                counter += 1
            ElseIf divm Mod 1 = 0 Or counter > 7 Then
                loopexit = True
            End If
        Loop Until loopexit
        loopexit = False
        counter = 0

        Do
            If pitchm Mod 1 > 0 Then
                multiplierp *= 10
                pitchm = Math.Round(pitch * multiplierp, 3)
                loopexit = False
                counter += 1
            ElseIf pitchm Mod 1 = 0 Or counter > 7 Then
                loopexit = True
            End If
        Loop Until loopexit

        mp = Math.Max(multiplierp, multiplierdiv)
        value1 = Math.Round(div * mp, 1)
        value2 = Math.Round(pitch * mp, 1)
        modxy = Math.Round(value1 Mod value2, 3)

        Return modxy
    End Function

    Shared Function GetUniqueStrings(valuelist As List(Of String), Optional value2list As List(Of String) = Nothing) As List(Of String)
        Dim skipthis As Boolean
        Dim resultlist, tempvaluelist As New List(Of String)

        If valuelist IsNot Nothing Then
            For Each entry In valuelist
                tempvaluelist.Add(entry)
            Next
        End If

        If value2list IsNot Nothing Then
            For Each entry In value2list
                tempvaluelist.Add(entry)
            Next
        End If

        For Each Value In tempvaluelist
            skipthis = False
            If resultlist.Count = 0 Then
                resultlist.Add(Value)
            Else
                For Each Entry In resultlist
                    If Entry = Value Then
                        skipthis = True
                    End If
                Next
                If skipthis = False Then
                    resultlist.Add(Value)
                End If
            End If
        Next

        Return resultlist
    End Function

    Shared Function GetUniqueValues(ByVal valuelist As List(Of Double)) As List(Of Double)
        Dim skipthis As Boolean
        Dim resultlist As New List(Of Double)

        For Each Value In valuelist
            skipthis = False
            If resultlist.Count = 0 Then
                resultlist.Add(Value)
            Else
                For Each Entry In resultlist
                    If Entry = Value Then
                        skipthis = True
                    End If
                Next
                If skipthis = False Then
                    resultlist.Add(Value)
                End If
            End If
        Next

        Return resultlist
    End Function

    Shared Function GetShortName(fullname As String) As String
        Dim shortname As String = fullname.Substring(fullname.LastIndexOf("\") + 1)
        Return shortname
    End Function

    Shared Function ConvertList(textlist As List(Of String)) As List(Of Double)
        Dim doublelist As New List(Of Double)
        Dim doublevalue As Double

        For Each entry As String In textlist
            doublevalue = TextToDouble(entry)
            doublelist.Add(doublevalue)
        Next

        Return doublelist
    End Function

    Shared Function FindOccInList(occlist As List(Of PartData), occname As String) As Integer
        Dim occindex As Integer = -1
        For i As Integer = 0 To occlist.Count - 1
            If occlist(i).Occname.Substring(0, occlist(i).Occname.IndexOf(":")).Contains(occname) Then
                occindex = occlist(i).Occindex
                Exit For
            End If
        Next
        Return occindex
    End Function

    Shared Function GetConfigname(occlist As List(Of PartData), occname As String) As String
        Dim configref As String = ""
        For i As Integer = 0 To occlist.Count - 1
            If occlist(i).Occname.Contains(occname) Then
                configref = occlist(i).Configref
                Exit For
            End If
        Next
        Return configref
    End Function

    Shared Function RenameFiles() As Integer
        Dim masterasmdoc, coilasmdoc As SolidEdgeAssembly.AssemblyDocument
        Dim coilnames As New Dictionary(Of String, String)
        Dim checkforbows As Boolean = False
        Dim coildft, consysdft As String

        Try
            'masterassembly should be open in SE with new name
            masterasmdoc = seapp.ActiveDocument

            'update jobtable with PDMID
            Database.UpdateJob(currentjob, {"Status", "PDMID"}, {"50", masterasmdoc.Name.Substring(0, 10)})

            currentunit.UnitFile.Fullfilename = masterasmdoc.FullName
            currentjob.PDMID = GetShortName(masterasmdoc.Name)

            For Each objocc As SolidEdgeAssembly.Occurrence In masterasmdoc.Occurrences
                For i As Integer = 0 To currentunit.Occlist.Count - 1
                    If objocc.Index = currentunit.Occlist(i).Occindex Then
                        coilnames.Add(currentunit.Occlist(i).Occname.Substring(0, currentunit.Occlist(i).Occname.IndexOf(":")), objocc.PartFileName)
                    End If
                Next
            Next

            seapp.Documents.CloseDocument(masterasmdoc.FullName, SaveChanges:=False, DoIdle:=False)

            For Each c In currentunit.Coillist
                'replace the filename in the class
                For Each entry In coilnames
                    If c.CoilFile.Fullfilename.Contains(entry.Key) Then
                        c.CoilFile.Fullfilename = entry.Value
                    End If
                Next

                coilasmdoc = seapp.Documents.Open(c.CoilFile.Fullfilename)
                seapp.DoIdle()

                For i As Integer = 0 To c.Frontbowids.Count - 1
                    If c.Frontbowids(i).Contains(Library.TemplateParts.BOW1) OrElse c.Frontbowids(i).Contains(Library.TemplateParts.BOW9) Then
                        checkforbows = True
                        Exit For
                    End If
                Next

                If Not checkforbows Then
                    For i As Integer = 0 To c.Backbowids.Count - 1
                        If c.Backbowids(i).Contains(Library.TemplateParts.BOW1) OrElse c.Backbowids(i).Contains(Library.TemplateParts.BOW9) Then
                            checkforbows = True
                            Exit For
                        End If
                    Next
                End If

                coildft = GetFullFilename(currentjob.Workspace, GetShortName(c.CoilFile.Fullfilename).Substring(0, 10), "dft")
                If checkforbows Then
                    'rename the drawing views containing template parts
                    SEDrawing.ChangeCaptionForBowViews(c, coilasmdoc, coildft)
                End If

                'replace the consys names
                For i As Integer = 1 To c.Circuits.Count
                    Dim csindex As Integer = FindOccInList(c.Occlist, "Consys" + i.ToString + "_" + c.Number.ToString)
                    If csindex > -1 Then
                        c.ConSyss(i - 1).ConSysFile.Fullfilename = coilasmdoc.Occurrences.Item(csindex).PartFileName

                        'open and update consys dft
                        consysdft = GetFullFilename(currentjob.Workspace, GetShortName(c.ConSyss(i - 1).ConSysFile.Fullfilename).Substring(0, 10), "dft")
                        SEDrawing.UpdateDrawingAfterSave(consysdft)
                    End If
                Next

                'open and update coil dft
                SEDrawing.UpdateDrawingAfterSave(coildft)

                If currentunit.ModelRangeName = "GACV" And isProdavit Then
                    'check for coversheet
                    For Each consys In c.ConSyss
                        If consys.CoverSheetCutouts.Count > 0 Then
                            'is not renamed yet, because not saved in PDM
                            If File.Exists(consys.CoverSheetCutouts.First.Filename) Then
                                WSM.SaveCS(consys.CoverSheetCutouts.First.Filename)
                            End If
                        End If
                    Next
                End If
                WSM.SaveDFT()
                WSM.WaitforWSMDialog()

                seapp.Documents.CloseDocument(c.CoilFile.Fullfilename, SaveChanges:=False, DoIdle:=True)

                Database.UpdateJob(currentjob, {"Status", "FinishedTime", "Saved"}, {"100", Database.ConvertDatetoStr(Date.UtcNow), "true"})
                currentjob.Status = 100
            Next

        Catch ex As Exception
            Database.UpdateJob(currentjob, {"Status"}, {Library.JobStatus.ERRORSAVE})
            CreateLogEntry(ex.ToString)
            currentjob.Status = Library.JobStatus.ERRORSAVE
        End Try

        Return currentjob.Status
    End Function

    Shared Sub WriteDatatoFile(path As String)
        Dim payload, filename As String
        Dim utime As Integer

        Try
            currentunit.APPVersion = apptime.ToString
            currentunit.DLLVersion = dlltime.ToString
            utime = (Date.UtcNow - New DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds
            filename = path + "_" + utime.ToString + ".json"
            payload = CreatePayload(currentunit)
            My.Computer.FileSystem.WriteAllText(filename, payload, False)
        Catch ex As Exception

        End Try
    End Sub

    Shared Function CreatePayload(unit As UnitData) As String
        Return JsonConvert.SerializeObject(unit)
    End Function

    Shared Function DeserializeUnit(payload As String) As UnitData
        Dim newunit As New UnitData
        Try
            newunit = JsonConvert.DeserializeObject(Of UnitData)(payload)
        Catch ex As Exception

        End Try
        Return newunit
    End Function

    Shared Sub SendErrorMail(subject As String, body As String)
        Dim SmtpPort As Integer = 25
        Dim MailTo As String = "Martin.Lewin@spark-radiance.eu"
        Dim MailFrom As String = "cad-konfigurator@guentner.com"
        Dim ServerIP As String = "10.1.11.135"
        Dim client As New SmtpClient(ServerIP, SmtpPort)
        Dim eMail As New MailMessage

        Try
            eMail.Subject = subject
            eMail.From = New MailAddress(MailFrom)
            eMail.To.Add(MailTo)
            eMail.To.Add("Hagen.Conrad@spark-radiance.eu")
            eMail.To.Add("Daniel.Pereira@spark-radiance.eu")
            eMail.Body = body
            client.Send(eMail)
        Catch ex As Exception

        End Try

    End Sub

    Shared Sub GetLanguage()
        Try
            'userlangID 1031 - Deutsch // 1033 - English // 1038 - Hungarian // 1048 - Romanian
            seapp.GetGlobalParameter(SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalDefaultUserLangID, userlangID)
        Catch ex As Exception

        End Try
    End Sub

    Shared Function SendDataToWebservice(website As String, modifier As String) As String
        Dim response As String = ""
        Try
            response = Order.CSGWebservice(website, currentunit, modifier)
        Catch ex As Exception
            CreateLogEntry(ex.ToString)
        End Try
        Return response
    End Function

    Shared Sub OverrideData(payload As String)

        Try
            currentunit = DeserializeUnit(payload)
            currentjob = GetPropValue(currentunit, "OrderData")
            errorlogfile = currentjob.OrderDir + "\" + currentjob.OrderNumber + "_" + currentjob.OrderPosition + "_error.txt"
            actionlogfile = currentjob.OrderDir + "\" + currentjob.OrderNumber + "_" + currentjob.OrderPosition + "_action.txt"
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Shared Function GetPropValue(src As Object, propName As String) As Object
        Return src.[GetType]().GetProperty(propName).GetValue(src, Nothing)
    End Function

    Private Sub SetPropValue(src As Object, propName As String, value As Object)
        src.[GetType]().GetProperty(propName).SetValue(src, value)
    End Sub

    Shared Sub CreateWinDir(path As String)
        Try
            If Not Directory.Exists(path) Then
                Directory.CreateDirectory(path)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Shared Sub CopyDir(sourcdir As String, targetdir As String)
        Dim dir As New DirectoryInfo(sourcdir)
        Dim dirs As DirectoryInfo() = dir.GetDirectories

        Try
            Directory.CreateDirectory(targetdir)

            For Each f As FileInfo In dir.GetFiles()
                File.Copy(f.FullName, targetdir + "\" + f.Name)
            Next

            For Each subdir In dirs
                Dim newtargetdir As String = targetdir + "\" + subdir.Name
                CopyDir(subdir.FullName, newtargetdir)
            Next
        Catch ex As Exception
            CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub DeleteFolder(folder As String)
        Dim removelist As New List(Of String)

        Try
            Dim dir As New DirectoryInfo(folder)
            Try
                dir.Attributes = FileAttributes.Normal
                SetNormal(dir)
                Directory.Delete(folder, True)
            Catch ex As Exception
                Debug.Print(ex.ToString)
            End Try
        Catch ex As Exception

        End Try

    End Sub

    Shared Sub SetNormal(dir As DirectoryInfo)
        For Each subdir In dir.GetDirectories
            SetNormal(subdir)
            subdir.Attributes = FileAttributes.Normal
        Next
        For Each f In dir.GetFiles
            f.Attributes = FileAttributes.Normal
        Next
    End Sub

    Shared Function TextUpperCase(reftext As String) As String
        Dim newtext As String
        newtext = reftext.Substring(0, 1).ToUpper + reftext.Substring(1)
        Return newtext
    End Function
End Class
