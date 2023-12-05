Imports System.IO
Imports Newtonsoft.Json
Imports System.Management

Public Class WSM
    Public Shared fullpartids As New List(Of String)
    Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Integer) As Integer

    Shared Sub CheckWSFiles(csgdir As String, infosdir As String)
        Dim batfile, wscallfile, batcontent, nsjsondll, oracledll As String
        Dim subinfos(), subinfo, tempname, mandfiles() As String
        Dim strmfile As StreamWriter
        Dim enc As Text.Encoding = Text.Encoding.GetEncoding(1252)  'ANSI encoder necessary, standard is UTF-8

        batfile = csgdir + "\cdbwscall_local.bat"
        wscallfile = "checkout_cdbwscall.cmd"
        nsjsondll = "Newtonsoft.Json.dll"
        oracledll = "Oracle.ManagedDataAccess.dll"

        mandfiles = {wscallfile, nsjsondll, oracledll}

        If Not File.Exists(batfile) Then
            batcontent = wscallfile + " %1 %2"

            strmfile = New StreamWriter(batfile, False, enc)
            strmfile.WriteLine(batcontent)
            strmfile.Close()
        End If

        For Each mandfile In mandfiles
            Dim i As Integer = 0
            Dim loopexit As Boolean = False

            If Not File.Exists(csgdir + "\" + mandfile) Then
                'find file in infos subfolder
                subinfos = Directory.GetDirectories(infosdir)

                Do
                    subinfo = subinfos(i)
                    If subinfo.Contains("_Appl") Then
                        loopexit = True
                        tempname = subinfo + "\" + mandfile
                        If File.Exists(tempname) Then
                            File.Copy(tempname, csgdir + "\" + mandfile)
                        End If
                    End If

                    i += 1
                    If i = subinfos.Count Then
                        loopexit = True
                    End If
                Loop Until loopexit
            End If
        Next

    End Sub

    Shared Function StartWSM(wsmfolder As String, username As String, Optional isCDBtest As Boolean = False) As String
        Dim newprocess As New Process
        Dim wsmdir, wsmexe As String
        Dim tasklist As New List(Of String)
        Dim count As Integer
        Dim loopexit As Boolean = False
        Dim existed As Boolean = False

        Try

            wsmdir = "C:\Users\" + username + "\WSM\"

            For Each f As String In Directory.GetDirectories(wsmdir)
                If f = wsmdir + wsmfolder Then
                    existed = True
                    Exit For
                End If
            Next

            General.KillTask("WorkspacesDesktop")
            wsmexe = "C:\Program Files\CONTACT Workspaces Desktop\bin\WorkspacesDesktop.exe"
            If Not existed Then
                Directory.CreateDirectory(wsmdir + wsmfolder)
                Dim dir As New DirectoryInfo(wsmdir + wsmfolder) With {
                    .Attributes = FileAttributes.Normal
                }
            End If

            newprocess.StartInfo.FileName = wsmexe
            newprocess.StartInfo.Arguments = "--context Konfigurator " + wsmdir + wsmfolder
            newprocess.Start()
            newprocess.WaitForInputIdle(True)

            Console.WriteLine("finished launch")
            Threading.Thread.Sleep(5000)
            Console.WriteLine("Process ID: " + newprocess.Id.ToString)

            Do
                'find the loading workspace window
                Dim firsthwd As IntPtr = WindowsAPI.FindWindowHandle("Loading workspace", True)

                If firsthwd <> IntPtr.Zero Then
                    Console.WriteLine("First Handle found")
                    WindowsAPI.SetForegroundWindow(firsthwd)
                    WindowsAPI.PostMessage(firsthwd, WindowsAPI.Messages.WM_KEYDOWN, WindowsAPI.VKeys.KEY_Return, 0)
                    Threading.Thread.Sleep(1000)
                    Console.WriteLine("Window confirmed")

                    If Not existed Then
                        'find the created workspace window
                        Dim secondhwd As IntPtr = WindowsAPI.FindWindowHandle("Create new Workspace", True)

                        If secondhwd <> IntPtr.Zero Then
                            Console.WriteLine("Second Handle found")
                            WindowsAPI.SetForegroundWindow(secondhwd)
                            WindowsAPI.PostMessage(secondhwd, WindowsAPI.Messages.WM_KEYDOWN, WindowsAPI.VKeys.KEY_Return, 0)
                            Console.WriteLine("Window confirmed")
                        End If
                    End If
                End If

                Threading.Thread.Sleep(2000)
                If Directory.Exists(wsmdir + wsmfolder) Then
                    loopexit = True
                    Console.WriteLine("WS Created: " + wsmdir + wsmfolder)
                End If

                count += 1
                Console.WriteLine("Count #" + count.ToString)
                If count > 20 Then
                    loopexit = True
                End If
            Loop Until loopexit
            wsmfolder = wsmdir + wsmfolder
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
        Console.WriteLine("Start WSM Method finished")

        Return wsmfolder
    End Function

    Shared Function CreateWSName() As String
        Dim currenthour, currentmin, currentsec As Integer
        Dim today, wsmfolder As String
        Dim currentdate As Date = Date.Today

        currenthour = Date.Now.Hour
        currentmin = Date.Now.Minute
        currentsec = Date.Now.Second
        today = currentdate.ToShortDateString

        today = today.Replace(".", "_")
        today = today.Replace("/", "_")
        today = today.Replace(" ", "")
        wsmfolder = "CSG_" + today + "_" + currenthour.ToString + "_" + currentmin.ToString + "_" + currentsec.ToString

        Return wsmfolder
    End Function

    Shared Sub CheckoutPart(singlepdmid As String, orderdir As String, workspace As String, batfile As String, Optional filetype As String = "par")
        Dim full_artnr, parjson, pyarg As String
        Dim slist As List(Of Integer)

        Try
            full_artnr = GetFullArtnumber(singlepdmid)

            If General.GetFullFilename(General.currentjob.Workspace, full_artnr, filetype) = "" Then
                'get status
                If filetype = "par" Then
                    slist = Database.GetStatusfromDB(full_artnr, "cad_part")
                Else
                    slist = Database.GetStatusfromDB(full_artnr, "cad_assembly")
                End If

                If slist.Count > 0 Then
                    parjson = CreateJSON(full_artnr, orderdir, filetype, slist.Max)

                    pyarg = parjson + " " + workspace
                    Dim p As Process = Process.Start(batfile, pyarg)
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    General.WaitForFile(workspace, full_artnr, filetype, 150)
                End If
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub CheckoutCircs(singlepdmid As String, orderdir As String, workspace As String, batfile As String)
        Dim full_artnr, dftjson1, pyarg As String
        Dim slist As List(Of Integer)

        Try
            full_artnr = GetFullArtnumber(singlepdmid)

            If General.GetFullFilename(General.currentjob.Workspace, full_artnr, "dft") = "" Then
                'get status
                slist = Database.GetStatusfromDB(full_artnr, "cad_drawing")

                If slist.Count > 0 Then
                    dftjson1 = CreateJSON(full_artnr, orderdir, "dft", slist.Max)

                    pyarg = dftjson1 + " " + workspace
                    Dim p As Process = Process.Start(batfile, pyarg)
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    General.WaitForFile(workspace, full_artnr, "dft", 150)
                End If
            Else
                Debug.Print("File exists: " + full_artnr + ".dft")
            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function GetFullArtnumber(pdmid As String) As String
        Dim full_pdmid As String = pdmid

        While full_pdmid.Length < 10
            full_pdmid = "0" + full_pdmid
        End While

        Return full_pdmid
    End Function

    Shared Function CreateJSON(artnr As String, orderdir As String, filetype As String, Optional status As Integer = 0) As String
        Dim filename As String
        Dim jobj As Linq.JObject

        Select Case filetype
            Case "asm"
                jobj = New Linq.JObject(New Linq.JProperty("teilenummer", artnr), New Linq.JProperty("z_art", "cad_assembly"), New Linq.JProperty("z_status", status))
            Case "par"
                jobj = New Linq.JObject(New Linq.JProperty("teilenummer", artnr), New Linq.JProperty("z_art", "cad_part"), New Linq.JProperty("z_status", status))
            Case Else
                jobj = New Linq.JObject(New Linq.JProperty("teilenummer", artnr), New Linq.JProperty("z_art", "cad_drawing"), New Linq.JProperty("z_status", status))
        End Select
        filename = orderdir + "\cond_" + artnr + "_" + filetype + "_" + status.ToString + ".json"

        If Not File.Exists(filename) Then
            Using file As StreamWriter = IO.File.CreateText(filename)
                Using writer As JsonTextWriter = New JsonTextWriter(file)
                    jobj.WriteTo(writer)
                End Using
            End Using
        End If

        Return filename
    End Function

    Shared Function CDB_Material(materialcodeletter As String) As String
        Dim matkey As String = ""

        Select Case materialcodeletter
            Case "C"
                matkey = "C (SP01-A)"
            Case "D"
                matkey = "C (SP01-A1)"
            Case "X"
                matkey = "X (SP01-C)"
            Case "R"
                matkey = "R (SP01-B1)"
            Case "K"
                matkey = "C (SP01-A1)"
            Case "F"
                matkey = "F (SP02-A)"
            Case "G"
                matkey = "AlMg (SP12)"
            Case "S"
                matkey = "St galv (SP15-1)"
            Case "V"
                matkey = "V (SP03-1)"
            Case "W"
                matkey = "W (SP03-2)"
        End Select

        Return matkey
    End Function

    Shared Function GetPDMDescription(keyname As String) As String()
        Dim CDB_de, CDB_en As String

        CDB_de = ""
        CDB_en = ""

        If keyname.ToLower.Contains("header") Then
            If keyname.ToLower.Contains("inlet") Then
                CDB_de = "Verteilrohr"
                CDB_en = "Header Inlet"
            Else
                CDB_de = "Sammelrohr"
                CDB_en = "Header Outlet"
            End If
        ElseIf keyname.ToLower.Contains("nipple") Then
            CDB_de = "Rohrstutzen"
            CDB_en = "Tube connection"
            If keyname.ToLower.Contains("inlet") Then
            End If
        End If

        Return {CDB_de, CDB_en}
    End Function

    Shared Function GetAGPNumber(PDMdescription As String, material As String, onebranch As Boolean) As String
        Dim agpno As String = ""

        'use PDM description from file properties instead of part ID
        Select Case PDMdescription
            Case "Anschlußsystem"
                'use headermaterial for 1 branch (incl AP/CP), but nipple material for 2+
                If onebranch Then
                    If material = "C" Then
                        agpno = "103"
                    Else
                        agpno = "104"
                    End If
                Else
                    If material = "C" Then
                        agpno = "101"
                    ElseIf material = "V" Then
                        agpno = "102"
                    Else
                        agpno = "105"
                    End If
                End If
            Case "Verteilrohr"
                If material = "C" Then
                    agpno = "113"
                Else
                    agpno = "114"
                End If
            Case "Sammelrohr"
                If material = "C" Then
                    agpno = "113"
                Else
                    agpno = "114"
                End If
            Case "Rohrstutzen"
                If material = "V" Or material = "W" Then
                    agpno = "125"
                Else
                    agpno = "123"
                End If
            Case "Stutzen"
                If material = "C" Then
                    agpno = "121"
                Else
                    agpno = "122"
                End If
            Case "Rohrbogen"
                If material = "C" Then
                    agpno = "131"
                Else
                    agpno = "132"
                End If
        End Select

        Return agpno
    End Function

    Shared Sub SaveAsm()
        Dim seenvs As SolidEdgeFramework.Environments
        Dim seassenv As SolidEdgeFramework.Environment
        Dim secmdbars As SolidEdgeFramework.CommandBars
        Dim cntctbar As SolidEdgeFramework.CommandBar
        Dim cmdbarcntrls As SolidEdgeFramework.CommandBarControls
        Dim cntctcntrl As SolidEdgeFramework.CommandBarControl
        Dim windowname As String = "Transfer changes to the PDM system"
        Dim loopexit As Boolean = False
        Dim loopcount As Integer = 1
        Dim mname As String
        Dim wsmprocess As Process
        Dim success As Boolean = False

        Try

            'Get environments
            seenvs = General.seapp.Environments

            'Find Assembly environment
            seassenv = FindEnv(seenvs, "Assembly")

            If seassenv IsNot Nothing Then
                Console.WriteLine("Found env")
                'Get command bars
                secmdbars = seassenv.CommandBars

                'Find command bar
                cntctbar = FindCmdBar(secmdbars)

                If cntctbar IsNot Nothing Then
                    Console.WriteLine("Found contact control bar")
                    'Get the controls of the contact addin
                    cmdbarcntrls = cntctbar.Controls

                    'Find the Save command
                    cntctcntrl = FindCntrl(cmdbarcntrls, "wsmsave")

                    If cntctcntrl IsNot Nothing Then
                        Console.WriteLine("Found wsmsave button")
                        'Execute Save command
                        cntctcntrl.Execute()
                        Threading.Thread.Sleep(3000)

                        'Find Tranfer command
                        cntctcntrl = FindCntrl(cmdbarcntrls, "wsmsaveinpdm")

                        If cntctcntrl IsNot Nothing Then
                            Console.WriteLine("Found save to PDM button")
                            'Execute Transfer command
                            cntctcntrl.Execute()

                            'running in a loop with set max sleep time
                            Console.WriteLine(Date.Now.ToString + " - start save command")

                            'find wsm task and wait until window pops up
                            wsmprocess = General.FindProcess

                            Threading.Thread.Sleep(10000)
                            Do
                                Dim firsthwd As IntPtr = WindowsAPI.FindWindowHandle(windowname, True)
                                If firsthwd <> IntPtr.Zero Then
                                    Console.WriteLine("Window confirmed")
                                    mname = "First Method"
                                Else
                                    Console.WriteLine("2nd search with contain")
                                    firsthwd = WindowsAPI.FindWindowHandle(windowname, False)
                                    mname = "Second Method"
                                End If

                                If firsthwd <> IntPtr.Zero Then
                                    Console.WriteLine("First Handle found")
                                    WindowsAPI.SetForegroundWindow(firsthwd)

                                    For i As Integer = 1 To 7
                                        WindowsAPI.PostMessage(firsthwd, WindowsAPI.Messages.WM_KEYDOWN, WindowsAPI.VKeys.KEY_TAB, 0)
                                        Threading.Thread.Sleep(300)
                                    Next

                                    WindowsAPI.PostMessage(firsthwd, WindowsAPI.Messages.WM_KEYDOWN, WindowsAPI.VKeys.KEY_Return, 0)
                                    Threading.Thread.Sleep(1000)
                                    success = True
                                    loopexit = True
                                Else
                                    Console.WriteLine("loop #" + loopcount.ToString + " searching Commit window")
                                    Threading.Thread.Sleep(5000)
                                    loopcount += 1
                                    If loopcount = 25 Then
                                        loopexit = True
                                    End If
                                    Debug.Print(Date.Now.ToString + " - loopnumber " + loopcount.ToString)
                                End If
                            Loop Until loopexit
                            Debug.Print(Date.Now.ToString + " - loopexit")

                            If success Then
                                General.currentjob.Status = Library.JobStatus.FINISHED
                            Else
                                General.currentjob.Status = Library.JobStatus.FAILED
                            End If

                        Else
                            General.currentjob.Status = Library.JobStatus.NOTRANSFER
                        End If

                    Else
                        General.currentjob.Status = Library.JobStatus.NOSAVECMD
                    End If

                Else
                    General.currentjob.Status = Library.JobStatus.NOADDIN
                End If

            Else
                General.currentjob.Status = Library.JobStatus.NOENV
            End If
        Catch ex As Exception

        End Try

    End Sub

    Shared Function FindEnv(seenvs As SolidEdgeFramework.Environments, checkenvname As String) As SolidEdgeFramework.Environment
        Dim seassenv As SolidEdgeFramework.Environment = Nothing
        Dim currentenv As SolidEdgeFramework.Environment
        Dim envname As String
        Dim envfound As Boolean = False
        Dim i As Integer = 0

        Do
            currentenv = seenvs(i)
            envname = currentenv.Name
            If envname = checkenvname Then
                seassenv = currentenv
                envfound = True
            End If

            i += 1
        Loop Until envfound Or i = seenvs.Count

        Return seassenv
    End Function

    Shared Function FindCmdBar(secmdbars As SolidEdgeFramework.CommandBars) As SolidEdgeFramework.CommandBar
        Dim cntcbar As SolidEdgeFramework.CommandBar = Nothing

        For Each cmdbar As SolidEdgeFramework.CommandBar In secmdbars
            If cmdbar.Name.Contains("CONTACT") Or cmdbar.Name.Contains("Workspace") Then
                cntcbar = cmdbar
            End If
        Next

        Return cntcbar
    End Function

    Shared Function FindCntrl(cmdbarcntrls As SolidEdgeFramework.CommandBarControls, controlname As String) As SolidEdgeFramework.CommandBarControl
        Dim cmdcntrl As SolidEdgeFramework.CommandBarControl = Nothing
        Dim currentcntrl As SolidEdgeFramework.CommandBarControl
        Dim cntrltag As String
        Dim cmdfound As Boolean = False
        Dim i As Integer = 0

        Do
            currentcntrl = cmdbarcntrls(i)
            cntrltag = currentcntrl.Tag
            If cntrltag = controlname Then
                cmdcntrl = currentcntrl
                cmdfound = True
            End If

            i += 1
        Loop Until i = cmdbarcntrls.Count Or cmdfound

        Return cmdcntrl
    End Function

    Shared Function WaitforWSMDialog() As Boolean
        Dim isfinished As Boolean
        Dim checktime, controltime As Date

        checktime = Date.UtcNow
        Dim i As Integer = 1
        Do
            Dim windic As IDictionary(Of IntPtr, String) = WindowsAPI.GetOpenWindows()
            isfinished = True
            Threading.Thread.Sleep(1000)
            Console.WriteLine("Loop #" + i.ToString)
            For Each win In windic
                If win.Value = "Application jobs progress" Then
                    isfinished = False
                    Console.WriteLine("Step application")
                ElseIf win.Value = "Waiting for server..." Then
                    isfinished = False
                    Console.WriteLine("Step waiting")
                ElseIf win.Value = "Transfer files" Then
                    isfinished = False
                    Console.WriteLine("Step transfer")
                ElseIf win.Value.Contains("changes to the PDM system") Then
                    isfinished = False
                End If
            Next
            i += 1
            controltime = Date.UtcNow
        Loop Until isfinished Or (controltime - checktime).TotalSeconds > 180

        Return isfinished
    End Function

    Shared Function SaveDFT() As Boolean
        Dim seenvs As SolidEdgeFramework.Environments
        Dim seassenv As SolidEdgeFramework.Environment
        Dim secmdbars As SolidEdgeFramework.CommandBars
        Dim cntctbar As SolidEdgeFramework.CommandBar
        Dim cmdbarcntrls As SolidEdgeFramework.CommandBarControls
        Dim cntctcntrl As SolidEdgeFramework.CommandBarControl
        Dim wsmprocess As Process
        Dim windowname As String = "Transfer changes to the PDM system"
        Dim loopexit As Boolean = False
        Dim loopcount As Integer = 1
        Dim mname As String
        Dim success As Boolean = False

        Try

            'Get environments
            seenvs = General.seapp.Environments

            'Find Assembly environment
            seassenv = FindEnv(seenvs, "Detail")

            If seassenv IsNot Nothing Then
                Console.WriteLine("Found asm env")
                'Get command bars
                secmdbars = seassenv.CommandBars

                'Find command bar
                cntctbar = FindCmdBar(secmdbars)

                If cntctbar IsNot Nothing Then
                    Console.WriteLine("Found contact control bar")
                    'Get the controls of the contact addin
                    cmdbarcntrls = cntctbar.Controls

                    'Find the Save command
                    cntctcntrl = FindCntrl(cmdbarcntrls, "wsmsave")

                    If cntctcntrl IsNot Nothing Then
                        Console.WriteLine("Found wsmsave button")
                        'Execute Save command
                        cntctcntrl.Execute()
                        Threading.Thread.Sleep(3000)

                        'Find Tranfer command
                        cntctcntrl = FindCntrl(cmdbarcntrls, "wsmsaveinpdm")

                        If cntctcntrl IsNot Nothing Then
                            Console.WriteLine("Found save to PDM button")
                            'Execute Transfer command
                            cntctcntrl.Execute()

                            'running in a loop with set max sleep time
                            Console.WriteLine(Date.Now.ToString + " - start save command")

                            'find WSM task And wait until window pops up
                            wsmprocess = General.FindProcess

                            Threading.Thread.Sleep(10000)
                            Do
                                Dim firsthwd As IntPtr = WindowsAPI.FindWindowHandle(windowname, True)
                                If firsthwd <> IntPtr.Zero Then
                                    Console.WriteLine("Window confirmed")
                                    mname = "First Method"
                                Else
                                    Console.WriteLine("2nd search with contain")
                                    firsthwd = WindowsAPI.FindWindowHandle(windowname, False)
                                    mname = "Second Method"
                                End If

                                If firsthwd <> IntPtr.Zero Then
                                    Console.WriteLine("First Handle found")
                                    WindowsAPI.SetForegroundWindow(firsthwd)

                                    For i As Integer = 1 To 7
                                        WindowsAPI.PostMessage(firsthwd, WindowsAPI.Messages.WM_KEYDOWN, WindowsAPI.VKeys.KEY_TAB, 0)
                                        Threading.Thread.Sleep(300)
                                    Next

                                    WindowsAPI.PostMessage(firsthwd, WindowsAPI.Messages.WM_KEYDOWN, WindowsAPI.VKeys.KEY_Return, 0)
                                    Threading.Thread.Sleep(1000)
                                    success = True
                                    loopexit = True
                                Else
                                    Console.WriteLine("loop #" + loopcount.ToString + " searching Commit window")
                                    Threading.Thread.Sleep(5000)
                                    loopcount += 1
                                    If loopcount = 40 Then
                                        loopexit = True
                                    End If
                                    Debug.Print(Date.Now.ToString + " - loopnumber " + loopcount.ToString)
                                End If
                            Loop Until loopexit
                            Debug.Print(Date.Now.ToString + " - loopexit")
                            If success Then
                                General.currentjob.Status = Library.JobStatus.FINISHED
                            Else
                                General.currentjob.Status = Library.JobStatus.NOTRANSFER
                            End If
                        Else
                            General.currentjob.Status = Library.JobStatus.NOTRANSFER
                        End If

                    Else
                        General.currentjob.Status = Library.JobStatus.NOSAVECMD
                    End If

                Else
                    General.currentjob.Status = Library.JobStatus.NOADDIN
                End If

            Else
                General.currentjob.Status = Library.JobStatus.NOENV
            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
        Return success
    End Function

    Shared Sub SaveCS(coversheetfile As String)
        Dim dftfile As String

        Try
            dftfile = SEDraft.UpdateDrawing(coversheetfile)
            'rename files and update modellink
            coversheetfile = SEPart.RenameCoversheet(coversheetfile)
            SEDraft.RenameCoversheet(dftfile, coversheetfile)
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub KillTask(pname As String)
        Dim WSMtask() As Process
        Dim currenttask, owner As String

        Try
            WSMtask = Process.GetProcesses()

            For Each task As Process In WSMtask
                currenttask = task.ToString()
                If currenttask.Contains(pname) And Not currenttask.Contains("Micro") Then
                    owner = GetProcessOwner(task.Id)
                    If owner = Environment.UserDomainName + "\" + Environment.UserName Then
                        task.Kill()
                        Exit For
                    End If
                End If
            Next
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Function GetProcessOwner(ByVal processId As Integer) As String
        Dim query As String = "Select * From Win32_Process Where ProcessID = " & processId
        Dim searcher As ManagementObjectSearcher = New ManagementObjectSearcher(query)
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

End Class
