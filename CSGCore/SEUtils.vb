Imports SEFramework = SolidEdgeFramework
Imports System.Runtime.InteropServices
Imports System.Management

Public Class SEUtils

    Public Shared Function Connect() As SEFramework.Application
        Dim processes() As Process = Process.GetProcesses
        Dim owner, progID As String
        Dim SEApplication As SEFramework.Application = Nothing

        Try
            progID = ""
            For Each proc In processes
                If proc.ProcessName = "Edge" Then
                    owner = GetProcessOwner(proc.Id)
                    If owner = Environment.UserDomainName + "\" + Environment.UserName Then
                        progID = proc.Id
                    End If
                End If
            Next

            If progID <> "" Then
                Console.WriteLine("SE already running, ID = " + progID)
                SEApplication = DirectCast(Marshal.GetActiveObject("SolidEdge.Application"), SEFramework.Application)
            Else
                Console.WriteLine("Starting SE")
                StartSE("C:\Program Files\Siemens\Solid Edge 2021\Program\Edge.exe")
            End If

        Catch ex As COMException

        End Try

        Return SEApplication

    End Function

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

    Public Shared Function Start() As SEFramework.Application

        ' On a system where Solid Edge is installed, the COM ProgID will be
        ' defined in registry: HKEY_CLASSES_ROOT\SolidEdge.Application
        Dim t As Type = Type.GetTypeFromProgID(progID:="SolidEdge.Application", throwOnError:=True)

        ' Using the discovered Type, create and return a new instance of Solid Edge.
        Return TryCast(Activator.CreateInstance(type:=t), SEFramework.Application)

    End Function

    Public Shared Sub SwitchWindow(seapp As SEFramework.Application, caption As String, env As String)

        Try
            Dim objwindows As SEFramework.Windows
            Dim asmwindow As SEFramework.Window
            Dim dftwindow As SolidEdgeDraft.SheetWindow

            objwindows = seapp.Windows

            For i As Integer = 0 To objwindows.Count - 1
                If env = "Assembly" Then
                    asmwindow = TryCast(objwindows(i), SEFramework.Window)
                    If asmwindow IsNot Nothing Then
                        If asmwindow.Caption.Contains(caption) Then
                            asmwindow.Activate()
                            seapp.DoIdle()
                            Exit For
                        End If
                    End If
                Else
                    dftwindow = TryCast(objwindows(i), SolidEdgeDraft.SheetWindow)
                    If dftwindow IsNot Nothing Then
                        If dftwindow.Caption.Contains(caption) Then
                            dftwindow.Activate()
                            seapp.DoIdle()
                            Exit For
                        End If
                    End If
                End If
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Public Shared Sub StartSE(ProcessPath)
        Dim objProcess As Process

        Try
            objProcess = New Process()
            objProcess.StartInfo.FileName = ProcessPath
            objProcess.Start()
            'Wait until the process passes back an exit code 
            objProcess.WaitForInputIdle()

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Public Shared Sub CloseWindows()
        Dim sewindows As SEFramework.Windows

        Try
            sewindows = General.seapp.Windows
            For Each sewin As SEFramework.Window In sewindows
                sewin.Close()
                General.seapp.DoIdle()
            Next

        Catch ex As Exception
            Console.WriteLine()
        End Try

    End Sub

    Public Shared Function HandleProcess(starttime As Date) As String
        Dim setasks() As Process
        Dim currenttask As String = ""
        Dim nowtime As Date = Date.UtcNow
        Try
            setasks = Process.GetProcessesByName("Edge")
            currenttask = setasks(0).MainWindowTitle

            Dim delta As TimeSpan = nowtime - starttime
            If delta.TotalSeconds > 35 Then
                currenttask = "stop"
            End If
        Catch ex As Exception

        End Try
        Return currenttask
    End Function

    Shared Function ReConnect() As SEFramework.Application
        Dim SEwindowname As String
        Dim seapp As SEFramework.Application

        'waiting for SE to start
        Dim starttime As Date = Date.UtcNow
        Do
            SEwindowname = HandleProcess(starttime)
        Loop Until SEwindowname <> ""

        seapp = Connect()
        Console.WriteLine("SE Windowname: " + SEwindowname)

        Return seapp
    End Function

    Shared Function FindProp(objgsets As SEFramework.Properties, propname As String) As Integer
        Dim propindex As Integer = -1

        For i As Integer = 1 To objgsets.Count
            Dim currentpropname As String = objgsets.Item(i).Name
            If currentpropname = propname Then
                propindex = i
                Exit For
            End If
        Next

        Return propindex
    End Function

    Shared Function ChangeProp(ByRef objgsets As SEFramework.Properties, index As Integer, propname As String, propvalue As String) As SEFramework.Property
        Dim objprop As SEFramework.Property

        If index < 1 Then
            objprop = objgsets.Add(propname, propvalue)
        Else
            objprop = objgsets.Item(index)
            objprop.Value = propvalue
        End If

        Return objprop
    End Function

End Class
