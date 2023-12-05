Imports System.Text
Imports System.Runtime.InteropServices


Public Class WindowsAPI
    Delegate Function EnumWindowsProc(ByVal hWnd As IntPtr, ByVal lParam As Integer) As Boolean

    Shared Function FindWindowHandle(windowname As String, isexact As Boolean) As IntPtr
        Dim windowdic As IDictionary(Of IntPtr, String)
        Dim winhwnd As IntPtr = IntPtr.Zero

        Try
            windowdic = GetOpenWindows()

            For Each entry In windowdic
                If isexact Then
                    If entry.Value.ToLower = windowname.ToLower Then
                        winhwnd = entry.Key
                    End If
                Else
                    If entry.Value.ToLower.Contains(windowname.ToLower) Then
                        winhwnd = entry.Key
                    End If
                End If
            Next

        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try

        Return winhwnd
    End Function

    Shared Function GetOpenWindows() As IDictionary(Of IntPtr, String)
        Dim shellWindow As IntPtr = GetShellWindow()
        Dim windows As Dictionary(Of IntPtr, String) = New Dictionary(Of IntPtr, String)()
        EnumWindows(Function(ByVal hWnd As IntPtr, ByVal lParam As Integer)
                        If hWnd = shellWindow Then Return True
                        If Not IsWindowVisible(hWnd) Then Return True
                        Dim length As Integer = GetWindowTextLength(hWnd)
                        Dim builder As StringBuilder = New StringBuilder(length)
                        GetWindowText(hWnd, builder, length + 1)
                        windows(hWnd) = builder.ToString()
                        Return True
                    End Function, 0)
        Return windows
    End Function

    Shared Function GetChildWindows(phWnd As IntPtr) As IDictionary(Of IntPtr, String)
        Dim shellWindow As IntPtr = GetShellWindow()
        Dim windows As Dictionary(Of IntPtr, String) = New Dictionary(Of IntPtr, String)()
        EnumChildWindows(phWnd, Function(ByVal hWnd As IntPtr, ByVal lParam As Integer)
                                    If hWnd = shellWindow Then Return True
                                    If Not IsWindowVisible(hWnd) Then Return True
                                    Dim length As Integer = GetWindowTextLength(hWnd)
                                    Dim builder As StringBuilder = New StringBuilder(length)
                                    GetWindowText(hWnd, builder, length + 1)
                                    windows(hWnd) = builder.ToString()
                                    Return True
                                End Function, 0)
        Return windows
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetClassName(ByVal hWnd As IntPtr, ByVal lpClassName As StringBuilder, ByVal nMaxCount As Integer) As Integer
    End Function
    <DllImport("user32.dll")>
    Shared Function EnumChildWindows(ByVal hwnd As IntPtr, ByVal func As EnumWindowsProc, ByVal lParam As IntPtr) As Boolean
    End Function
    <DllImport("USER32.DLL")>
    Shared Function EnumWindows(ByVal enumFunc As EnumWindowsProc, ByVal lParam As Integer) As Boolean
    End Function
    <DllImport("USER32.DLL")>
    Shared Function GetWindowText(ByVal hWnd As IntPtr, ByVal lpString As StringBuilder, ByVal nMaxCount As Integer) As Integer
    End Function
    <DllImport("USER32.DLL")>
    Shared Function GetWindowTextLength(ByVal hWnd As IntPtr) As Integer
    End Function
    <DllImport("USER32.DLL")>
    Shared Function IsWindowVisible(ByVal hWnd As IntPtr) As Boolean
    End Function
    <DllImport("USER32.DLL")>
    Shared Function GetShellWindow() As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function PostMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function FindWindowEx(ByVal hwndParent As IntPtr, ByVal hwndChildAfter As IntPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As IntPtr
    End Function
    Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As IntPtr) As IntPtr

    Public Enum VKeys
        KEY_TAB = &H9
        KEY_SHIFT = &H10
        KEY_ALT = &H12
        KEY_CAPS = &H14
        KEY_Return = &HD
        KEY_C = &H43
        KEY_D = &H44
        KEY_E = &H45
        KEY_G = &H47
        KEY_M = &H4D
        KEY_N = &H4E
        KEY_P = &H50
        KEY_R = &H52
        KEY_S = &H53
        KEY_T = &H54
        KEY_1 = &H31
        KEY_2 = &H32
        KEY_PLUS = &H6B
        KEY_MINUS = &H6D
    End Enum

    Public Enum Messages
        WM_KEYDOWN = &H100
        WM_KEYUP = &H101
        WM_CHAR = &H102
    End Enum

End Class
