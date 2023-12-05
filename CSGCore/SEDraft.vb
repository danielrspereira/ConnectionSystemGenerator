Imports SED = SolidEdgeDraft
Imports SEFS = SolidEdgeFrameworkSupport
Imports System.IO
Imports SolidEdgeGeometry
Imports Microsoft.SqlServer.Server

Public Class SEDraft

    Shared Function OpenDFT(filename As String, Optional loopnumber As Integer = 1) As SED.DraftDocument
        Dim dftdoc As SED.DraftDocument = Nothing
        Dim windowdic As IDictionary(Of IntPtr, String) = WindowsAPI.GetOpenWindows()
        Dim childdic As IDictionary(Of IntPtr, String)

        Try
            dftdoc = General.seapp.Documents.Open(filename)
            General.seapp.DoIdle()
            Threading.Thread.Sleep(5000)
            For Each win In windowdic
                If win.Value = "Solid Edge" Then
                    childdic = WindowsAPI.GetChildWindows(win.Key)
                    For Each chil In childdic
                        If chil.Value = "OK" Then
                            WindowsAPI.SetForegroundWindow(chil.Key)
                            WindowsAPI.PostMessage(chil.Key, WindowsAPI.Messages.WM_KEYDOWN, WindowsAPI.VKeys.KEY_Return, 0)
                        End If
                    Next
                    Exit For
                End If
            Next
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
            Threading.Thread.Sleep(2000)
            If filename <> "" Then
                General.StartSE()
                SEUtils.ReConnect()
                If loopnumber < 3 Then
                    OpenDFT(filename, loopnumber + 1)
                End If
            End If
        End Try
        Return dftdoc
    End Function

    Shared Sub FitWindow()
        Try
            Dim sewindow As SED.SheetWindow = General.seapp.ActiveWindow
            sewindow.Fit()
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Function GetCoilPosition(objsheet As SED.Sheet) As String
        Dim coilposition = ""
        Dim objgroups As SEFS.Groups
        Dim objgroup As SEFS.Group
        Dim objlines As SEFS.Lines2d
        Dim objline As SEFS.Line2d
        Dim xs, ys, xe, ye, crossingpoint() As Double
        Dim i As Integer
        Dim xslist, xelist, yslist, yelist As New List(Of Double)

        Try
            objgroups = objsheet.Groups
            i = 0
            Do
                objgroup = objgroups(i)
                If objgroup.UserDefinedName = "LR_AD" Then
                    'Using the longest line of the LR_AD group to get the air direction
                    objlines = objgroup.Lines2d
                    For Each objline In objlines
                        objline.GetStartPoint(xs, ys)
                        objline.GetEndPoint(xe, ye)
                        'pick the diagonal lines
                        If xs <> xe And ys <> ye Then
                            xslist.Add(Math.Round(xs, 6))
                            xelist.Add(Math.Round(xe, 6))
                            yslist.Add(Math.Round(ys, 6))
                            yelist.Add(Math.Round(ye, 6))
                        End If
                    Next

                    crossingpoint = CircProps.ContactPoint(xslist(0), yslist(0), xslist(1), yslist(1), xelist(0), yelist(0), xelist(1), yelist(1))

                    If crossingpoint(1) < Math.Max(yslist.Max, yelist.Max) And crossingpoint(1) > Math.Min(yslist.Min, yelist.Min) Then 'And crossingpoint(0) = Math.Max(xslist.Max, xelist.Max)
                        'coil vertical, airflow horizontal
                        coilposition = "vertical"
                    Else
                        'coil horizontal, airflow vertical
                        coilposition = "horizontal"
                    End If

                End If
                i += 1
            Loop Until i >= objgroups.Count Or coilposition <> ""

        Catch ex As Exception
            Debug.Print("Error getting coil position")
        End Try

        Return coilposition
    End Function

    Shared Function GetCoilFrame(objsheet As SED.Sheet, coilposition As String, mcd As String, coilnumber As Integer) As Double()
        Dim objgroups As SEFS.Groups
        Dim objgroup As SEFS.Group
        Dim objlines As SEFS.Lines2d
        Dim objline As SEFS.Line2d
        Dim xslist, yslist, xelist, yelist As New List(Of Double)
        Dim xs, ys, xe, ye, xmax, xmin, ymax, coilframe(), rangelimit As Double
        Dim i As Integer = 0
        Dim issplit As Boolean = False

        Try
            objgroups = objsheet.Groups
            xmax = 0

            If mcd <> "2" Or coilnumber = 1 Then
                rangelimit = GetRangeLimit(objsheet, coilposition)
            End If

            If mcd = "2" AndAlso GetRangeLimit(objsheet, coilposition) > 0 Then
                'check for split circuiting
                issplit = True
            End If

            Do
                objgroup = objgroups(i)
                If objgroup.UserDefinedName = "coil frame" Or objgroup.UserDefinedName = "BlockRahmen" Then
                    objlines = objgroup.Lines2d
                    'Origin is always at 0/0 → top right point of the frame = size
                    For Each objline In objlines
                        objline.GetStartPoint(xs, ys)
                        objline.GetEndPoint(xe, ye)
                        xs = Math.Round(xs * 1000, 3)
                        ys = Math.Round(ye * 1000, 3)
                        xe = Math.Round(xe * 1000, 3)
                        ye = Math.Round(ye * 1000, 3)

                        xslist.Add(xs)
                        yslist.Add(ys)
                        xelist.Add(xe)
                        yelist.Add(ye)

                        xmax = Math.Max(Math.Abs(xslist.Max), Math.Abs(xelist.Max))
                        xmin = Math.Min(Math.Abs(xslist.Max), Math.Abs(xelist.Min))
                        If xmin <> 0 Then
                            xmax -= xmin
                        End If

                        ymax = Math.Max(yslist.Max, yelist.Max)
                        If rangelimit <> 0 Or issplit Then
                            If coilposition = "horizontal" Then
                                ymax /= 2
                            Else
                                xmax /= 2
                            End If
                        End If
                    Next
                    coilframe = {xmax, ymax}
                End If
                i += 1
            Loop Until i >= objgroups.Count Or xmax > 0

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        coilframe = {xmax, ymax}

        Return coilframe
    End Function

    Shared Function GetCTPositions(objsheet As SED.Sheet) As List(Of Double)()
        Dim objcircles As SEFS.Circles2d
        Dim objcircle As SEFS.Circle2d
        Dim diameterlist As New List(Of Double) From {7, 9.52, 9.525, 12, 15, 22}
        Dim circdiameter, xc, yc As Double
        Dim xlist, ylist As New List(Of Double)
        Dim coordlist() As List(Of Double)

        Try
            objcircles = objsheet.Circles2d
            For Each objcircle In objcircles
                Dim objstyle As SolidEdgeFrameworkSupport.GeometryStyle2d = objcircle.Style

                If objcircle.Layer = "Default" And objstyle.DashGapCount = 0 Then
                    circdiameter = objcircle.Diameter
                    circdiameter = Math.Round(circdiameter * 1000, 3)
                    For Each diameter In diameterlist
                        If circdiameter = diameter Then
                            'Get center point
                            objcircle.GetCenterPoint(xc, yc)
                            xc = Math.Round(xc * 1000, 3)
                            yc = Math.Round(yc * 1000, 3)
                            xlist.Add(xc)
                            ylist.Add(yc)
                        End If
                    Next
                End If
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        coordlist = {xlist, ylist}

        Return coordlist
    End Function

    Shared Sub CreateMirroredDFT(draftdoc As SED.DraftDocument, usemirror As Boolean, conside As String, alignment As String, customcirc As Boolean)
        Dim objDV As SED.DrawingView
        Dim x0, y0 As Double

        Try
            objDV = FindDV(draftdoc)

            'move it out of the focus
            FitWindow()

            objDV.GetOrigin(x0, y0)
            objDV.SetOrigin(x0 + 2, y0)

            'check if mirror needed
            If Not customcirc Then
                If MirrorNeeded(objDV.Sheet) Then
                    MirrorDV(objDV, usemirror)
                ElseIf General.currentunit.UnitDescription <> "VShape" And General.currentunit.ApplicationType = "Condenser" And ControlCTonDV(objDV.Sheet, alignment, conside) Then
                    '((conside = "left" And ControlCTonDV(objDV.Sheet, alignment)) Or (conside = "right" And Not ControlCTonDV(objDV.Sheet, alignment))) Then
                    MirrorDV(objDV, True)
                End If
            End If

            draftdoc.Application.ScreenUpdating = True

            DirectCast(draftdoc.Application.ActiveWindow, SED.SheetWindow).Update()

            objDV.SetOrigin(x0, y0)

            draftdoc.Save()
            General.seapp.DoIdle()
            'draftdoc.Close()
            'General.seapp.DoIdle()
            'General.ReleaseObject(draftdoc)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Function ControlCTonDV(objsheet1 As SED.Sheet, alignment As String, conside As String) As Boolean
        Dim needmirror As Boolean = False

        If alignment = "horizontal" Then
            If ObjectAtLoc(objsheet1, {Math.Round(25 / 2, 2), Math.Round(25 / 2, 2)}, "circle") And conside = "right" Then
                needmirror = True
            End If
        Else
            If Not ObjectAtLoc(objsheet1, {Math.Round(25 * 3 / 2, 2), Math.Round(25 / 2, 2)}, "circle") And conside = "left" Then
                needmirror = True
            End If
        End If

        Return needmirror
    End Function

    Shared Function MirrorNeeded(DVSheet As SED.Sheet) As Boolean
        Dim DVGroup As SEFS.Group
        Dim xlist As New List(Of Double)
        Dim needed As Boolean = False

        Try
            DVGroup = GetGroupByName(DVSheet, "coil frame")
            If DVGroup Is Nothing Then
                DVGroup = GetGroupByName(DVSheet, "BlockRahmen")
            End If

            Dim x1 As Double, y1 As Double
            Dim x2 As Double, y2 As Double

            If DVGroup IsNot Nothing Then
                For Each l As SEFS.Line2d In DVGroup.Lines2d
                    l.GetStartPoint(x1, y1)
                    l.GetEndPoint(x2, y2)
                    xlist.Add(Math.Min(x1, x2))
                Next

                If xlist.Min < -0.01 Then
                    needed = True
                End If
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return needed
    End Function

    Shared Sub MirrorDV(objDV As SED.DrawingView, usemirror As Boolean)
        Dim objsheet As SED.Sheet
        Dim framewidth As Double

        objsheet = objDV.Sheet
        framewidth = GetWidthFromGroup(objsheet, "coil frame")
        If framewidth = 0 Then
            framewidth = GetWidthFromGroup(objsheet, "BlockRahmen")
        End If
        MirrorObjects(objsheet, framewidth, usemirror)

    End Sub

    Shared Function GetWidthFromGroup(DVSheet As SED.Sheet, groupname As String) As Double
        Dim DVGroup As SEFS.Group
        Dim GroupWidth As Double

        DVGroup = GetGroupByName(DVSheet, groupname)
        If DVGroup IsNot Nothing Then
            Dim x1 As Double, y1 As Double
            Dim x2 As Double, y2 As Double
            For Each l As SEFS.Line2d In DVGroup.Lines2d
                l.GetStartPoint(x1, y1)
                l.GetEndPoint(x2, y2)
                If Math.Abs(x1 - x2) > GroupWidth Then GroupWidth = Math.Round(Math.Abs(x1 - x2), 5)
            Next
        Else
            GroupWidth = 0
        End If

        Return GroupWidth
    End Function

    Shared Function GetGroupByName(oSheet As SED.Sheet, name As String) As SEFS.Group
        Dim DVGroups As SEFS.Groups = oSheet.Groups

        For g As Integer = 1 To DVGroups.Count
            If DVGroups.Item(g).UserDefinedName = name Then
                Return DVGroups.Item(g)
            End If
        Next
        Return Nothing
    End Function

    Shared Sub MirrorObjects(objSheet As SED.Sheet, framewidth As Double, usemirror As Boolean)

        For Each lineelement As SEFS.Line2d In objSheet.Lines2d
            If usemirror Then
                lineelement.Mirror(0, 0, 0, 1)
            End If
            lineelement.Move(0, 0, framewidth, 0)
        Next
        For Each circelement As SEFS.Circle2d In objSheet.Circles2d
            If usemirror Then
                circelement.Mirror(0, 0, 0, 1)
            End If
            circelement.Move(0, 0, framewidth, 0)
        Next
        For Each tbelement As SEFS.TextBox In objSheet.TextBoxes
            If usemirror Then
                tbelement.Mirror(0, 0, 0, 1)
            End If
            tbelement.Move(0, 0, framewidth, 0)
        Next
        For Each groupelement As SEFS.Group In objSheet.Groups
            If usemirror Then
                groupelement.Mirror(0, 0, 0, 1)
            End If
            groupelement.Move(0, 0, framewidth, 0)
        Next

    End Sub

    Shared Function FindDV(dftdoc As SED.DraftDocument) As SED.DrawingView
        Dim objDV As SED.DrawingView = Nothing
        Dim DVList As New List(Of SED.DrawingView)
        Dim objSheet As SED.Sheet
        Dim sheetfound As Boolean = False
        Dim framewidth As Double

        Try
            For Each DV As SED.DrawingView In dftdoc.ActiveSheet.DrawingViews
                framewidth = GetWidthFromGroup(DV.Sheet, "coil frame")
                If framewidth = 0 Then
                    framewidth = GetWidthFromGroup(DV.Sheet, "BlockRahmen")
                End If
                If framewidth <> 0 Then
                    objSheet = DV.Sheet
                    DVList.Add(DV)
                End If
            Next

            'Usually the key equals the name for normal drawing views
            For i As Integer = 0 To DVList.Count - 1
                objSheet = DVList(i).Sheet
                If objSheet.Name = objSheet.Key Then
                    objDV = DVList(i)
                    sheetfound = True
                    Exit For
                End If
            Next

            'Sometimes the first way doesn't work, so the first drawing view will be used no matter what
            If sheetfound = False Then
                objDV = DVList.First
                General.seapp.DoIdle()
                Threading.Thread.Sleep(2000)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return objDV
    End Function

    Shared Function FindSheet(dftdoc As SED.DraftDocument) As SED.Sheet
        Dim objDV As SED.DrawingView
        Dim DVList As New List(Of SED.DrawingView)
        Dim objSheet As SED.Sheet = Nothing
        Dim sheetfound As Boolean = False
        Dim framewidth As Double

        Try
            For Each DV As SED.DrawingView In dftdoc.ActiveSheet.DrawingViews
                framewidth = GetWidthFromGroup(DV.Sheet, "coil frame")
                If framewidth = 0 Then
                    framewidth = GetWidthFromGroup(DV.Sheet, "BlockRahmen")
                End If
                If framewidth <> 0 Then
                    objSheet = DV.Sheet
                    DVList.Add(DV)
                End If
            Next

            'Usually the key equals the name for normal drawing views
            For i As Integer = 0 To DVList.Count - 1
                objSheet = DVList(i).Sheet
                If objSheet.Name = objSheet.Key Then
                    DVList(i).Sheet.Activate()
                    General.seapp.DoIdle()
                    objSheet = dftdoc.ActiveSheet
                    Threading.Thread.Sleep(2000)
                    sheetfound = True
                    Exit For
                End If
            Next

            'Sometimes the first way doesn't work, so the first drawing view will be used no matter what
            If sheetfound = False Then
                objDV = DVList.First
                objSheet = objDV.Sheet
                objSheet.Activate()
                General.seapp.DoIdle()
                Threading.Thread.Sleep(2000)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return objSheet
    End Function

    Shared Function FindSheetX(side As String, pitchx As Double, pitchy As Double) As SED.Sheet
        Dim dftdoc As SED.DraftDocument
        Dim mainsheet, objsheet1 As SED.Sheet
        Dim objDVList As New List(Of SED.DrawingView)
        Dim objsheet As SED.Sheet = Nothing
        Dim objDVA As SED.DrawingView = Nothing
        Dim needmirror As Boolean
        Dim xlist, ylist, sflist As New List(Of Double)
        Dim framewidth, xmin, ymin, xmax, ymax, dx, dy, yoffset As Double
        Dim coilside As String = ""

        Try
            dftdoc = General.seapp.ActiveDocument

            mainsheet = dftdoc.ActiveSheet

            'first identify all 6 drawing views by scaling factor, must be bigger than the examples for alignment
            For Each DV As SED.DrawingView In mainsheet.DrawingViews
                framewidth = GetWidthFromGroup(DV.Sheet, "coil frame")
                If framewidth = 0 Then
                    framewidth = GetWidthFromGroup(DV.Sheet, "BlockRahmen")
                End If
                If framewidth > 0 Then
                    sflist.Add(Math.Round(DV.ScaleFactor, 5))
                End If
            Next

            For Each DV As SED.DrawingView In mainsheet.DrawingViews
                If sflist.Max = Math.Round(DV.ScaleFactor, 5) Then
                    Dim x0, y0 As Double
                    DV.GetOrigin(x0, y0)
                    xlist.Add(Math.Round(x0 * 1000))
                    ylist.Add(Math.Round(y0 * 1000))
                    objDVList.Add(DV)
                End If
            Next

            xlist.Sort()
            ylist.Sort()
            objDVList.First.Range(xmin, ymin, xmax, ymax)
            dx = Math.Round(Math.Abs(xmax - xmin) * 1000)
            dy = Math.Round(Math.Abs(ymax - ymin) * 1000)

            If General.currentunit.UnitSize = "Compact" Then
                'add 25mm offset for y when check location
                yoffset = 25
            End If

            If xlist.Count > 2 Then
                For i As Integer = 0 To objDVList.Count - 1
                    Dim x0, y0 As Double
                    objDVList(i).GetOrigin(x0, y0)
                    x0 = Math.Round(x0 * 1000)
                    y0 = Math.Round(y0 * 1000)
                    If dx > dy Then
                        'vertically stacked
                        Debug.Print(objDVList(i).Key.ToString + ": y=" + y0.ToString)
                        If (y0 = ylist(1)) And side = "left" Then
                            'probably A
                            objDVA = objDVList(i)
                            coilside = "A"
                        ElseIf y0 = ylist(ylist.Count - 2) And side = "right" Then
                            'probably B
                            objDVA = objDVList(i)
                            coilside = "B"
                        End If
                    Else
                        'horizontally stacked
                        Debug.Print(objDVList(i).Key.ToString + ": x=" + x0.ToString)
                        If (x0 = xlist(1) And side = "left") Or (x0 = xlist(xlist.Count - 2) And side = "right") Then
                            'probably A
                            objDVA = objDVList(i)
                            coilside = "A"
                        End If
                    End If
                Next

                'get AD icon position, left = side A // right = side B 
                'only checking one neccessary, other DV has the opposite
                If objDVA IsNot Nothing Then
                    objsheet1 = objDVA.Sheet
                    objsheet1.Activate()

                    'check core tube position, bottom left = side A
                    If ObjectAtLoc(objsheet1, {Math.Round(pitchx / 2, 2), Math.Round(pitchy / 2 + yoffset, 2)}, "circle") Then
                        'side A
                        If (side = "right" And General.currentunit.UnitSize <> "Compact") Or (General.currentunit.UnitSize = "Compact" And side = "left") Then
                            needmirror = True
                        End If
                    ElseIf ObjectAtLoc(objsheet1, {Math.Round(pitchx * 3 / 2, 2, MidpointRounding.AwayFromZero), Math.Round(pitchy / 2 + yoffset, 2)}, "circle") Then
                        If side = "right" And coilside = "B" Then
                            needmirror = True
                        End If
                    Else
                        If side = "left" Then
                            needmirror = True
                        End If
                    End If

                    If needmirror Then
                        'mirror elements
                        framewidth = GetWidthFromGroup(objsheet1, "coil frame")
                        If framewidth = 0 Then
                            framewidth = GetWidthFromGroup(objsheet1, "BlockRahmen")
                        End If
                        MirrorObjects(objsheet1, framewidth, True)
                    End If

                    objsheet = objDVA.Sheet
                End If
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return objsheet
    End Function

    Shared Function ObjectAtLoc(objsheet As SED.Sheet, location() As Double, objtype As String) As Boolean
        Dim posmatch As Boolean = False
        Try
            If objtype = "circle" Then
                For Each objcircle As SEFS.Circle2d In objsheet.Circles2d
                    If objcircle.Style.DashGapCount = 0 Then
                        Dim x, y As Double
                        objcircle.GetCenterPoint(x, y)
                        x = Math.Round(x * 1000, 3)
                        y = Math.Round(y * 1000, 3)
                        If x = location(0) And y = location(1) Then
                            posmatch = True
                            Exit For
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
        Return posmatch
    End Function

    Shared Sub VShapeMirrorDV(objsheet As SED.Sheet, connectionside As String)
        Dim framewidth As Double
        Dim needmirror As Boolean
        'check core tube position, bottom left = top right = ok
        If ObjectAtLoc(objsheet, {12.5, 12.5}, "circle") Then
            If connectionside = "left" Then
                needmirror = True
            End If
        Else
            If connectionside = "right" Then
                needmirror = True
            End If
        End If

        If needmirror Then
            'mirror elements
            framewidth = GetWidthFromGroup(objsheet, "coil frame")
            If framewidth = 0 Then
                framewidth = GetWidthFromGroup(objsheet, "BlockRahmen")
            End If
            MirrorObjects(objsheet, framewidth, True)
        End If
    End Sub

    Shared Function HeatingPos(objsheet As SED.Sheet) As List(Of Double)()
        Dim angle As Double
        Dim xs, xe, ys, ye, xp, yp As Double
        Dim xplist, yplist As New List(Of Double)
        Dim v2 As Integer

        For Each objgroup As SEFS.Group In objsheet.Groups
            If objgroup.Layer.ToLower.Contains("element") Then
                If objgroup.Lines2d.Count = 3 Then
                    Dim xlist, ylist As New List(Of Double)
                    'one is horizontal (not important), two are diagonal
                    For Each objline As SEFS.Line2d In objgroup.Lines2d
                        angle = Math.Round(objline.Angle, 4)
                        If angle <> 0 And Math.Round(Math.PI - angle, 2) <> 0 Then
                            objline.GetStartPoint(xs, ys)
                            objline.GetEndPoint(xe, ye)
                            xs = Math.Round(xs * 1000, 3)
                            xe = Math.Round(xe * 1000, 3)
                            ys = Math.Round(ys * 1000, 3)
                            ye = Math.Round(ye * 1000, 3)
                            If xlist.IndexOf(xs) > -1 Then
                                xp = xs
                                v2 = General.IntegerRem(Math.Max(ye, ys), 12.5)
                                yp = v2 * 12.5
                                xplist.Add(xp)

                                yplist.Add(yp)
                                Exit For
                            Else
                                xlist.Add(xs)
                                If xlist.IndexOf(xe) > -1 Then
                                    xp = xe
                                    v2 = General.IntegerRem(Math.Max(ye, ys), 12.5)
                                    yp = v2 * 12.5
                                    xplist.Add(xp)
                                    yplist.Add(yp)
                                    Exit For
                                Else
                                    xlist.Add(xe)
                                End If

                            End If
                        End If
                    Next
                ElseIf objgroup.Circles2d.Count = 1 And objgroup.UserDefinedName.ToLower.Contains("sensor") Then
                    objgroup.Circles2d.Item(1).GetCenterPoint(xp, yp)
                    xplist.Add(Math.Round(xp * 1000, 3))
                    yplist.Add(Math.Round(yp * 1000, 3))
                End If

            End If
        Next

        Return {xplist, yplist, xplist, yplist}
    End Function

    Shared Function HeatingBows(objsheet As SED.Sheet, heatingpos As List(Of Double)()) As List(Of Double)()
        Dim xs, ys, xe, ye As Double
        Dim xslist, xelist, yslist, yelist, x1list, y1list, x2list, y2list As New List(Of Double)

        Try
            For Each objgroup As SEFS.Group In objsheet.Groups
                If objgroup.UserDefinedName.Contains("heating") OrElse objgroup.UserDefinedName.Contains("Heizbügel") Then
                    For Each objarc As SEFS.Arc2d In objgroup.Arcs2d
                        If objarc.Style.LinearColor = 210 Then
                            objarc.GetStartPoint(xs, ys)
                            objarc.GetEndPoint(xe, ye)
                            xslist.Add(Math.Round(xs * 1000, 3))
                            xelist.Add(Math.Round(xe * 1000, 3))
                            yslist.Add(Math.Round(ys * 1000, 3))
                            yelist.Add(Math.Round(ye * 1000, 3))
                        End If
                    Next
                End If
            Next

            For i As Integer = 0 To xslist.Count - 1
                For j As Integer = 0 To heatingpos(0).Count - 1
                    Dim dxs, dys As Double
                    dxs = Math.Round(Math.Abs(xslist(i) - heatingpos(0)(j)), 3)
                    dys = Math.Round(Math.Abs(yslist(i) - heatingpos(1)(j)), 3)
                    If Math.Sqrt(dxs ^ 2 + dys ^ 2) < 15 Then
                        For k As Integer = 0 To heatingpos(0).Count - 1
                            Dim dxe, dye As Double
                            dxe = Math.Round(Math.Abs(xelist(i) - heatingpos(0)(k)), 3)
                            dye = Math.Round(Math.Abs(yelist(i) - heatingpos(1)(k)), 3)
                            If Math.Sqrt(dxe ^ 2 + dye ^ 2) < 15 Then
                                x1list.Add(heatingpos(0)(j))
                                y1list.Add(heatingpos(1)(j))
                                x2list.Add(heatingpos(0)(k))
                                y2list.Add(heatingpos(1)(k))
                                Exit For
                            End If
                        Next

                        Exit For
                    End If
                Next
            Next


        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return {x1list, y1list, x2list, y2list}
    End Function

    Shared Function GetBowLines(objsheet As SED.Sheet, coilposition As String, mcd As String, coilnumber As Integer) As List(Of SEFS.Line2d)()
        Dim objlines As SEFS.Lines2d
        Dim objline As SEFS.Line2d
        Dim bowlines() As List(Of SEFS.Line2d) = Nothing
        Dim frontlines, backlines As New List(Of SEFS.Line2d)
        Dim layername As String
        Dim rangelimit, x, y As Double
        Dim checkposition As Boolean = False
        Dim addline As Boolean

        Try
            objlines = objsheet.Lines2d
            'First check if fin is split
            rangelimit = GetRangeLimit(objsheet, coilposition)

            If rangelimit <> 0 Then
                checkposition = True
            End If

            For Each objline In objlines
                layername = objline.Layer
                'Check for the correct layers
                If layername.Contains("front") Or layername.Contains("Vorne") Or layername.Contains("back") Or layername.Contains("Hinten") Then
                    If checkposition Then
                        'Check if line is in the important area
                        objline.GetStartPoint(x, y)
                        addline = False
                        If coilposition = "horizontal" Then
                            If coilnumber = 1 Then
                                If y > rangelimit Then
                                    addline = True
                                End If
                            Else
                                If y < rangelimit Then
                                    addline = True
                                End If
                            End If
                        Else
                            If coilnumber = 1 Then
                                If x > rangelimit Then
                                    addline = True
                                End If
                            Else
                                If x < rangelimit Then
                                    addline = True
                                End If
                            End If
                        End If
                    Else
                        addline = True
                    End If
                    If addline Then
                        If layername.Contains("front") Or layername.Contains("Vorne") Then
                            frontlines.Add(objline)
                        Else
                            backlines.Add(objline)
                        End If
                    End If
                End If
            Next
            bowlines = {frontlines, backlines}
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return bowlines
    End Function

    Shared Function GetRangeLimit(objsheet As SED.Sheet, coilposition As String) As Double
        Dim objlines As SEFS.Lines2d
        Dim objline As SEFS.Line2d
        Dim rangelimit As Double = 0
        Dim x, y As Double

        Try
            objlines = objsheet.Lines2d
            'Search for a dash line
            For Each objline In objlines
                If objline.Style.DashGapCount > 0 Then
                    objline.GetStartPoint(x, y)
                    If coilposition = "vertical" Then
                        'Only left part important
                        rangelimit = x
                    Else
                        'Only upper part important
                        rangelimit = y
                    End If
                End If
            Next
        Catch ex As Exception

        End Try

        Return Math.Round(rangelimit, 7)
    End Function

    Shared Function GetBowProps(objsheet As SED.Sheet, bowlines As List(Of SEFS.Line2d), pitchx As Double, pitchy As Double, coilposition As String,
                                 connectionside As String, circtype As String, passnumber As Integer, hasEDefrost As String, mcd As String, coilnumber As Integer) As List(Of Double)()
        Dim bowline As SEFS.Line2d
        Dim lineone, linezero As New List(Of SEFS.Line2d)
        Dim bowprops() As List(Of Double)
        Dim inoutlist() As List(Of Double)
        Dim x1list, y1list, x2list, y2list, lengthlist, typelist, levellist As New List(Of Double)
        Dim xs, ys, xe, ye, length, bowpoints(), rangelimit, xoffset, yoffset, circframe() As Double
        Dim gcstart, gcend As Integer

        Try
            If circtype = "defrost" Then
                xoffset = 25
                yoffset = 25
            Else
                xoffset = 0
                yoffset = 0
            End If
            'Get the position // list only contains correct lines anyway
            For Each bowline In bowlines
                bowline.GetStartPoint(xs, ys)
                bowline.GetEndPoint(xe, ye)
                xs = Math.Round(xs * 1000, 4)
                ys = Math.Round(ys * 1000, 4)
                xe = Math.Round(xe * 1000, 4)
                ye = Math.Round(ye * 1000, 4)

                'Check how many coords are in the grid (should be 0/2/4)
                gcstart = CircProps.CheckCoords({xs, ys}, pitchx, pitchy, xoffset, yoffset)
                gcend = CircProps.CheckCoords({xe, ye}, pitchx, pitchy, xoffset, yoffset)
                Select Case gcstart + gcend
                    Case 4
                        'Bow is perfectly in the grid
                        x1list.Add(xs)
                        y1list.Add(ys)
                        x2list.Add(xe)
                        y2list.Add(ye)
                        length = Math.Round(bowline.Length * 1000, 3)
                        lengthlist.Add(length)
                        typelist.Add(1)
                        levellist.Add(1)
                    Case 3
                        'outbreak line for diagonal cropped
                        lineone.Add(bowline)
                    Case 2
                        If gcstart = gcend Then
                            'diagonal cropped bowline
                            linezero.Add(bowline)
                        Else
                            'Start or endpoint of an outbreak line
                            lineone.Add(bowline)
                        End If
                    Case 0
                        'line of the cropped bow
                        linezero.Add(bowline)
                End Select
            Next

            If lineone.Count = 2 * linezero.Count Then

                'Find the 2 outbreak lines for the bow
                For Each bowline In linezero
                    'Get the start and endpoint of the bow
                    bowpoints = GetBowPoint(bowline, lineone)

                    If bowpoints(0) > 0 Then
                        'No point can be in the origin
                        x1list.Add(bowpoints(0))
                        y1list.Add(bowpoints(1))
                        x2list.Add(bowpoints(2))
                        y2list.Add(bowpoints(3))
                        length = GetLengthByPoints(bowpoints)
                        lengthlist.Add(length)
                        levellist.Add(1)
                        'Check if cropping is necessary - note: cropping only for evaporators needed

                        If length < 80 Then     '80 because 70.7mm is 1 diagonal step for N pattern - AGHN has diagonal cropped bows, 79.2mm is 1x2 diagonal in F Pattern
                            'Definitely cropped because of heating rod
                            typelist.Add(9)
                        Else
                            'to avoid unneccessary work, set type to 0 
                            'later check if a type 0 bow exists, if so then 
                            '-get in/out position and check if they are on a line
                            '-check if a line is below and the outbreak is only for better display in the drawing
                            typelist.Add(0)
                        End If
                    End If
                Next

                bowprops = {x1list, y1list, x2list, y2list, lengthlist}

                If typelist.Min = 0 Then
                    If connectionside = "front" Then        'check for inlet / outlet only on frontside
                        inoutlist = GetInOutCoords(objsheet, pitchx, pitchy, coilposition, "cooling", passnumber, coilnumber)
                        'for bow vs cp, only first 2 values used for initial check → extend list with switched values
                        For i As Integer = 0 To inoutlist(0).Count - 1
                            inoutlist(0).Add(inoutlist(2)(i))
                            inoutlist(1).Add(inoutlist(3)(i))
                            inoutlist(2).Add(inoutlist(0)(i))
                            inoutlist(3).Add(inoutlist(1)(i))
                        Next
                        For i As Integer = 0 To x1list.Count - 1
                            'check each bow with type = 0
                            If typelist(i) = 0 Then
                                bowpoints = {bowprops(0)(i), bowprops(1)(i), bowprops(2)(i), bowprops(3)(i)}
                                If CircProps.CheckBowvsCP(bowpoints, inoutlist) Then
                                    typelist(i) = 9
                                End If
                            End If

                        Next
                    End If
                    'check if another bow is "below"
                    For i As Integer = 0 To x1list.Count - 1
                        'check each bow with type = 0
                        If hasEDefrost <> "" Then
                            bowpoints = {bowprops(0)(i), bowprops(1)(i), bowprops(2)(i), bowprops(3)(i)}
                            If CircProps.CheckBowvsCP(bowpoints, Unit.heatingcoords) Then
                                typelist(i) = 9
                            End If
                        End If
                        If typelist(i) = 0 Then
                            'check bow vs bow - input: i,bowprops
                            If CircProps.CheckBowinBow(i, bowprops) Then
                                typelist(i) = 9
                            Else
                                typelist(i) = 1
                            End If
                        End If
                    Next
                End If
            End If

            'If fin is split, then the position of bows has to be recalculated for horizontal coilposition
            rangelimit = GetRangeLimit(objsheet, coilposition)
            If rangelimit > 0 And mcd <> "2" Then
                circframe = GetCoilFrame(objsheet, coilposition, General.currentunit.MultiCircuitDesign, coilnumber)
                If coilposition = "horizontal" Then
                    'Reduce all y values by rangelimit
                    For i As Integer = 0 To y1list.Count - 1
                        y1list(i) = Math.Round(y1list(i) - Math.Round(circframe(1), 2), 5)
                        y2list(i) = Math.Round(y2list(i) - Math.Round(circframe(1), 2), 5)
                    Next
                Else
                    'reduce all x values by rangelimit
                    For i As Integer = 0 To x1list.Count - 1
                        x1list(i) = Math.Round(x1list(i) - Math.Round(circframe(0), 2), 5)
                        x2list(i) = Math.Round(x2list(i) - Math.Round(circframe(0), 2), 5)
                    Next
                End If
            End If

            If typelist.IndexOf(9) > -1 Then
                For i As Integer = 0 To typelist.Count - 1
                    If typelist(i) = 9 Then
                        If connectionside = "front" Then
                            If x1list(i) > x2list(i) Then
                                xs = x2list(i)
                                ys = y2list(i)
                                xe = x1list(i)
                                ye = y1list(i)
                                x1list(i) = xs
                                y1list(i) = ys
                                x2list(i) = xe
                                y2list(i) = ye
                            End If
                        Else
                            If x2list(i) > x1list(i) Then
                                xs = x2list(i)
                                ys = y2list(i)
                                xe = x1list(i)
                                ye = y1list(i)
                                x1list(i) = xs
                                y1list(i) = ys
                                x2list(i) = xe
                                y2list(i) = ye
                            End If
                        End If
                    End If
                Next
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        bowprops = {x1list, y1list, x2list, y2list, lengthlist, typelist, levellist}

        Return bowprops
    End Function

    Shared Function GetBowPoint(bowline As SEFS.Line2d, outbreaklines As List(Of SEFS.Line2d)) As Double()
        Dim breakline As SEFS.Line2d
        Dim sp() As Double = {0, 0}
        Dim ep() As Double = {0, 0}
        Dim bowpoint(), xs, ys, xe, ye, xsb, ysb, xeb, yeb As Double

        bowline.GetStartPoint(xs, ys)
        bowline.GetEndPoint(xe, ye)
        xs = Math.Round(xs * 1000, 4)
        ys = Math.Round(ys * 1000, 4)
        xe = Math.Round(xe * 1000, 4)
        ye = Math.Round(ye * 1000, 4)

        For Each breakline In outbreaklines
            breakline.GetStartPoint(xsb, ysb)
            breakline.GetEndPoint(xeb, yeb)
            xsb = Math.Round(xsb * 1000, 4)
            ysb = Math.Round(ysb * 1000, 4)
            xeb = Math.Round(xeb * 1000, 4)
            yeb = Math.Round(yeb * 1000, 4)
            'If both lines share a point, then the other point of the breakline is in the grid and needed for the bowpoint
            If xs = xsb And ys = ysb Then
                sp = {xeb, yeb}
            ElseIf xs = xeb And ys = yeb Then
                sp = {xsb, ysb}
            ElseIf xe = xsb And ye = ysb Then
                ep = {xeb, yeb}
            ElseIf xe = xeb And ye = yeb Then
                ep = {xsb, ysb}
            End If
        Next

        bowpoint = {sp(0), sp(1), ep(0), ep(1)}

        Return bowpoint
    End Function

    Shared Function GetLengthByPoints(points() As Double) As Double
        Dim dx, dy, length As Double

        dx = Math.Abs(points(2) - points(0))
        dy = Math.Abs(points(3) - points(1))

        length = Math.Round(Math.Sqrt(dx ^ 2 + dy ^ 2), 3)

        Return length
    End Function

    Shared Function GetInOutCoords(objsheet As SED.Sheet, pitchx As Double, pitchy As Double, coilposition As String, circtype As String, passnumber As Integer, coilnumber As Integer) As List(Of Double)()
        Dim objlines As SEFS.Lines2d
        Dim grouplines As SEFS.Lines2d
        Dim objline As SEFS.Line2d
        Dim groupline As SEFS.Line2d
        Dim arrowlines As New List(Of SEFS.Line2d)
        Dim objgroups As SEFS.Groups
        Dim objgroup As SEFS.Group
        Dim xinlist, yinlist, xoutlist, youtlist, xtotallist, ytotallist As New List(Of Double)
        Dim totalcoords, uniquecoords As New List(Of String)
        Dim inoutlist() As List(Of Double)
        Dim rangelimit, xs, ys, xe, ye, xtemp, ytemp, xoffset, yoffset, circframe() As Double
        Dim gridcount, coordcounter As Integer
        Dim checkposition As Boolean = False
        Dim addline As Boolean

        Try

            objlines = objsheet.Lines2d
            'First check if fin is split
            rangelimit = GetRangeLimit(objsheet, coilposition)
            circframe = GetCoilFrame(objsheet, coilposition, General.currentunit.MultiCircuitDesign, coilnumber)

            If circtype.Contains("Defrost") Then
                xoffset = 25
                yoffset = 25
            Else
                xoffset = 0
                yoffset = 0
            End If

            If rangelimit <> 0 Then
                checkposition = True
            End If

            'Arrow lines for inlet / outlet are in a group → searching for all groups with "strand" or "Strang"
            objgroups = objsheet.Groups

            For Each objgroup In objgroups
                'Get group name, check for s[tran]d / S[tran]g
                If objgroup.UserDefinedName.Contains("tran") Then
                    'Get lines
                    grouplines = objgroup.Lines2d
                    For Each groupline In grouplines
                        arrowlines.Add(groupline)
                    Next
                End If
            Next

            For Each objline In arrowlines
                If objline.Length < 0.015 Then        'line belongs to and arrow (objline.Layer = "Default" Or objline.Layer = "Layer1" Or objline.Layer.Contains("")) And
                    'Check if one point of the line is in the grid
                    objline.GetStartPoint(xs, ys)
                    If checkposition Then
                        addline = False                       'set default value
                        If coilposition = "horizontal" Then
                            If coilnumber = 1 Then
                                If ys > rangelimit Then
                                    addline = True
                                    If General.currentunit.MultiCircuitDesign <> "2" Then
                                        ys = Math.Round(ys - circframe(1) / 1000, 5)
                                    End If
                                End If
                            Else
                                If ys < rangelimit Then
                                    addline = True
                                End If
                            End If
                        Else
                            If coilnumber = 1 Then
                                If xs > rangelimit Then
                                    If General.currentunit.MultiCircuitDesign <> "2" Then
                                        xs = Math.Round(xs - circframe(0) / 1000, 5)
                                    End If
                                    addline = True
                                End If
                            Else
                                If xs < rangelimit Then
                                    addline = True
                                End If
                            End If
                        End If
                    Else
                        addline = True
                    End If
                    'only if the line is in the correct area
                    If addline Then
                        xs = Math.Round(xs * 1000, 4)
                        ys = Math.Round(ys * 1000, 4)
                        gridcount = CircProps.CheckCoords({xs, ys}, pitchx, pitchy, xoffset, yoffset)
                        If gridcount = 0 Then
                            objline.GetEndPoint(xe, ye)
                            xe = Math.Round(xe * 1000, 4)
                            ye = Math.Round(ye * 1000, 4)
                            gridcount = CircProps.CheckCoords({xe, ye}, pitchx, pitchy, xoffset, yoffset)
                            If gridcount = 2 Then
                                xtotallist.Add(xe)
                                ytotallist.Add(ye)
                            End If
                        ElseIf gridcount = 2 Then
                            xtotallist.Add(xs)
                            ytotallist.Add(ys)
                        End If
                    End If
                End If
            Next

            'Rewrite the points into a string each and get unique points
            For i As Integer = 0 To xtotallist.Count - 1
                totalcoords.Add(xtotallist(i).ToString + "/" + ytotallist(i).ToString)
            Next

            uniquecoords = General.GetUniqueStrings(totalcoords)

            'Split the list into inlet and outlet - double entries = inlet / single entry = outlet
            For i As Integer = 0 To uniquecoords.Count - 1
                xtemp = CDbl(uniquecoords(i).Substring(0, uniquecoords(i).IndexOf("/")))
                ytemp = CDbl(uniquecoords(i).Substring(uniquecoords(i).IndexOf("/") + 1))

                coordcounter = 0
                For Each entry In totalcoords
                    If entry = uniquecoords(i) Then
                        coordcounter += 1
                    End If
                Next
                If coordcounter = 1 Then
                    xoutlist.Add(xtemp)
                    youtlist.Add(ytemp)
                Else
                    xinlist.Add(xtemp)
                    yinlist.Add(ytemp)
                End If
            Next

            If passnumber = 1 Then
                xoutlist.AddRange(xinlist.ToArray)
                youtlist.AddRange(yinlist.ToArray)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        inoutlist = {xinlist, yinlist, xoutlist, youtlist}

        Return inoutlist
    End Function

    Shared Sub CreateDummyDrawing(originalname As String, newname As String, f As FileData)
        Dim dftdoc As SED.DraftDocument
        Dim modellink As SED.ModelLink

        Try
            If File.Exists(originalname) Then
                dftdoc = OpenDFT(originalname)
                modellink = dftdoc.ActiveSheet.DrawingViews.Item(1).ModelLink

                If File.Exists(newname.Replace(".dft", ".par")) Then
                    modellink.ChangeSource(newname.Replace(".dft", ".par"))
                    For Each DV As SED.DrawingView In dftdoc.ActiveSheet.DrawingViews
                        DV.Update()
                    Next
                End If

                WriteCostumProps(dftdoc, f)
                'SEPart.OrderProps("dft", CoolingConSys.OrderData.Orderno, CoolingConSys.OrderData.Position, sematerial:=material)

                dftdoc.SaveAs(newname)
                General.seapp.Documents.CloseDocument(dftdoc.FullName, SaveChanges:=True, DoIdle:=True)

            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function GetSupportPositions(dftdoc As SED.DraftDocument, conside As String, circtype As String) As List(Of Double)()
        Dim objsheet As SED.Sheet
        Dim angle As Double
        Dim xs, xe, ys, ye, xp, yp As Double
        Dim xplist, yplist As New List(Of Double)
        Dim v2 As Integer

        Try

            'check if mirrored
            If (conside = "left" And Not circtype.Contains("defrost")) Or (conside = "right" And circtype.Contains("defrost")) Then
                'move all objects by frame width in x direction
                CreateMirroredDFT(dftdoc, False, conside, "vertical", False)
            End If

            objsheet = FindSheet(dftdoc)

            If objsheet IsNot Nothing Then
                For Each objgroup As SEFS.Group In objsheet.Groups
                    If objgroup.UserDefinedName.Contains("contact tube") Or objgroup.UserDefinedName.Contains("Kontaktrohr") Then
                        If objgroup.Lines2d.Count = 3 Then
                            Dim xlist, ylist As New List(Of Double)
                            For Each objline As SEFS.Line2d In objgroup.Lines2d
                                angle = Math.Round(objline.Angle, 4)
                                If angle <> 0 And Math.Round(Math.PI - angle, 2) <> 0 Then
                                    objline.GetStartPoint(xs, ys)
                                    objline.GetEndPoint(xe, ye)
                                    xs = Math.Round(xs * 1000, 3)
                                    xe = Math.Round(xe * 1000, 3)
                                    ys = Math.Round(ys * 1000, 3)
                                    ye = Math.Round(ye * 1000, 3)
                                    If xlist.IndexOf(xs) > -1 Then
                                        xp = xs
                                        v2 = General.IntegerRem(Math.Max(ye, ys), 12.5)
                                        yp = v2 * 12.5
                                        xplist.Add(xp)

                                        yplist.Add(yp)
                                        Exit For
                                    Else
                                        xlist.Add(xs)
                                        If xlist.IndexOf(xe) > -1 Then
                                            xp = xe
                                            v2 = General.IntegerRem(Math.Max(ye, ys), 12.5)
                                            yp = v2 * 12.5
                                            xplist.Add(xp)
                                            yplist.Add(yp)
                                            Exit For
                                        Else
                                            xlist.Add(xe)
                                        End If

                                    End If
                                End If
                            Next
                        End If
                    End If
                Next

            End If

            General.seapp.Documents.CloseDocument(dftdoc.FullName, SaveChanges:=False)
        Catch ex As Exception

        End Try

        Return {xplist, yplist}
    End Function

    Shared Sub WriteCostumProps(dftdoc As SED.DraftDocument, f As FileData)
        SEPart.GetSetCustomProp(dftdoc, "CSG", "1", "write")
        SEPart.GetSetCustomProp(dftdoc, "Auftragsnummer", f.Orderno, "write")
        SEPart.GetSetCustomProp(dftdoc, "Position", f.Orderpos, "write")
        SEPart.GetSetCustomProp(dftdoc, "Order_Projekt", f.Projectno, "write")
        SEPart.GetSetCustomProp(dftdoc, "AGP_Nummer", f.AGPno + ".", "write")
        SEPart.GetSetCustomProp(dftdoc, "AGP_Plant", f.Plant, "write")

        If CInt(f.AGPno.Substring(0, 3)) < 106 Then
            SEPart.GetSetCustomProp(dftdoc, "Z_Kategorie", "2D-Baugruppenzeichnung", "write")
        Else
            SEPart.GetSetCustomProp(dftdoc, "Z_Kategorie", "2D-Einzelteilzeichnung", "write")
        End If

    End Sub

    Shared Function UpdateDrawing(psmfile As String) As String
        Dim dftfile As String = General.GetFullFilename(General.currentjob.Workspace, psmfile.Substring(psmfile.LastIndexOf("\") + 1, 10), ".dft")
        Dim dftdoc As SED.DraftDocument

        Try
            If dftfile <> "" Then
                dftdoc = OpenDFT(dftfile)

                For Each DV As SED.DrawingView In dftdoc.ActiveSheet.DrawingViews
                    DV.Update()
                Next

                SEPart.CoversheetProps("dft")

                General.seapp.Documents.CloseDocument(dftfile, SaveChanges:=True, DoIdle:=True)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
        Return dftfile
    End Function

    Shared Sub RenameCoversheet(oldfilename As String, newpsmfile As String)
        Dim newfilename As String
        Dim tempname As String
        Dim dftdoc As SED.DraftDocument = Nothing

        Try
            tempname = "CS" + oldfilename.Substring(oldfilename.LastIndexOf("\") + 5, 10) + ".dft"
            newfilename = General.currentjob.Workspace + "\" + tempname
            My.Computer.FileSystem.RenameFile(oldfilename, tempname)
            General.seapp.DisplayAlerts = False
            dftdoc = General.seapp.Documents.Open(newfilename)
            General.seapp.DoIdle()

            SEDrawing.ChangeModellink(dftdoc, General.currentjob.Workspace, tempname.Replace("dft", "psm"))
            For Each dv As SED.DrawingView In dftdoc.ActiveSheet.DrawingViews
                dv.Update()
            Next
            dftdoc.Save()
            General.seapp.DoIdle()
            General.seapp.DisplayAlerts = True
            'save it in PDM!
            WSM.SaveDFT()
            'wait until finished
            WSM.WaitforWSMDialog()
            Threading.Thread.Sleep(3000)

            Do
                dftdoc = TryCast(General.seapp.ActiveDocument, SED.DraftDocument)
                Threading.Thread.Sleep(1000)
            Loop Until dftdoc IsNot Nothing
            General.seapp.Documents.CloseDocument(dftdoc.FullName, SaveChanges:=True, DoIdle:=True)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function GFinSupport(objsheet As SED.Sheet, coil As CoilData) As List(Of Double)()
        Dim boundaries As SEFS.Boundaries2d = objsheet.Boundaries2d
        Dim bojectlist As New List(Of Object)
        Dim xlist, ylist As New List(Of Double)

        Try
            For Each boundary As SEFS.Boundary2d In boundaries
                If boundary.BoundingObjects.Count > 0 Then
                    bojectlist.Add(boundary.BoundingObjects.Item(1))
                End If
            Next

            For Each b In bojectlist
                'Debug.Print(TypeName(b))
                Dim objcirc As SEFS.Circle2d = TryCast(b, SEFS.Circle2d)
                If objcirc IsNot Nothing Then
                    Dim x, y As Double
                    objcirc.GetCenterPoint(x, y)
                    If coil.Alignment = "horizontal" Then
                        xlist.Add(Math.Round(x, 6))
                        ylist.Add(Math.Round(y, 6))
                    Else
                        'x = FH-y   // y = x
                        xlist.Add(Math.Round(coil.FinnedHeight / 1000 - y, 6))
                        ylist.Add(Math.Round(x, 6))
                    End If
                End If
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return {xlist, ylist}
    End Function

    Shared Function GetADPosition(objsheet As SED.Sheet) As String
        Dim objgroup As SEFS.Group
        Dim xmin, ymin, xmax, ymax As Double
        Dim location As String = ""
        Try
            objgroup = GetGroupByName(objsheet, "LR_AD")
            If objgroup IsNot Nothing Then
                objgroup.Range(xmin, ymin, xmax, ymax)
                If xmin < 0 Then
                    location = "left"
                Else
                    location = "right"
                End If
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
        Return location
    End Function

    Shared Function FindSheet2(dftfile As String, revsheetname As String) As SED.Sheet
        Dim dftdoc As SED.DraftDocument
        Dim objDV As SED.DrawingView
        Dim DVList As New List(Of SED.DrawingView)
        Dim objSheet As SED.Sheet = Nothing
        Dim sheetfound As Boolean = False
        Dim framewidth As Double

        Try
            dftdoc = OpenDFT(dftfile)

            If dftdoc.ActiveSheet.DrawingViews.Count > 1 Then
                For Each DV As SED.DrawingView In dftdoc.ActiveSheet.DrawingViews
                    Dim tempsheet As SED.Sheet = DV.Sheet
                    If tempsheet.Name <> revsheetname Then
                        framewidth = GetWidthFromGroup(DV.Sheet, "coil frame")
                        If framewidth = 0 Then
                            framewidth = GetWidthFromGroup(DV.Sheet, "BlockRahmen")
                        End If
                        If framewidth <> 0 Then
                            objSheet = DV.Sheet
                            DVList.Add(DV)
                        End If
                    End If
                Next

                For i As Integer = 0 To DVList.Count - 1
                    objSheet = DVList(i).Sheet
                    If objSheet.Name = objSheet.Key Then
                        DVList(i).Sheet.Activate()
                        General.seapp.DoIdle()
                        objSheet = dftdoc.ActiveSheet
                        Threading.Thread.Sleep(2000)
                        sheetfound = True
                        Exit For
                    End If
                Next

                'Sometimes the first way doesn't work, so the second drawing view will be used no matter what
                If sheetfound = False Then
                    objDV = DVList(1)
                    objSheet = objDV.Sheet
                    objSheet.Activate()
                    General.seapp.DoIdle()
                    Threading.Thread.Sleep(2000)
                End If

            Else
                objSheet = dftdoc.ActiveSheet.DrawingViews.Item(1).Sheet
                objSheet.Activate()
                General.seapp.DoIdle()
                Threading.Thread.Sleep(2000)
            End If

        Catch ex As Exception

        End Try

        Return objSheet
    End Function

    Shared Function CountStrandGroups(objsheet As SED.Sheet) As Integer
        Dim objgroups As SEFS.Groups = objsheet.Groups
        Dim count As Integer = 0

        For Each objgroup In objgroups
            'Get group name, check for s[tran]d / S[tran]g
            If objgroup.UserDefinedName.Contains("tran") Then
                count += 1
            End If
        Next
        Return count
    End Function
End Class
