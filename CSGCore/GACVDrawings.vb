Imports Microsoft.SqlServer.Server

Public Class GACVDrawings
    Shared frontDVLinelist As New List(Of DVLineElement)
    Shared sideDVArclist As New List(Of SolidEdgeDraft.DVArc2d)
    Shared inletIDs, outletIDs As New List(Of String)
    Shared uniqueinletsdata, uniqueoutletsdata As New List(Of StutzenData)

    Shared Sub ResetsLists()
        frontDVLinelist.Clear()
        sideDVArclist.Clear()
        inletIDs.Clear()
        outletIDs.Clear()
        uniqueinletsdata.Clear()
        uniqueoutletsdata.Clear()
    End Sub

    Shared Sub MainCoil(dftdoc As SolidEdgeDraft.DraftDocument, coil As CoilData)
        Dim adposition As String

        Try
            AddBowViews(dftdoc, coil)

            If coil.Circuits.First.ConnectionSide = "right" Then
                adposition = "left"
            Else
                adposition = "right"
            End If

            CreateADBows(dftdoc, adposition, coil.Circuits.First)

            'if needed display electrical defrost
            If coil.EDefrostPDMID <> "" Then
                CreateHeatingBows(dftdoc, coil.ConSyss.First.VType, coil)
            End If

            'display fin separation line
            If coil.FinnedHeight = 1600 And (coil.Circuits.First.FinType = "N" Or (coil.Circuits.First.FinType = "M")) Then
                FinSeparation(dftdoc, coil)
            End If

            CreatePartListCoil(dftdoc)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub MainConsys(consys As ConSysData, coil As CoilData, circuit As CircuitData, dftdoc As SolidEdgeDraft.DraftDocument)

        Try
            'reset lists
            ResetsLists()
            'fill the lists
            FillStutzenLists(consys)

            If consys.HeaderAlignment = "vertical" Or circuit.IsOnebranchEvap Then
                If CreateDrawingVer(consys, coil, circuit, dftdoc) Then
                    CreateAirDirection(dftdoc, circuit.ConnectionSide, coil.FinnedDepth)

                    If circuit.IsOnebranchEvap Then
                        'different logic for FP and DX!
                        OneBranchFrontViewDims(dftdoc, consys, coil, circuit)
                    Else
                        If General.currentunit.ModelRangeSuffix.Substring(1, 1) = "X" And circuit.CircuitType <> "Defrost" Then
                            FrontViewDims(dftdoc, consys, coil, circuit, "outlet")

                            SideViewDims(dftdoc, consys, coil, circuit)

                            TopViewDims(dftdoc, consys, circuit)
                        Else
                            FrontViewDims(dftdoc, consys, coil, circuit, "outlet")
                            FrontViewDims(dftdoc, consys, coil, circuit, "inlet")

                            'for each header a single drawing view
                            FPSideViewDims(dftdoc, consys, dftdoc.ActiveSheet.DrawingViews.Item(3), consys.OutletHeaders.First, circuit, coil)
                            FPSideViewDims(dftdoc, consys, dftdoc.ActiveSheet.DrawingViews.Item(4), consys.InletHeaders.First, circuit, coil)

                            FPTopViewDims(dftdoc, dftdoc.ActiveSheet.DrawingViews.Item(2), consys, consys.OutletHeaders.First, circuit, coil)
                            FPTopViewDims(dftdoc, dftdoc.ActiveSheet.DrawingViews.Item(2), consys, consys.InletHeaders.First, circuit, coil)
                        End If
                    End If
                    IsoView(dftdoc, consys, circuit, coil)

                    Partlist(dftdoc, circuit, consys.HeaderAlignment)
                End If
            Else
                If CreateDrawingHor(consys, coil, circuit, dftdoc) Then
                    dftdoc = General.seapp.ActiveDocument

                    CreateAirDirection(dftdoc, circuit.ConnectionSide, coil.FinnedDepth)

                    HorFrontViewDims(dftdoc, consys, circuit)

                    HorHeaderViews(dftdoc, dftdoc.ActiveSheet.DrawingViews.Item(2), consys.OutletHeaders.First, circuit, General.GetUniqueStrings(outletIDs), uniqueoutletsdata)
                    HorHeaderViews(dftdoc, dftdoc.ActiveSheet.DrawingViews.Item(3), consys.InletHeaders.First, circuit, General.GetUniqueStrings(inletIDs), uniqueinletsdata)

                    TopViewHor(dftdoc, consys, circuit)

                    Partlist(dftdoc, circuit, consys.HeaderAlignment)

                    MoveIsoView(dftdoc, coil.FinnedHeight)

                    If consys.HasHotgas Then
                        DimHotGas(dftdoc, consys, circuit)
                    End If

                    'handle visibility of drawings views now


                End If
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub MainConsysDefrost(consys As ConSysData, coil As CoilData, circuit As CircuitData, dftdoc As SolidEdgeDraft.DraftDocument)

        Try
            ResetsLists()

            If CreateDrawingHor(consys, coil, circuit, dftdoc) Then
                'fill the lists
                FillStutzenLists(consys)

                CreateAirDirection(dftdoc, circuit.ConnectionSide, coil.FinnedDepth)

                If circuit.NoDistributions = 1 Then
                    'different logic for FP and DX!
                    OneBranchFrontViewDims(dftdoc, consys, coil, circuit)
                ElseIf consys.HeaderAlignment = "vertical" Then
                    FrontViewDims(dftdoc, consys, coil, circuit, "outlet")
                    FrontViewDims(dftdoc, consys, coil, circuit, "inlet")

                    'for each header a single drawing view
                    BrineSideViewDims(dftdoc, consys, dftdoc.ActiveSheet.DrawingViews.Item(2), consys.OutletHeaders.First, circuit, coil)
                    BrineSideViewDims(dftdoc, consys, dftdoc.ActiveSheet.DrawingViews.Item(3), consys.InletHeaders.First, circuit, coil)

                    FPTopViewDims(dftdoc, dftdoc.ActiveSheet.DrawingViews.Item(4), consys, consys.OutletHeaders.First, circuit, coil)
                    FPTopViewDims(dftdoc, dftdoc.ActiveSheet.DrawingViews.Item(4), consys, consys.InletHeaders.First, circuit, coil)

                    Partlist(dftdoc, circuit, consys.HeaderAlignment)

                    MoveIsoView(dftdoc, coil.FinnedHeight)
                Else
                    HorFrontViewDims(dftdoc, consys, circuit)

                    HorHeaderViews(dftdoc, dftdoc.ActiveSheet.DrawingViews.Item(2), consys.OutletHeaders.First, circuit, General.GetUniqueStrings(outletIDs), uniqueoutletsdata)
                    HorHeaderViews(dftdoc, dftdoc.ActiveSheet.DrawingViews.Item(3), consys.InletHeaders.First, circuit, General.GetUniqueStrings(inletIDs), uniqueinletsdata)

                    TopViewHor(dftdoc, consys, circuit)

                    Partlist(dftdoc, circuit, consys.HeaderAlignment)

                    MoveIsoView(dftdoc, coil.FinnedHeight)
                End If
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub AddBowViews(dftdoc As SolidEdgeDraft.DraftDocument, coil As CoilData)
        Dim objmls As SolidEdgeDraft.ModelLinks
        Dim asmlink As SolidEdgeDraft.ModelLink
        Dim firstfrontDV, firstbackDV As SolidEdgeDraft.DrawingView
        Dim scalefactor, xmin, ymin, xmax, ymax, sheetframe(), maxDVcount, x0, y0, xm, x0calc As Double
        Dim mp As Integer

        Try
            dftdoc.Sheets.Item("A3").Delete()

            sheetframe = {0.59, 0.41}
            dftdoc.ActiveSheet.Name = "Coil1"

            If General.currentjob.Plant = "Beji" Then
                'switch layer for background
                SwitchLayers(dftdoc, General.currentjob.Plant)
            End If

            objmls = dftdoc.ModelLinks
            asmlink = objmls.Add(coil.CoilFile.Fullfilename)

            scalefactor = GetScaling(coil.FinnedHeight, coil.FinnedDepth, "Bows")

            If coil.Frontbowids.Count = 0 Then
                'only backbows → one row
                For i As Integer = 0 To coil.Backbowids.Count - 1
                    SEDrawing.AddBowViewtoSheet(dftdoc.ActiveSheet, asmlink, scalefactor, coil.Backbowids(i), "Back", coil)
                Next
                firstbackDV = dftdoc.ActiveSheet.DrawingViews.Item(1)
                firstbackDV.Range(xmin, ymin, xmax, ymax)
                y0 = 0.41 - ymax - 2 * 0.0065

                'calculate x location depending of DVcount
                For i As Integer = 0 To coil.Backbowids.Count - 1
                    x0 = Math.Round(sheetframe(0) / (1 + coil.Backbowids.Count), 6) * (i + 1)
                    dftdoc.ActiveSheet.DrawingViews.Item(i + 1).SetOrigin(x0, y0)
                Next

            ElseIf coil.Backbowids.Count = 0 Then
                'only frontbows → one row
                For i As Integer = 0 To coil.Frontbowids.Count - 1
                    SEDrawing.AddBowViewtoSheet(dftdoc.ActiveSheet, asmlink, scalefactor, coil.Frontbowids(i), "Front", coil)
                Next
                firstfrontDV = dftdoc.ActiveSheet.DrawingViews.Item(1)
                firstfrontDV.Range(xmin, ymin, xmax, ymax)
                y0 = 0.41 - ymax - 2 * 0.0065

                'calculate x location depending of DVcount
                For i As Integer = 0 To coil.Frontbowids.Count - 1
                    x0 = Math.Round(sheetframe(0) / (1 + coil.Frontbowids.Count), 6) * (i + 1)
                    dftdoc.ActiveSheet.DrawingViews.Item(i + 1).SetOrigin(x0, y0)
                Next

            Else
                firstfrontDV = SEDrawing.AddBowViewtoSheet(dftdoc.ActiveSheet, asmlink, scalefactor, coil.Frontbowids.First, "Front", coil)

                'check how many drawing views fit in one line
                maxDVcount = Calculation.MaxDVcount(firstfrontDV, coil.FinnedDepth)

                If maxDVcount < coil.Frontbowids.Count + coil.Backbowids.Count Then
                    'additional sheet
                    SEDrawing.AddSheet(dftdoc.ActiveSheet)
                    dftdoc.ActiveSheet.Name = "Front"
                    dftdoc.Sheets.Item("Bows").Name = "Back"

                    firstfrontDV.Range(xmin, ymin, xmax, ymax)

                    'check front & back if maxDVcount bigger than needed DVcount → divide into 2 rows and check again
                    If maxDVcount >= coil.Frontbowids.Count Then
                        'normal procedure
                        If coil.Frontbowids.Count > 1 Then
                            'first DV has been placed already
                            For i As Integer = 1 To coil.Frontbowids.Count - 1
                                SEDrawing.AddBowViewtoSheet(dftdoc.ActiveSheet, asmlink, scalefactor, coil.Frontbowids(i), "Front", coil)
                            Next
                        End If
                        y0 = 0.41 - ymax - 2 * 0.0065

                        'calculate x location depending of DVcount
                        For i As Integer = 0 To coil.Frontbowids.Count - 1
                            x0 = Math.Round(sheetframe(0) / (1 + coil.Frontbowids.Count), 6) * (i + 1)
                            dftdoc.ActiveSheet.DrawingViews.Item(i + 1).SetOrigin(x0, y0)
                        Next
                    Else
                        If maxDVcount >= Math.Ceiling(coil.Frontbowids.Count / 2) Then
                            'check for rescale
                            Do
                                scalefactor = Calculation.RescaleFactor(scalefactor, ymax, "down", 0.17)
                                firstfrontDV.ScaleFactor = scalefactor
                                firstfrontDV.Range(xmin, ymin, xmax, ymax)
                            Loop Until Math.Round(2 * ymax, 6) <= 0.17
                        Else
                            'recalc maxDVcount with smaller scalefactors until 2 rows fit
                            Do
                                scalefactor = Calculation.RescaleFactor(scalefactor, ymax, "down", 0)
                                firstfrontDV.ScaleFactor = scalefactor
                                maxDVcount = Calculation.MaxDVcount(firstfrontDV, coil.FinnedDepth)
                            Loop Until maxDVcount <= Math.Ceiling(coil.Frontbowids.Count / 2)
                            'check for rescale
                            firstfrontDV.Range(xmin, ymin, xmax, ymax)
                            Do
                                scalefactor = Calculation.RescaleFactor(scalefactor, ymax, "down", 0.17)
                                firstfrontDV.ScaleFactor = scalefactor
                                firstfrontDV.Range(xmin, ymin, xmax, ymax)
                            Loop Until Math.Round(2 * ymax, 6) <= 0.17
                        End If

                        'top row
                        y0 = 0.41 - ymax - 2 * 0.0065
                        For i As Integer = 1 To Math.Ceiling(coil.Frontbowids.Count / 2) - 1
                            SEDrawing.AddBowViewtoSheet(dftdoc.ActiveSheet, asmlink, scalefactor, coil.Frontbowids(i), "Front", coil)
                        Next
                        For i As Integer = 0 To Math.Ceiling(coil.Frontbowids.Count / 2) - 1
                            x0 = Math.Round(sheetframe(0) / (1 + Math.Ceiling(coil.Frontbowids.Count / 2)), 6) * (i + 1)
                            dftdoc.ActiveSheet.DrawingViews.Item(i + 1).SetOrigin(x0, y0)
                        Next

                        'bottom row
                        y0 = 0.41 - 3 * ymax - 4 * 0.0065
                        For i As Integer = Math.Ceiling(coil.Frontbowids.Count / 2) To coil.Frontbowids.Count - 1
                            SEDrawing.AddBowViewtoSheet(dftdoc.ActiveSheet, asmlink, scalefactor, coil.Frontbowids(i), "Front", coil)
                        Next
                        Dim j As Integer = 0
                        For i As Integer = Math.Ceiling(coil.Frontbowids.Count / 2) To coil.Frontbowids.Count - 1
                            x0 = Math.Round(sheetframe(0) / (1 + Math.Ceiling(coil.Frontbowids.Count / 2)), 6) * (j + 1)
                            dftdoc.ActiveSheet.DrawingViews.Item(i + 1).SetOrigin(x0, y0)
                            j += 1
                        Next

                    End If

                    'switch to 2nd sheet
                    dftdoc.Sheets.Item("Back").Activate()
                    General.seapp.DoIdle()

                    'start with defaul scalefactor & maxDVcount again, front could have changed it
                    scalefactor = GetScaling(coil.FinnedHeight, coil.FinnedDepth, "Bows")

                    firstbackDV = SEDrawing.AddBowViewtoSheet(dftdoc.ActiveSheet, asmlink, scalefactor, coil.Backbowids.First, "Back", coil)
                    firstbackDV.Range(xmin, ymin, xmax, ymax)
                    maxDVcount = Calculation.MaxDVcount(firstbackDV, coil.FinnedDepth)

                    'same for backbows
                    If maxDVcount >= coil.Backbowids.Count Then
                        If coil.Backbowids.Count > 1 Then
                            For i As Integer = 1 To coil.Backbowids.Count - 1
                                SEDrawing.AddBowViewtoSheet(dftdoc.ActiveSheet, asmlink, scalefactor, coil.Backbowids(i), "Back", coil)
                            Next
                        End If
                        y0 = 0.41 - ymax - 2 * 0.0065

                        'calculate x location depending of DVcount
                        For i As Integer = 0 To coil.Backbowids.Count - 1
                            x0 = Math.Round(sheetframe(0) / (1 + coil.Backbowids.Count), 6) * (i + 1)
                            dftdoc.ActiveSheet.DrawingViews.Item(i + 1).SetOrigin(x0, y0)
                        Next
                    Else
                        If maxDVcount >= Math.Ceiling(coil.Backbowids.Count / 2) Then
                            'check for rescale
                            Do
                                scalefactor = Calculation.RescaleFactor(scalefactor, ymax, "down", 0.17)
                                firstbackDV.ScaleFactor = scalefactor
                                firstbackDV.Range(xmin, ymin, xmax, ymax)
                            Loop Until Math.Round(2 * ymax, 6) <= 0.17
                        Else
                            'recalc maxDVcount with smaller scalefactors until 2 rows fit
                            Do
                                scalefactor = Calculation.RescaleFactor(scalefactor, ymax, "down", 0)
                                firstbackDV.ScaleFactor = scalefactor
                                maxDVcount = Calculation.MaxDVcount(firstbackDV, coil.FinnedDepth)
                            Loop Until maxDVcount <= Math.Ceiling(coil.Backbowids.Count / 2)
                            firstbackDV.Range(xmin, ymin, xmax, ymax)
                            'check for rescale
                            Do
                                scalefactor = Calculation.RescaleFactor(scalefactor, ymax, "down", 0.17)
                                firstbackDV.ScaleFactor = scalefactor
                                firstbackDV.Range(xmin, ymin, xmax, ymax)
                            Loop Until Math.Round(2 * ymax, 6) <= 0.17
                        End If

                        'top row
                        y0 = 0.41 - ymax - 2 * 0.0065
                        For i As Integer = 1 To Math.Ceiling(coil.Backbowids.Count / 2) - 1
                            SEDrawing.AddBowViewtoSheet(dftdoc.ActiveSheet, asmlink, scalefactor, coil.Backbowids(i), "Back", coil)
                        Next
                        For i As Integer = 0 To Math.Ceiling(coil.Backbowids.Count / 2) - 1
                            x0 = Math.Round(sheetframe(0) / (1 + Math.Ceiling(coil.Backbowids.Count / 2)), 6) * (i + 1)
                            dftdoc.ActiveSheet.DrawingViews.Item(i + 1).SetOrigin(x0, y0)
                        Next

                        'bottom row
                        y0 = 0.41 - 3 * ymax - 4 * 0.0065
                        For i As Integer = Math.Ceiling(coil.Backbowids.Count / 2) To coil.Backbowids.Count - 1
                            SEDrawing.AddBowViewtoSheet(dftdoc.ActiveSheet, asmlink, scalefactor, coil.Backbowids(i), "Back", coil)
                        Next
                        Dim j As Integer = 0
                        For i As Integer = Math.Ceiling(coil.Backbowids.Count / 2) To coil.Backbowids.Count - 1
                            x0 = Math.Round(sheetframe(0) / (1 + Math.Ceiling(coil.Backbowids.Count / 2)), 6) * (j + 1)
                            dftdoc.ActiveSheet.DrawingViews.Item(i + 1).SetOrigin(x0, y0)
                            j += 1
                        Next
                    End If

                Else
                    'place them all in one line
                    Dim k As Integer = 1
                    Do
                        If coil.Frontbowids.Count > 1 Then
                            For i As Integer = 1 To coil.Frontbowids.Count - 1
                                SEDrawing.AddBowViewtoSheet(dftdoc.ActiveSheet, asmlink, scalefactor, coil.Frontbowids(i), "Front", coil)
                                k += 1
                            Next
                        End If
                        For i As Integer = 0 To coil.Backbowids.Count - 1
                            SEDrawing.AddBowViewtoSheet(dftdoc.ActiveSheet, asmlink, scalefactor, coil.Backbowids(i), "Back", coil)
                            k += 1
                        Next
                    Loop Until k >= coil.Frontbowids.Count + coil.Backbowids.Count

                    dftdoc.ActiveSheet.DrawingViews.Item(1).Range(xmin, ymin, xmax, ymax)

                    y0 = 0.41 - ymax - 2 * 0.0065

                    'set to 0/0 → current center xm → offset to origin of the DV
                    xm = (xmax + xmin) / 2

                    'calculate x location depending of DVcount
                    For i As Integer = 0 To (coil.Frontbowids.Count + coil.Backbowids.Count) - 1
                        dftdoc.ActiveSheet.DrawingViews.Item(i + 1).Range(xmin, ymin, xmax, ymax)

                        'calculated position for DV
                        x0calc = Math.Round(sheetframe(0) / (1 + coil.Frontbowids.Count + coil.Backbowids.Count), 6) * (i + 1)

                        'reverse the offset, if DV is for backside (left side of origin)
                        If i > coil.Frontbowids.Count - 1 Then
                            mp = -1
                        Else
                            mp = 1
                        End If

                        x0 = x0calc - xm * mp

                        dftdoc.ActiveSheet.DrawingViews.Item(i + 1).SetOrigin(x0, y0)
                    Next

                End If
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub CreateHeatingBows(dftdoc As SolidEdgeDraft.DraftDocument, vtype As String, coil As CoilData)
        Dim sheetlist As New List(Of SolidEdgeDraft.Sheet)
        Dim objsheet, mainsheet As SolidEdgeDraft.Sheet
        Dim caption, notcaption As String

        mainsheet = dftdoc.ActiveSheet

        Try
            If vtype = "X" Then
                caption = "Front"
                notcaption = "Back"
            Else
                caption = "Back"
                notcaption = "Front"
            End If

            If mainsheet.Name <> "Coil" + coil.Number.ToString Then
                sheetlist.Add(dftdoc.Sheets.Item(caption))
                sheetlist.Add(dftdoc.Sheets.Item(notcaption))
            Else
                sheetlist.Add(mainsheet)
            End If

            For Each objsheet In sheetlist
                objsheet.Activate()
                General.seapp.DoIdle()
                For Each objDV As SolidEdgeDraft.DrawingView In objsheet.DrawingViews
                    Dim objcircles As List(Of SolidEdgeDraft.DVCircle2d) = SEDrawing.GetCirclesFromOcc(objDV, New List(Of String) From {"CoilFin11.par"})
                    Dim stubelist As New List(Of SolidEdgeDraft.DVCircle2d)
                    Dim xclist, yclist As New List(Of Double)
                    For Each objcircle As SolidEdgeDraft.DVCircle2d In objcircles
                        If objcircle.Diameter = 0.011 Then
                            Dim xc, yc As Double
                            stubelist.Add(objcircle)
                            objcircle.GetCenterPoint(xc, yc)
                            xclist.Add(xc)
                            yclist.Add(yc)
                        End If
                    Next

                    Dim subsheet As SolidEdgeDraft.Sheet = objDV.Sheet
                    subsheet.Activate()

                    For i As Integer = 0 To stubelist.Count - 1
                        SEDrawing.CreateBoundary(subsheet, xclist(i), yclist(i), 11)
                    Next

                    If objDV.CaptionDefinitionTextPrimary.Contains(caption) Then
                        'transfer the heating bow coords (0,0) based into the consys drawing (x,y)
                        Dim dx, dy As Double

                        If caption.Contains("Back") Then
                            dx = coil.FinnedDepth / 1000 - Math.Max(coil.Defrostbows(0).Max, coil.Defrostbows(2).Max) / 1000 - xclist.Min
                        Else
                            dx = Math.Min(coil.Defrostbows(0).Min, coil.Defrostbows(2).Min) / 1000 - xclist.Min
                        End If
                        dy = Math.Min(coil.Defrostbows(1).Min, coil.Defrostbows(3).Min) / 1000 - yclist.Min

                        Dim dvlines As SolidEdgeFrameworkSupport.Lines2d = subsheet.Lines2d

                        For i As Integer = 0 To coil.Defrostbows(0).Count - 1
                            Dim dvline As SolidEdgeFrameworkSupport.Line2d
                            Dim xs, ys, xe, ye As Double

                            If objDV.CaptionDefinitionTextPrimary.Contains("Back") Then
                                xs = coil.FinnedDepth / 1000 - coil.Defrostbows(0)(i) / 1000 - dx
                                xe = coil.FinnedDepth / 1000 - coil.Defrostbows(2)(i) / 1000 - dx
                            Else
                                xs = coil.Defrostbows(0)(i) / 1000 - dx
                                xe = coil.Defrostbows(2)(i) / 1000 - dx
                            End If

                            ys = coil.Defrostbows(1)(i) / 1000 - dy
                            ye = coil.Defrostbows(3)(i) / 1000 - dy

                            dvline = dvlines.AddBy2Points(xs, ys, xe, ye)
                            dvline.Style.LinearColor = 255
                            dvline.Style.Width = 0.001
                        Next
                    End If

                Next

            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        Finally
            mainsheet.Activate()
        End Try

    End Sub

    Shared Sub FinSeparation(dftdoc As SolidEdgeDraft.DraftDocument, coil As CoilData)
        Dim sheetlist As New List(Of SolidEdgeDraft.Sheet)
        Dim objsheet, mainsheet As SolidEdgeDraft.Sheet
        Dim objTBs As SolidEdgeFrameworkSupport.TextBoxes
        Dim objTB As SolidEdgeFrameworkSupport.TextBox
        Dim mainline As SolidEdgeFrameworkSupport.Line2d
        Dim mainstyle As SolidEdgeFrameworkSupport.GeometryStyle2d

        mainsheet = dftdoc.ActiveSheet
        Try
            If mainsheet.Name <> "Coil1" Then
                sheetlist.Add(dftdoc.Sheets.Item("Front"))
                sheetlist.Add(dftdoc.Sheets.Item("Back"))
            Else
                sheetlist.Add(mainsheet)
            End If

            For Each objsheet In sheetlist
                objsheet.Activate()

                'add line
                mainline = objsheet.Lines2d.AddBy2Points(0.03, 0.025, 0.05, 0.025)
                mainstyle = mainline.Style
                With mainstyle
                    .Width = 0.0005
                    .LinearColor = 0
                    .LinearName = "Normal"
                    .DashName = "Dash"
                End With

                'add textbox
                objTBs = objsheet.TextBoxes
                objTB = objTBs.Add(0.055, 0.035, 0)
                objTB.Edit.TextSize = 0.006
                objTB.Text = "Lamellentrennung 18+14 Rohrlagen (900 / 700 mm)" + vbNewLine + "Fin separation 18+14 tube rows (900 / 700 mm)"


                General.seapp.DoIdle()
                For Each objDV As SolidEdgeDraft.DrawingView In objsheet.DrawingViews
                    'get left and right frame line for start and end 
                    Dim frame() As Double = GetFrame(objDV, coil.FinnedDepth)

                    Dim dvsheet As SolidEdgeDraft.Sheet = objDV.Sheet
                    'draw the line
                    Dim objline As SolidEdgeFrameworkSupport.Line2d = dvsheet.Lines2d.AddBy2Points(frame(0), frame(2) + 0.7, frame(1), frame(2) + 0.7)
                    Dim objstyle As SolidEdgeFrameworkSupport.GeometryStyle2d = objline.Style
                    objstyle.Width = 0.0005
                    objstyle.LinearColor = 0
                    objstyle.LinearName = "Normal"
                    objstyle.DashName = "Dash"
                Next
            Next
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        Finally
            mainsheet.Activate()
        End Try
    End Sub

    Shared Function GetFrame(objDV As SolidEdgeDraft.DrawingView, finneddepth As Double) As Double()
        Dim horlines, vertlines As New List(Of SolidEdgeDraft.DVLine2d)
        Dim xs, ys As Double
        Dim xlist, ylist As New List(Of Double)

        For Each objDVline As SolidEdgeDraft.DVLine2d In objDV.DVLines2d
            Dim length As Double = Math.Round(objDVline.Length * 1000, 3)
            If length = 1600 Then
                vertlines.Add(objDVline)
            ElseIf Length = finneddepth Then
                horlines.Add(objDVline)
            End If
        Next

        For Each vertline In vertlines
            vertline.GetStartPoint(xs, ys)
            xlist.Add(xs)
        Next

        For Each horline In horlines
            horline.GetStartPoint(xs, ys)
            ylist.Add(ys)
        Next

        Return {xlist.Min, xlist.Max, ylist.Min}
    End Function

    Shared Sub FillStutzenLists(consys As ConSysData)

        inletIDs.Clear()
        outletIDs.Clear()
        uniqueinletsdata.Clear()
        uniqueinletsdata.Clear()

        For Each h In consys.InletHeaders
            For i As Integer = 0 To h.StutzenDatalist.Count - 1
                If inletIDs.IndexOf(h.StutzenDatalist(i).ID) = -1 Then
                    uniqueinletsdata.Add(h.StutzenDatalist(i))
                    uniqueinletsdata.Last.Angle = Math.Abs(uniqueinletsdata.Last.Angle)
                End If
                inletIDs.Add(h.StutzenDatalist(i).ID)
            Next
        Next

        If uniqueinletsdata.Count > 0 Then
            Dim templist = From plist In uniqueinletsdata Order By plist.Angle

            uniqueinletsdata = templist.ToList
        End If

        For Each h In consys.OutletHeaders
            For i As Integer = 0 To h.StutzenDatalist.Count - 1
                If outletIDs.IndexOf(h.StutzenDatalist(i).ID) = -1 Then
                    uniqueoutletsdata.Add(h.StutzenDatalist(i))
                    uniqueoutletsdata.Last.Angle = Math.Abs(uniqueoutletsdata.Last.Angle)
                End If
                outletIDs.Add(h.StutzenDatalist(i).ID)
            Next
        Next

        If uniqueoutletsdata.Count > 0 Then
            Dim templist = From plist In uniqueoutletsdata Order By plist.Angle

            uniqueoutletsdata = templist.ToList
        End If

    End Sub

    Shared Function CreateDrawingVer(consys As ConSysData, coil As CoilData, circuit As CircuitData, dftdoc As SolidEdgeDraft.DraftDocument) As Boolean
        Dim sheetsize, material As String
        Dim stutzenlist As New List(Of String)
        Dim objDVs As SolidEdgeDraft.DrawingViews
        Dim topDV, frontDV, sideDV, isoDV As SolidEdgeDraft.DrawingView
        Dim DVlist As New List(Of SolidEdgeDraft.DrawingView)
        Dim objmls As SolidEdgeDraft.ModelLinks
        Dim asmlink As SolidEdgeDraft.ModelLink
        Dim swatch As New Stopwatch
        Dim scalefactor, mp As Double
        Dim sideno, isono As Integer
        Dim success As Boolean = True

        Try

            'add the CSG Property
            material = consys.HeaderMaterial

            objmls = dftdoc.ModelLinks
            asmlink = objmls.Add(consys.ConSysFile.Fullfilename)

            mp = 1
            If coil.FinnedHeight > 500 Or General.currentunit.ModelRangeSuffix = "FP" Or General.currentunit.ModelRangeSuffix = "WP" Then
                dftdoc.Sheets.Item("A2").Activate()
                General.seapp.DoIdle()
                dftdoc.Sheets.Item("A3").Delete()
                sheetsize = "A2"
                If coil.FinnedHeight <= 500 Then
                    mp = 1.5
                End If
            Else
                dftdoc.Sheets.Item("A3").Activate()
                General.seapp.DoIdle()
                dftdoc.Sheets.Item("A2").Delete()
                sheetsize = "A3"
            End If
            scalefactor = GetScaling(coil.FinnedHeight, coil.FinnedDepth, "Main") * mp

            If General.currentunit.ModelRangeSuffix.Substring(1, 1) = "P" And circuit.ConnectionSide = "left" Then
                scalefactor = GetScaling(coil.FinnedHeight, coil.FinnedDepth, "MainRescale") * mp
            End If

            General.seapp.DoIdle()

            objDVs = dftdoc.ActiveSheet.DrawingViews

            frontDV = objDVs.AddAssemblyView(asmlink, 4, scalefactor, -0.8, 0, 0)
            frontDV.CaptionDefinitionTextPrimary = "FrontView"
            DVlist.Add(frontDV)

            'top view, needs cropping!
            topDV = objDVs.AddByFold(frontDV, 2, -0.8, 0)
            topDV.CaptionDefinitionTextPrimary = "TopView"
            topDV.Update()

            '3x (a+Ø)?
            'topDV.CropTop = -topDV.CropTop + 0.35 * scalefactor
            DVlist.Add(topDV)

            'side view, left for normal, right for mirrored
            If circuit.ConnectionSide = "left" Then
                sideno = 3
                isono = 8
            Else
                sideno = 4
                isono = 7
            End If

            If Not circuit.IsOnebranchEvap Then
                If consys.HeaderAlignment = "vertical" Then
                    sideDV = objDVs.AddByFold(frontDV, sideno, -0.8, 0)
                    sideDV.CaptionDefinitionTextPrimary = "SideView"
                Else
                    sideDV = objDVs.AddAssemblyView(asmlink, 1, scalefactor, -0.8, 0, 0, "Outlet")
                End If

                DVlist.Add(sideDV)

                If General.currentunit.ModelRangeSuffix.Substring(1, 1) = "P" Then
                    sideDV.CaptionDefinitionTextPrimary = "Outlet"
                    sideDV.Configuration = "Outlet"
                    sideDV.CaptionLocation = SolidEdgeFrameworkSupport.DimViewCaptionLocationConstants.igDimViewCaptionLocationTop
                    sideDV.DisplayCaption = True
                    sideDV.MatchConfiguration = True
                    sideDV.Update()

                    sideDV = Nothing
                    If consys.HeaderAlignment = "vertical" Then
                        sideDV = objDVs.AddByFold(frontDV, sideno, 0, 0)
                    Else
                        sideDV = objDVs.AddAssemblyView(asmlink, 1, scalefactor, -0.8, 0, 0, "Inlet")
                    End If
                    sideDV.CaptionDefinitionTextPrimary = "Inlet"
                    sideDV.Configuration = "Inlet"
                    sideDV.CaptionLocation = SolidEdgeFrameworkSupport.DimViewCaptionLocationConstants.igDimViewCaptionLocationTop
                    sideDV.DisplayCaption = True
                    sideDV.MatchConfiguration = True
                    sideDV.Update()

                    DVlist.Add(sideDV)
                Else
                    If circuit.ConnectionSide = "left" Then
                        sideDV.CropLeft = topDV.CropTop
                    Else
                        sideDV.CropRight = topDV.CropTop
                    End If
                End If
            Else
                consys.HeaderAlignment = "vertical"
                sideDV = objDVs.AddByFold(frontDV, sideno, -0.8, 0)
                sideDV.CaptionDefinitionTextPrimary = "SideView"
                DVlist.Add(sideDV)

            End If

            'set position of the drawing views, starting with top right
            DVPositions(DVlist, consys, sheetsize, coil, circuit)

            If circuit.IsOnebranchEvap Then
                sideDV.Delete()
            End If

            isoDV = objDVs.AddByFold(frontDV, isono, 0, 0)
            isoDV.Update()

            If General.currentjob.Plant = "Beji" Then
                'switch layer for background
                SwitchLayers(dftdoc, General.currentjob.Plant)
            End If

        Catch ex As Exception
            success = False
            General.CreateLogEntry(ex.ToString)
        End Try

        Return success
    End Function

    Shared Function CreateDrawingHor(consys As ConSysData, coil As CoilData, circuit As CircuitData, dftdoc As SolidEdgeDraft.DraftDocument) As Boolean
        Dim sheetsize, material As String
        Dim stutzenlist As New List(Of String)
        Dim objDVs As SolidEdgeDraft.DrawingViews
        Dim topDV, frontDV, isoDV, inletDV, outletDV As SolidEdgeDraft.DrawingView
        Dim DVlist As New List(Of SolidEdgeDraft.DrawingView)
        Dim objmls As SolidEdgeDraft.ModelLinks
        Dim asmlink As SolidEdgeDraft.ModelLink
        Dim swatch As New Stopwatch
        Dim scalefactor, mp As Double
        Dim sideno, isono, frontno As Integer
        Dim success As Boolean = True

        Try

            'add the CSG Property
            material = consys.HeaderMaterial

            objmls = dftdoc.ModelLinks
            asmlink = objmls.Add(consys.ConSysFile.Fullfilename)

            mp = 1
            If coil.FinnedHeight > 500 Then
                dftdoc.Sheets.Item("A2").Activate()
                General.seapp.DoIdle()
                dftdoc.Sheets.Item("A3").Delete()
                sheetsize = "A2"
                If coil.FinnedHeight <= 500 Then
                    mp = 1.5
                End If
            Else
                dftdoc.Sheets.Item("A3").Activate()
                General.seapp.DoIdle()
                dftdoc.Sheets.Item("A2").Delete()
                sheetsize = "A3"
            End If
            scalefactor = GetScaling(coil.FinnedHeight, coil.FinnedDepth, "Main") * mp

            If coil.FinnedDepth > 399 Then
                scalefactor = GetScaling(coil.FinnedHeight, coil.FinnedDepth, "MainRescale")
            End If
            General.seapp.DoIdle()

            objDVs = dftdoc.ActiveSheet.DrawingViews

            'Defrost consys is placed on the backside / behind the fin → different DVs
            If circuit.CircuitType = "Defrost" Then
                frontno = 6
            Else
                frontno = 4
            End If
            frontDV = objDVs.AddAssemblyView(asmlink, frontno, scalefactor, -0.8, 0, 0)
            frontDV.CaptionDefinitionTextPrimary = "FrontView"
            DVlist.Add(frontDV)

            'side view, left for normal, right for mirrored
            If circuit.ConnectionSide = "right" Then
                sideno = 3
                isono = 8
            Else
                sideno = 4
                isono = 7
            End If

            'assign config only after positioning
            If consys.HeaderAlignment = "vertical" And circuit.NoDistributions > 1 Then
                outletDV = objDVs.AddByFold(frontDV, sideno, -0.8, 0)
                outletDV.CaptionDefinitionTextPrimary = "Outlet"
                DVlist.Add(outletDV)

                inletDV = objDVs.AddByFold(frontDV, sideno, -0.8, 0)
                inletDV.CaptionDefinitionTextPrimary = "Inlet"
                DVlist.Add(inletDV)
            Else
                outletDV = objDVs.AddAssemblyView(asmlink, frontno, scalefactor, -0.8, 0, 0)
                outletDV.CaptionDefinitionTextPrimary = "Outlet"
                DVlist.Add(outletDV)

                inletDV = objDVs.AddAssemblyView(asmlink, frontno, scalefactor, -0.8, 0, 0)
                inletDV.CaptionDefinitionTextPrimary = "Inlet"
                DVlist.Add(inletDV)
            End If

            'top view, needs cropping!
            topDV = objDVs.AddByFold(frontDV, 2, -0.8, 0)
            topDV.CaptionDefinitionTextPrimary = "TopView"
            topDV.Update()

            DVlist.Add(topDV)

            'set position of the drawing views, starting with top right
            DVPositionsHor(DVlist, consys, sheetsize, coil, circuit)

            inletDV.Configuration = "Inlet"
            inletDV.MatchConfiguration = True
            General.seapp.DoIdle()

            inletDV.Update()
            General.seapp.DoIdle()

            outletDV.Configuration = "Outlet"
            outletDV.MatchConfiguration = True
            General.seapp.DoIdle()

            outletDV.Update()
            General.seapp.DoIdle()

            isoDV = objDVs.AddByFold(frontDV, isono, 0, 0)

            General.seapp.DoIdle()
            isoDV.Update()

            If General.currentjob.Plant = "Beji" Then
                'switch layer for background
                SwitchLayers(dftdoc, General.currentjob.Plant)
            End If

        Catch ex As Exception
            success = False
            General.CreateLogEntry(ex.ToString)
        End Try

        Return success
    End Function

    Shared Sub FrontViewDims(dftdoc As SolidEdgeDraft.DraftDocument, consys As ConSysData, coil As CoilData, circuit As CircuitData, headertype As String)
        Dim objDVs As SolidEdgeDraft.DrawingViews
        Dim frontDV As SolidEdgeDraft.DrawingView
        Dim objDims As SolidEdgeFrameworkSupport.Dimensions
        Dim finID, SVID As String
        Dim finlines, headerlines, nipplelines, svlines, flangelines As New List(Of SolidEdgeDraft.DVLine2d)
        Dim findimlines(), headerdimlines(), nippledimlines(), flangedimlines() As SolidEdgeDraft.DVLine2d
        Dim x0, y0, scalefactor, headerlength, xs, ys, xe, ye As Double
        Dim searchneeded As Boolean = False
        Dim headerID, nippleID As String

        Try
            objDVs = dftdoc.ActiveSheet.DrawingViews
            frontDV = objDVs(0)
            frontDV.GetOrigin(x0, y0)
            scalefactor = frontDV.ScaleFactor

            If headertype = "outlet" Then
                headerID = "OutletHeader"
                nippleID = "OutletNipple"
            Else
                headerID = "InletHeader"
                nippleID = "InletNipple"
            End If

            finID = "Fin1"

            If circuit.Pressure > 50 Then
                SVID = "0000791913"
            Else
                SVID = "0000107560"
            End If

            If frontDVLinelist.Count = 0 Then
                searchneeded = True
                finlines = SEDrawing.GetLinesFromOcc(frontDV, finID, circuit.ConnectionSide)
                headerlines = SEDrawing.GetLinesFromOcc(frontDV, headerID, circuit.ConnectionSide)
                nipplelines = SEDrawing.GetLinesFromOcc(frontDV, nippleID, circuit.ConnectionSide)
                svlines = SEDrawing.GetLinesFromOcc(frontDV, SVID, circuit.ConnectionSide)
            Else
                Dim tempflist = From partiallist In frontDVLinelist Where partiallist.RefFileName.Contains(finID)

                For Each tempelement In tempflist
                    finlines.Add(tempelement.DVLine)
                Next

                Dim temphlist = From partiallist In frontDVLinelist Where partiallist.RefFileName.Contains(headerID)

                For Each tempelement In temphlist
                    headerlines.Add(tempelement.DVLine)
                Next
                Dim tempnlist = From partiallist In frontDVLinelist Where partiallist.RefFileName.Contains(nippleID)

                For Each tempelement In tempnlist
                    nipplelines.Add(tempelement.DVLine)
                Next

                Dim tempsvlist = From partiallist In frontDVLinelist Where partiallist.RefFileName.Contains(SVID)

                For Each tempelement In tempsvlist
                    svlines.Add(tempelement.DVLine)
                Next
            End If
            Debug.Print(headerlines.Count.ToString)

            findimlines = GetDimLines(finlines, circuit.ConnectionSide, "fin")
            headerdimlines = GetDimLines(headerlines, circuit.ConnectionSide, "header")
            nippledimlines = GetDimLines(nipplelines, circuit.ConnectionSide, "nipple")
            If consys.HasFTCon Then
                nipplelines(0).GetStartPoint(xs, ys)
                nipplelines(0).GetEndPoint(xe, ye)
                flangelines = SEDrawing.GetLinesFromOcc(frontDV, consys.FlangeID, circuit.ConnectionSide, Math.Round(Math.Max(ye, ys), 6))
                flangedimlines = GetDimLines(flangelines, circuit.ConnectionSide, "nipple")
                'will be the same for inlet and outlet, so checking y position
                nippledimlines(1) = flangedimlines(1)
            End If

            frontDV.ScaleFactor = 1
            frontDV.SetOrigin(0, 0)

            If headertype = "outlet" Then
                SetLengthDim(dftdoc.ActiveSheet, frontDV, findimlines(0), "fintop", circuit.ConnectionSide)
                SetLengthDim(dftdoc.ActiveSheet, frontDV, findimlines(1), "finside", circuit.ConnectionSide)
            End If
            SetLengthDim(dftdoc.ActiveSheet, frontDV, headerdimlines(0), "headerbottom", circuit.ConnectionSide)
            objDims = dftdoc.ActiveSheet.Dimensions
            If circuit.CircuitType = "Defrost" Then
                'switch break position
                SwitchBreakPosition(objDims.Item(objDims.Count))
            End If

            headerdimlines(1) = GetHeaderTopLine(headerlines, headerdimlines(0))
            If inletIDs.Count = 0 Then
                headerlength = SetHeaderLength(dftdoc.ActiveSheet, frontDV, headerdimlines(0), headerdimlines(1), circuit.ConnectionSide, General.currentunit.ModelRangeSuffix)
                If Math.Round(headerlength / coil.FinnedHeight, 2) < 0.6 Then
                    MoveBlock(dftdoc, 0, -0.03)
                End If
            End If

            'control position of dimensions
            If headertype = "outlet" Then
                ControlDimensions(frontDV, dftdoc.ActiveSheet.Dimensions, scalefactor, consys, coil, circuit)
            End If

            SetLengthDim(dftdoc.ActiveSheet, frontDV, nippledimlines(1), "nippleside", circuit.ConnectionSide)

            'Position of SV
            If Not consys.HasFTCon And ((headertype = "outlet" And circuit.CircuitType <> "Defrost") Or (headertype = "inlet" And circuit.CircuitType = "Defrost")) And
                General.currentunit.ModelRangeSuffix <> "AP" And General.currentunit.ModelRangeSuffix <> "CP" Then
                SetSVDistance(dftdoc.ActiveSheet, frontDV, nippledimlines(1), svlines, circuit.ConnectionSide)
            End If

            If General.currentunit.ModelRangeSuffix.Substring(1, 1) = "P" Or circuit.CircuitType = "Defrost" Then
                If headertype = "outlet" Then
                    SetNippleLengthFront(dftdoc.ActiveSheet, frontDV, headerdimlines(1), nippledimlines(1), consys, circuit)
                Else
                    SetNippleLengthFront(dftdoc.ActiveSheet, frontDV, headerdimlines(0), nippledimlines(1), consys, circuit)
                End If
            End If

            Dim headerocclist = From plist In consys.Occlist Where plist.Occname.Contains(headerID)

            If headerocclist.Count > 1 Then
                'only for DX units, so hardcoded part ID is fine
                'move air direction block
                MoveBlock(dftdoc, 0, 0.03)

                'clear headerlines 
                headerlines.Clear()

                If searchneeded Then
                    headerlines = SEDrawing.GetLinesFromOcc(frontDV, "OutletHeader2", circuit.ConnectionSide)
                Else
                    Dim temphlist = From partiallist In frontDVLinelist Where partiallist.RefFileName.Contains("OutletHeader2")

                    For Each tempelement In temphlist
                        headerlines.Add(tempelement.DVLine)
                    Next
                End If
                headerdimlines = GetDimLines(headerlines, circuit.ConnectionSide, "header")
                headerdimlines(1) = GetHeaderTopLine(headerlines, headerdimlines(0))
                SetHeaderLength(dftdoc.ActiveSheet, frontDV, headerdimlines(0), headerdimlines(1), circuit.ConnectionSide, General.currentunit.ModelRangeSuffix)

                'position of header lenght dim aligned with first header dim!
            End If

            frontDV.ScaleFactor = scalefactor
            frontDV.SetOrigin(x0, y0)

            If headertype = "outlet" Then
                HandleVisibility(frontDV, "front", consys)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub SideViewDims(dftdoc As SolidEdgeDraft.DraftDocument, consys As ConSysData, coil As CoilData, circuit As CircuitData)
        Dim objDVs As SolidEdgeDraft.DrawingViews
        Dim sideDV As SolidEdgeDraft.DrawingView
        Dim x0, y0, scalefactor As Double
        Dim headerlines, pipelines As List(Of SolidEdgeDraft.DVLine2d)
        Dim nipplearcs, headerarcs, partarcs As List(Of SolidEdgeDraft.DVArc2d)
        Dim partellipses As List(Of SolidEdgeDraft.DVEllipse2d)
        Dim partcircles As List(Of SolidEdgeDraft.DVCircle2d)
        Dim headerdimlines(), bottomline, topline As SolidEdgeDraft.DVLine2d
        Dim nipplecircs As New List(Of SolidEdgeDraft.DVCircle2d)

        Try
            'only for DX → no inlet!!

            objDVs = dftdoc.ActiveSheet.DrawingViews
            sideDV = objDVs(2)
            sideDV.GetOrigin(x0, y0)
            scalefactor = sideDV.ScaleFactor

            headerlines = SEDrawing.GetLinesFromOcc(sideDV, "OutletHeader", circuit.ConnectionSide)

            sideDV.ScaleFactor = 1
            sideDV.SetOrigin(0, 0)
            headerdimlines = GetDimLines(headerlines, circuit.ConnectionSide, "header")

            bottomline = headerdimlines(0)
            'header frame: defined by top and bottom line, using start and end point
            topline = GetHeaderTopLine(headerlines, bottomline)

            'distance first stutzen row
            If uniqueoutletsdata.First.Angle = 0 Then
                headerarcs = SEDrawing.GetArcsFromOcc(sideDV, "OutletHeader", circuit.CoreTube.Diameter)
                SetStutzenArcDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, headerarcs, scalefactor, circuit.ConnectionSide, coil.FinnedHeight)
            Else
                'must be between 0° and 90°
                partellipses = GetEllipsesFromOcc(sideDV, New List(Of String) From {uniqueoutletsdata.First.ID})
                If partellipses.Count > 0 Then
                    SetStutzenElDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, partellipses, scalefactor, circuit.ConnectionSide, coil.FinnedHeight)
                Else
                    'try arc
                    partarcs = SEDrawing.GetArcsFromOcc(sideDV, uniqueoutletsdata.First.ID, circuit.CoreTube.Diameter)
                    If partarcs.Count > 0 Then
                        SetStutzenArcDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, partarcs, scalefactor, circuit.ConnectionSide, coil.FinnedHeight)
                    End If
                End If
            End If

            HandleVisibility(sideDV, "side", consys)

            'distance second stutzen row
            If uniqueoutletsdata.Count > 1 And uniqueoutletsdata.Last.Angle > 0 Then
                If uniqueoutletsdata.Last.Angle = 90 Then
                    'circles
                    partcircles = SEDrawing.GetCirclesFromOcc(sideDV, New List(Of String) From {uniqueoutletsdata.Last.ID})
                    If partcircles.Count > 0 Then
                        SetStutzenCircDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, partcircles, scalefactor, circuit.ConnectionSide, coil.FinnedHeight)
                    End If
                Else
                    partellipses = GetEllipsesFromOcc(sideDV, New List(Of String) From {uniqueoutletsdata.Last.ID})
                    If partellipses.Count > 0 Then
                        SetStutzenElDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, partellipses, scalefactor, circuit.ConnectionSide, coil.FinnedHeight)
                    End If
                End If
            End If

            'start with position of nipple tube
            If consys.HeaderMaterial = "C" Then
                nipplearcs = SEDrawing.GetArcsFromOcc(sideDV, "OutletNipple", 0)
            Else
                nipplecircs = SEDrawing.GetCirclesFromOcc(sideDV, New List(Of String) From {"OutletNipple"})
            End If

            If circuit.CoreTube.Material.Substring(0, 1) = "C" Then
                pipelines = SEDrawing.GetLinesFromOcc(sideDV, "0000903626", circuit.ConnectionSide)
            Else
                pipelines = SEDrawing.GetLinesFromOcc(sideDV, "0000896642", circuit.ConnectionSide)
            End If

            SetNippleDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, nipplearcs, nipplecircs, circuit.ConnectionSide)

            'position of pressure pipe
            SetPipeDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, pipelines, circuit.ConnectionSide)

            Dim headerocclist = From plist In consys.Occlist Where plist.Occname.Contains("OutletHeader")

            If headerocclist.Count > 1 Then
                headerlines.Clear()
                headerlines = SEDrawing.GetLinesFromOcc(sideDV, "OutletHeader2", circuit.ConnectionSide)
                headerdimlines = GetDimLines(headerlines, circuit.ConnectionSide, "header")

                bottomline = headerdimlines(0)
                topline = GetHeaderTopLine(headerlines, bottomline)

                headerarcs = SEDrawing.GetArcsFromOcc(sideDV, "OutletHeader2", circuit.CoreTube.Diameter)
                If uniqueoutletsdata.Last.Angle = 90 Then
                    SetStutzenCircDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, partcircles, scalefactor, circuit.ConnectionSide, coil.FinnedHeight)
                Else
                    If partellipses.Count > 0 Then
                        SetStutzenElDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, partellipses, scalefactor, circuit.ConnectionSide, coil.FinnedHeight)
                    End If
                End If
                SetStutzenArcDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, headerarcs, scalefactor, circuit.ConnectionSide, coil.FinnedHeight)
                SetNippleDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, nipplearcs, nipplecircs, circuit.ConnectionSide)
                SetPipeDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, pipelines, circuit.ConnectionSide)
            End If

            sideDV.ScaleFactor = scalefactor
            sideDV.SetOrigin(x0, y0)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub TopViewDims(dftdoc As SolidEdgeDraft.DraftDocument, consys As ConSysData, circuit As CircuitData)
        Dim objDVs As SolidEdgeDraft.DrawingViews = dftdoc.ActiveSheet.DrawingViews
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = dftdoc.ActiveSheet.Dimensions
        Dim finline(), bottomfinline, nippleline(), nipplesideline, ctline(), leftctline, rightctline As SolidEdgeDraft.DVLine2d
        Dim finlines, nipplelines, ctlines As List(Of SolidEdgeDraft.DVLine2d)
        Dim headercircles As List(Of SolidEdgeDraft.DVCircle2d)
        Dim headercircle As SolidEdgeDraft.DVCircle2d
        Dim topDV As SolidEdgeDraft.DrawingView
        Dim x1, y1, ctd As Double
        Dim mmlist As New List(Of String)
        Dim keyinfolist, klist As New List(Of KeyInfo)
        Dim finID, topelement As String

        Try
            'only for DX → no inlet!!

            topDV = objDVs.Item(2)
            topDV.GetOrigin(x1, y1)
            ctd = topDV.ScaleFactor

            HandleVisibility(topDV, "top", consys)

            topDV.ScaleFactor = 1
            topDV.SetOrigin(0, 0)
            'it's a cap if used or header if closed
            If consys.OutletHeaders.First.Tube.TopCapID = "" Then
                topelement = "OutletHeader"
            Else
                topelement = consys.OutletHeaders.First.Tube.TopCapID
            End If
            headercircles = SEDrawing.GetCirclesFromOcc(topDV, New List(Of String) From {topelement})
            If headercircles.Count = 1 Then
                headercircle = headercircles.First
            Else
                If headercircles.First.Radius > headercircles.Last.Radius Then
                    headercircle = headercircles.First
                Else
                    headercircle = headercircles.Last
                End If
            End If

            finID = "Fin1"

            finlines = SEDrawing.GetLinesFromOcc(topDV, finID, circuit.ConnectionSide)
            finline = GetDimLines(finlines, circuit.ConnectionSide, "fins")
            bottomfinline = finline(0)

            nipplelines = SEDrawing.GetLinesFromOcc(topDV, "OutletNipple", circuit.ConnectionSide)
            nippleline = GetDimLines(nipplelines, circuit.ConnectionSide, "nipple")
            nipplesideline = nippleline(1)

            ctlines = SEDrawing.GetLinesFromOcc(topDV, General.GetShortName(circuit.CoreTube.FileName), circuit.ConnectionSide)
            ctline = GetCTLines(ctlines, "outlet", circuit.ConnectionSide, consys.OutletHeaders.First)
            leftctline = ctline(0)
            rightctline = ctline(1)

            'distance header first ct row, only when offset
            If circuit.ConnectionSide = "left" Then
                SetHeaderOffset(dftdoc.ActiveSheet, topDV, headercircle, rightctline, circuit.ConnectionSide)
                SetCTOverhang(dftdoc.ActiveSheet, topDV, bottomfinline, leftctline, circuit.ConnectionSide)
            Else
                SetHeaderOffset(dftdoc.ActiveSheet, topDV, headercircle, leftctline, circuit.ConnectionSide)
                SetCTOverhang(dftdoc.ActiveSheet, topDV, bottomfinline, rightctline, circuit.ConnectionSide)
            End If

            'Dim a
            SetHeaderDistance(dftdoc.ActiveSheet, topDV, headercircle, bottomfinline, circuit.ConnectionSide)

            SetNippleLengthTop(dftdoc.ActiveSheet, topDV, headercircle, nipplesideline, circuit.ConnectionSide)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        Finally
            topDV.ScaleFactor = ctd
            topDV.SetOrigin(x1, y1)
        End Try


    End Sub

    Shared Sub FPSideViewDims(dftdoc As SolidEdgeDraft.DraftDocument, consys As ConSysData, sideDV As SolidEdgeDraft.DrawingView, header As HeaderData, circuit As CircuitData, coil As CoilData)
        Dim headerlines As List(Of SolidEdgeDraft.DVLine2d)
        Dim headerdimlines(), bottomline, topline As SolidEdgeDraft.DVLine2d
        Dim headerarcs, nipplearcs As List(Of SolidEdgeDraft.DVArc2d)
        Dim partellipses, ventellipses As List(Of SolidEdgeDraft.DVEllipse2d)
        Dim partcircles As List(Of SolidEdgeDraft.DVCircle2d)
        Dim ventarcs As New List(Of SolidEdgeDraft.DVArc2d)
        Dim nipplecircs As New List(Of SolidEdgeDraft.DVCircle2d)
        Dim x0, y0, scalefactor, ctdiameter As Double
        Dim ventID As String = ""
        Dim headerID As String = header.Tube.HeaderType.Substring(0, 1).ToUpper + header.Tube.HeaderType.Substring(1) + "Header"

        Try
            sideDV.GetOrigin(x0, y0)
            scalefactor = sideDV.ScaleFactor

            headerlines = SEDrawing.GetLinesFromOcc(sideDV, headerID, circuit.ConnectionSide)

            sideDV.ScaleFactor = 1
            sideDV.SetOrigin(0, 0)
            headerdimlines = GetDimLines(headerlines, circuit.ConnectionSide, "header")

            bottomline = headerdimlines(0)

            'header frame: defined by top and bottom line, using start and end point
            topline = GetHeaderTopLine(headerlines, bottomline)

            'set headerlength
            SetHeaderLength(dftdoc.ActiveSheet, sideDV, bottomline, topline, circuit.ConnectionSide, "FP")

            'set nippletube dim
            nipplearcs = SEDrawing.GetArcsFromOcc(sideDV, "letNipple", 0)
            SetNippleDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, nipplearcs, nipplecircs, circuit.ConnectionSide)

            'for outlet dim drain
            If header.VentIDs IsNot Nothing Then
                ventID = header.VentIDs(2)
            End If

            If ventID IsNot Nothing Then
                If header.Tube.HeaderType = "outlet" Then
                    ventarcs = SEDrawing.GetArcsFromOcc(sideDV, "OutletHeader", GNData.GetVentDiameter(header.Tube.Materialcodeletter, header.Ventsize))
                End If

                ventellipses = GetEllipsesFromOcc(sideDV, New List(Of String) From {ventID})
                If ventellipses.Count = 0 And ventarcs.Count = 0 Then
                    ventarcs = SEDrawing.GetArcsFromOcc(sideDV, ventID, 0)
                    If ventarcs.Count = 0 Then
                        nipplecircs = SEDrawing.GetCirclesFromOcc(sideDV, New List(Of String) From {ventID})
                    End If
                End If
            End If

            If (header.Tube.HeaderType = "outlet" Or (Math.Round(bottomline.Length * 1000, 2) < 60 And consys.HeaderMaterial = "C") Or ((circuit.FinType = "N" Or circuit.FinType = "M") And consys.HeaderMaterial <> "C")) And circuit.Pressure < 17 Then
                SetVentDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, ventellipses, ventarcs, nipplecircs, circuit.ConnectionSide)
            End If

            Dim uniqueangles As New List(Of Double)
            For i As Integer = 0 To header.StutzenDatalist.Count - 1
                If uniqueangles.IndexOf(Math.Abs(header.StutzenDatalist(i).Angle)) = -1 Then
                    uniqueangles.Add(Math.Abs(header.StutzenDatalist(i).Angle))
                End If
            Next

            For i As Integer = 0 To uniqueangles.Count - 1
                Dim slist As New List(Of String)
                For j As Integer = 0 To header.StutzenDatalist.Count - 1
                    If Math.Abs(header.StutzenDatalist(j).Angle) = uniqueangles(i) Then
                        slist.Add(header.StutzenDatalist(j).ID)
                    End If
                Next

                Select Case uniqueangles(i)
                    Case 0
                        If circuit.CoreTube.Materialcodeletter = "V" Or circuit.CoreTube.Materialcodeletter = "W" Then
                            ctdiameter = circuit.CoreTube.Diameter + 1
                        Else
                            ctdiameter = circuit.CoreTube.Diameter
                        End If
                        headerarcs = SEDrawing.GetArcsFromOcc(sideDV, headerID, ctdiameter)
                        SetStutzenArcDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, headerarcs, scalefactor, circuit.ConnectionSide, coil.FinnedHeight)
                    Case 90
                        'use circles
                        partcircles = SEDrawing.GetCirclesFromOcc(sideDV, General.GetUniqueStrings(slist))
                        If partcircles.Count > 0 Then
                            SetStutzenCircDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, partcircles, scalefactor, circuit.ConnectionSide, coil.FinnedHeight)
                        End If
                    Case Else
                        'use ellipses
                        partellipses = GetEllipsesFromOcc(sideDV, General.GetUniqueStrings(slist))
                        If partellipses.Count > 0 Then
                            SetStutzenElDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, partellipses, scalefactor, circuit.ConnectionSide, coil.FinnedHeight)
                        End If
                End Select
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        Finally
            sideDV.ScaleFactor = scalefactor
            sideDV.SetOrigin(x0, y0)

            sideDV.CaptionLocation = SolidEdgeFrameworkSupport.DimViewCaptionLocationConstants.igDimViewCaptionLocationTop
            sideDV.DisplayCaption = True

            sideDV.GetCaptionPosition(x0, y0)
            sideDV.SetCaptionPosition(x0, y0 + 0.013)
        End Try

    End Sub

    Shared Sub BrineSideViewDims(dftdoc As SolidEdgeDraft.DraftDocument, consys As ConSysData, sideDV As SolidEdgeDraft.DrawingView, header As HeaderData, circuit As CircuitData, coil As CoilData)
        Dim headerlines As List(Of SolidEdgeDraft.DVLine2d)
        Dim headerdimlines(), bottomline, topline As SolidEdgeDraft.DVLine2d
        Dim headerarcs As List(Of SolidEdgeDraft.DVArc2d)
        Dim partellipses As List(Of SolidEdgeDraft.DVEllipse2d)
        Dim partcircles As List(Of SolidEdgeDraft.DVCircle2d)
        Dim nipplecircs As List(Of SolidEdgeDraft.DVCircle2d)
        Dim objDims As SolidEdgeFrameworkSupport.Dimensions
        Dim x0, y0, scalefactor As Double
        Dim ventID As String = ""
        Dim comparer, notconside As String
        Dim headerID As String = header.Tube.HeaderType.Substring(0, 1).ToUpper + header.Tube.HeaderType.Substring(1) + "Header"

        Try
            objDims = dftdoc.ActiveSheet.Dimensions
            If circuit.ConnectionSide = "left" Then
                comparer = "smaller"
                notconside = "right"
            Else
                comparer = "bigger"
                notconside = "left"
            End If

            sideDV.GetOrigin(x0, y0)
            scalefactor = sideDV.ScaleFactor

            headerlines = SEDrawing.GetLinesFromOcc(sideDV, headerID, circuit.ConnectionSide)

            sideDV.ScaleFactor = 1
            sideDV.SetOrigin(0, 0)
            headerdimlines = GetDimLines(headerlines, circuit.ConnectionSide, "header")

            bottomline = headerdimlines(0)

            'header frame: defined by top and bottom line, using start and end point
            topline = GetHeaderTopLine(headerlines, bottomline)

            'set headerlength
            SetHeaderLength(dftdoc.ActiveSheet, sideDV, bottomline, topline, circuit.ConnectionSide, "FP")
            SEDrawing.ControlDimDistance(objDims.Item(objDims.Count), {"x", comparer}, 0.02)

            'set nippletube dim
            nipplecircs = SEDrawing.GetCirclesFromOcc(sideDV, New List(Of String) From {"letNipple"})
            SetNippleDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, Nothing, nipplecircs, circuit.ConnectionSide)
            SEDrawing.ControlDimDistance(objDims.Item(objDims.Count), {"x", comparer}, 0.015)

            If header.VentIDs.Count > 0 Then
                ventID = header.VentIDs(2)
            End If

            If ventID IsNot Nothing Then
                If header.Tube.HeaderType = "outlet" Then
                    partcircles = SEDrawing.GetCirclesFromOcc(sideDV, New List(Of String) From {ventID})

                    SetBrineVentDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, partcircles.First, circuit.ConnectionSide)
                    SEDrawing.ControlDimDistance(objDims.Item(objDims.Count), {"x", comparer})
                Else
                    headerarcs = SEDrawing.GetArcsFromOcc(sideDV, headerID, 0)
                    Dim ventarc As SolidEdgeDraft.DVArc2d = FindVentArc(headerarcs, circuit.ConnectionSide)
                    SetBrineVentDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, ventarc, circuit.ConnectionSide)
                    SEDrawing.ControlDimDistance(objDims.Item(objDims.Count), {"x", comparer})
                End If
            End If

            Dim uniqueangles As New List(Of Double)
            For i As Integer = 0 To header.StutzenDatalist.Count - 1
                If uniqueangles.IndexOf(Math.Abs(header.StutzenDatalist(i).Angle)) = -1 Then
                    uniqueangles.Add(Math.Abs(header.StutzenDatalist(i).Angle))
                End If
            Next

            For i As Integer = 0 To uniqueangles.Count - 1
                Dim slist As New List(Of String)
                For j As Integer = 0 To header.StutzenDatalist.Count - 1
                    If Math.Abs(header.StutzenDatalist(j).Angle) = uniqueangles(i) Then
                        slist.Add(header.StutzenDatalist(j).ID)
                    End If
                Next

                Select Case uniqueangles(i)
                    Case 0
                        headerarcs = SEDrawing.GetArcsFromOcc(sideDV, headerID, circuit.CoreTube.Diameter)
                        SetStutzenArcDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, headerarcs, scalefactor, notconside, coil.FinnedHeight)
                    Case 90
                        'use circles
                        partcircles = SEDrawing.GetCirclesFromOcc(sideDV, General.GetUniqueStrings(slist))
                        If partcircles.Count > 0 Then
                            SetStutzenCircDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, partcircles, scalefactor, notconside, coil.FinnedHeight)
                        End If
                    Case Else
                        'use ellipses
                        partellipses = GetEllipsesFromOcc(sideDV, General.GetUniqueStrings(slist))
                        If partellipses.Count > 0 Then
                            SetStutzenElDistance(dftdoc.ActiveSheet, sideDV, bottomline, topline, partellipses, scalefactor, circuit.ConnectionSide, coil.FinnedHeight)
                        End If
                End Select
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        Finally
            sideDV.ScaleFactor = scalefactor
            sideDV.SetOrigin(x0, y0)

            sideDV.CaptionLocation = SolidEdgeFrameworkSupport.DimViewCaptionLocationConstants.igDimViewCaptionLocationTop
            sideDV.DisplayCaption = True

            sideDV.GetCaptionPosition(x0, y0)
            sideDV.SetCaptionPosition(x0, y0 + 0.013)
        End Try

    End Sub

    Shared Sub FPTopViewDims(dftdoc As SolidEdgeDraft.DraftDocument, topDV As SolidEdgeDraft.DrawingView, consys As ConSysData, header As HeaderData, circuit As CircuitData, coil As CoilData)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = dftdoc.ActiveSheet.Dimensions
        Dim finline(), bottomfinline, nippleline(), nipplesideline, ctline(), leftctline, rightctline, ventline As SolidEdgeDraft.DVLine2d
        Dim finlines, nipplelines, ctlines, ventlines As List(Of SolidEdgeDraft.DVLine2d)
        Dim headercircles As List(Of SolidEdgeDraft.DVCircle2d)
        Dim headercircle As SolidEdgeDraft.DVCircle2d
        Dim x1, y1, ctd As Double
        Dim finID, topviewelement, nippleID, headerID, comparer As String
        Dim nippletube As TubeData

        Try

            topDV.GetOrigin(x1, y1)
            ctd = topDV.ScaleFactor

            If header.Tube.HeaderType = "outlet" Then
                If header.Tube.TopCapID = "" Then
                    topviewelement = "OutletHeader"
                Else
                    topviewelement = header.Tube.TopCapID
                End If
                nippleID = "OutletNipple"
                headerID = "OutletHeader"
                nippletube = consys.OutletNipples.First
            Else
                If header.Tube.TopCapID = "" Then
                    topviewelement = "InletHeader"
                Else
                    topviewelement = header.Tube.TopCapID
                End If
                nippleID = "InletNipple"
                headerID = "InletHeader"
                nippletube = consys.InletNipples.First
            End If

            topDV.ScaleFactor = 1
            topDV.SetOrigin(0, 0)

            'it's a cap if used or header if closed
            headercircles = SEDrawing.GetCirclesFromOcc(topDV, New List(Of String) From {topviewelement})
            If headercircles.Count = 1 Then
                headercircle = headercircles.First
            ElseIf topviewelement.Contains("Header") Then
                If headercircles.First.Radius > headercircles.Last.Radius Then
                    headercircle = headercircles.First
                Else
                    headercircle = headercircles.Last
                End If
            Else
                'try header circles even if they are hidden
                headercircles = SEDrawing.GetCirclesFromOcc(topDV, New List(Of String) From {headerID})
                If headercircles.Count = 1 Then
                    headercircle = headercircles.First
                Else
                    If headercircles.First.Radius > headercircles.Last.Radius Then
                        headercircle = headercircles.First
                    Else
                        headercircle = headercircles.Last
                    End If
                End If
            End If

            finID = "Fin1"

            finlines = SEDrawing.GetLinesFromOcc(topDV, finID, circuit.ConnectionSide)
            finline = GetDimLines(finlines, circuit.ConnectionSide, "fins")
            bottomfinline = finline(0)

            nipplelines = SEDrawing.GetLinesFromOcc(topDV, nippleID, circuit.ConnectionSide)
            nippleline = GetDimLines(nipplelines, circuit.ConnectionSide, "nipple")
            nipplesideline = nippleline(1)

            If circuit.CircuitType = "Defrost" Then
                ctlines = SEDrawing.GetLinesFromOcc(topDV, "Support", circuit.ConnectionSide)
                If circuit.ConnectionSide = "left" Then
                    comparer = "smaller"
                Else
                    comparer = "bigger"
                End If
            Else
                ctlines = SEDrawing.GetLinesFromOcc(topDV, "Coretube", circuit.ConnectionSide)
            End If

            'ct overhang only if outlet 
            If header.Tube.HeaderType = "outlet" Then
                ctline = GetCTLines(ctlines, "outlet", circuit.ConnectionSide, header)
                leftctline = ctline(0)
                rightctline = ctline(1)
                If circuit.CircuitType = "Defrost" Then
                    SetCTOverhang(dftdoc.ActiveSheet, topDV, bottomfinline, leftctline, circuit.ConnectionSide)
                    SEDrawing.ControlDimDistance(objdims.Item(objdims.Count), {"x", comparer}, switchaxisdir:=True)
                Else
                    SetCTOverhang(dftdoc.ActiveSheet, topDV, bottomfinline, leftctline, "right")
                    If circuit.ConnectionSide = "left" Then
                        SetHeaderOffset(dftdoc.ActiveSheet, topDV, headercircle, rightctline, circuit.ConnectionSide)
                    Else
                        SetHeaderOffset(dftdoc.ActiveSheet, topDV, headercircle, leftctline, circuit.ConnectionSide)
                    End If
                End If
            Else
                If (headercircle.Diameter > 0.054 Or (circuit.FinType = "F" And consys.HeaderMaterial <> "C")) And circuit.CircuitType <> "Defrost" Then
                    'vent distance
                    ventlines = SEDrawing.GetLinesFromOcc(topDV, header.VentIDs(2), circuit.ConnectionSide)
                    ventline = GetVentLine(ventlines, circuit.ConnectionSide, bottomfinline)
                    If ventline IsNot Nothing Then
                        FPVentDistance(dftdoc.ActiveSheet, topDV, ventline, bottomfinline, circuit.ConnectionSide)
                    End If
                End If

                If circuit.CircuitType <> "Defrost" Then
                    ctline = GetCTLines(ctlines, "inlet", circuit.ConnectionSide, header)
                    SetHeaderOffset(dftdoc.ActiveSheet, topDV, headercircle, ctline(0), circuit.ConnectionSide)
                End If
            End If

            'Dim a
            If header.Tube.Diameter = nippletube.Diameter Then
                If circuit.CircuitType = "Defrost" Then
                    Dim notconside As String = "right"
                    If circuit.ConnectionSide = "right" Then
                        notconside = "left"
                    End If
                    FPSetHeaderDistance(dftdoc.ActiveSheet, topDV, nippleline(0), bottomfinline, headercircle, notconside, header.Tube.HeaderType, circuit.FinType)
                    If header.Tube.HeaderType = "outlet" And circuit.ConnectionSide = "left" Then
                        objdims.Item(objdims.Count - 1).TrackDistance *= 0.6
                        SEDrawing.ControlDimDistance(objdims.Item(objdims.Count - 1), {"x", comparer})
                        objdims.Item(objdims.Count).TrackDistance *= 0.6
                        SEDrawing.ControlDimDistance(objdims.Item(objdims.Count), {"x", comparer})
                    End If
                Else
                    FPSetHeaderDistance(dftdoc.ActiveSheet, topDV, nippleline(0), bottomfinline, headercircle, circuit.ConnectionSide, header.Tube.HeaderType, circuit.FinType)
                End If
            Else
                SetHeaderDistance(dftdoc.ActiveSheet, topDV, headercircle, bottomfinline, circuit.ConnectionSide)
            End If

            topDV.ScaleFactor = ctd
            topDV.SetOrigin(x1, y1)

            If header.Tube.HeaderType = "inlet" Then
                HandleVisibility(topDV, "top", consys)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub HorFrontViewDims(dftdoc As SolidEdgeDraft.DraftDocument, consys As ConSysData, circuit As CircuitData)
        Dim frontDV As SolidEdgeDraft.DrawingView
        Dim finlines, inheaderlines, outheaderlines As List(Of SolidEdgeDraft.DVLine2d)
        Dim findimlines(), headerdimlines() As SolidEdgeDraft.DVLine2d
        Dim x0, y0, scalefactor As Double
        Dim bottomdistance, topdistance, finneddim As SolidEdgeFrameworkSupport.Dimension
        Dim comparer As String

        Try
            If circuit.ConnectionSide = "left" Then
                comparer = "smaller"
            Else
                comparer = "bigger"
            End If

            frontDV = dftdoc.ActiveSheet.DrawingViews.Item(1)
            scalefactor = frontDV.ScaleFactor
            frontDV.GetOrigin(x0, y0)

            finlines = SEDrawing.GetLinesFromOcc(frontDV, "Fin1", circuit.ConnectionSide)
            inheaderlines = SEDrawing.GetLinesFromOcc(frontDV, "InletHeader", circuit.ConnectionSide)
            outheaderlines = SEDrawing.GetLinesFromOcc(frontDV, "OutletHeader", circuit.ConnectionSide)
            findimlines = GetDimLines(finlines, circuit.ConnectionSide, "fin")

            frontDV.ScaleFactor = 1
            frontDV.SetOrigin(0, 0)

            finneddim = SetLengthDim(dftdoc.ActiveSheet, frontDV, findimlines(0), "fintop", circuit.ConnectionSide)
            SEDrawing.ControlDimDistance(finneddim, {"y", "bigger"})
            finneddim = SetLengthDim(dftdoc.ActiveSheet, frontDV, findimlines(1), "finside", circuit.ConnectionSide)
            SEDrawing.ControlDimDistance(finneddim, {"x", comparer})
            'inlet
            headerdimlines = SEDrawing.GetHorHeaderDimLines(inheaderlines, consys.InletHeaders.First.Tube.Diameter)
            If headerdimlines(0) IsNot Nothing AndAlso headerdimlines(1) IsNot Nothing Then
                'diameter & length
                HorHeaderDim(dftdoc.ActiveSheet, frontDV, headerdimlines, circuit)

                'distance to bottom
                bottomdistance = HorHeaderDistance(dftdoc.ActiveSheet, frontDV, headerdimlines, circuit.ConnectionSide, findimlines(1), "bottom")
            End If

            'outlet
            headerdimlines = SEDrawing.GetHorHeaderDimLines(outheaderlines, consys.OutletHeaders.First.Tube.Diameter)
            If headerdimlines(0) IsNot Nothing AndAlso headerdimlines(1) IsNot Nothing Then
                'diameter & length
                HorHeaderDim(dftdoc.ActiveSheet, frontDV, headerdimlines, circuit)

                'distance to top
                topdistance = HorHeaderDistance(dftdoc.ActiveSheet, frontDV, headerdimlines, circuit.ConnectionSide, findimlines(1), "top")
            End If

            frontDV.ScaleFactor = scalefactor
            frontDV.SetOrigin(x0, y0)

            'check dim trackdistance
            If bottomdistance IsNot Nothing Then
                SEDrawing.ControlDimDistance(bottomdistance, {"x", comparer})
            End If
            If topdistance IsNot Nothing Then
                SEDrawing.ControlDimDistance(topdistance, {"x", comparer})
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub HorHeaderViews(dftdoc As SolidEdgeDraft.DraftDocument, DV As SolidEdgeDraft.DrawingView, header As HeaderData, circuit As CircuitData, IDlist As List(Of String), sdatalist As List(Of StutzenData))
        Dim headerlines As List(Of SolidEdgeDraft.DVLine2d)
        Dim headerdimlines() As SolidEdgeDraft.DVLine2d
        Dim stutzencircs As List(Of SolidEdgeDraft.DVCircle2d)
        Dim headerframe(), scalefactor, x0, y0 As Double

        Try

            scalefactor = DV.ScaleFactor
            DV.GetOrigin(x0, y0)

            DV.ScaleFactor = 1
            DV.SetOrigin(0, 0)
            'get vertical headerlines to create a frame
            headerlines = SEDrawing.GetLinesFromOcc(DV, header.Tube.TubeFile.Shortname, circuit.ConnectionSide)

            'left / right vertical line
            headerdimlines = SEDrawing.GetHorHeaderDimLines(headerlines, header.Tube.Diameter)

            headerframe = SEDrawing.GetHeaderFrame(headerdimlines)
            For j As Integer = 0 To IDlist.Count - 1
                stutzencircs = SEDrawing.GetCirclesFromOcc(DV, New List(Of String) From {IDlist(j)})
                Dim figure As Integer

                'if figure 4 → circ has to be inside the header frame
                For i As Integer = 0 To sdatalist.Count - 1
                    If sdatalist(i).ID = IDlist(j) Then
                        figure = sdatalist(i).Figure
                    End If
                Next

                Dim sortedcircs As List(Of SolidEdgeDraft.DVCircle2d) = StutzenHorPosition(dftdoc.ActiveSheet, DV, headerdimlines, headerframe, stutzencircs,
                                                                                           circuit, header.Tube.HeaderType, sdatalist.Count, figure)

                If figure < 8 Then
                    'vertical distance to header center
                    HeaderStutzenGap(dftdoc.ActiveSheet, DV, headerdimlines, stutzencircs, circuit)
                End If
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        Finally
            DV.ScaleFactor = scalefactor
            DV.SetOrigin(x0, y0)

            DV.CaptionLocation = SolidEdgeFrameworkSupport.DimViewCaptionLocationConstants.igDimViewCaptionLocationTop
            DV.DisplayCaption = True

            DV.GetCaptionPosition(x0, y0)
            DV.SetCaptionPosition(x0, y0 + 0.013)
        End Try

    End Sub

    Shared Sub TopViewHor(dftdoc As SolidEdgeDraft.DraftDocument, consys As ConSysData, circuit As CircuitData)
        Dim topDV As SolidEdgeDraft.DrawingView = dftdoc.ActiveSheet.DrawingViews.Item(4)
        Dim finlines, outheaderlines, ctlines As List(Of SolidEdgeDraft.DVLine2d)
        Dim inheaderlines As New List(Of SolidEdgeDraft.DVLine2d)
        Dim finline(), ctline(), outheaderline(), inheaderline(1) As SolidEdgeDraft.DVLine2d
        Dim x0, y0, scalefactor As Double
        Dim comparer As String
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions
        Dim headerdim As SolidEdgeFrameworkSupport.Dimension

        Try

            topDV.GetOrigin(x0, y0)
            scalefactor = topDV.ScaleFactor

            topDV.ScaleFactor = 1
            topDV.SetOrigin(0, 0)
            finlines = SEDrawing.GetLinesFromOcc(topDV, "Fin1", circuit.ConnectionSide)
            finline = GetDimLines(finlines, circuit.ConnectionSide, "fins")
            outheaderlines = SEDrawing.GetLinesFromOcc(topDV, "OutletHeader", circuit.ConnectionSide)
            outheaderline = SEDrawing.GetHorHeaderDimLines(outheaderlines, consys.OutletHeaders.First.Tube.Diameter)
            ctlines = SEDrawing.GetLinesFromOcc(topDV, General.GetShortName(circuit.CoreTube.FileName), circuit.ConnectionSide)
            ctline = GetCTLinesHor(ctlines)

            'coretube overhang & dim a outlet
            If circuit.ConnectionSide = "left" Then
                SetCTOverhang(dftdoc.ActiveSheet, topDV, finline(0), ctline(0), circuit.ConnectionSide)
                SetHeaderDistanceHor(dftdoc.ActiveSheet, topDV, outheaderline(1), finline(1), circuit.ConnectionSide, False, False)
            Else
                SetCTOverhang(dftdoc.ActiveSheet, topDV, finline(0), ctline(1), circuit.ConnectionSide)
                SetHeaderDistanceHor(dftdoc.ActiveSheet, topDV, outheaderline(0), finline(0), circuit.ConnectionSide, False, False)
            End If

            objdims = dftdoc.ActiveSheet.Dimensions

            'if dim a inlet <> dim a outlet → dim a inlet too
            If consys.InletHeaders.First.Dim_a <> consys.OutletHeaders.First.Dim_a Then
                inheaderlines = SEDrawing.GetLinesFromOcc(topDV, "InletHeader", circuit.ConnectionSide)
                inheaderline = SEDrawing.GetHorHeaderDimLines(inheaderlines, consys.InletHeaders.First.Tube.Diameter)
                If inheaderline(0) IsNot Nothing Then
                    If circuit.ConnectionSide = "left" Then
                        SetHeaderDistanceHor(dftdoc.ActiveSheet, topDV, inheaderline(0), finline(0), "right", onlya:=True, False)
                    Else
                        SetHeaderDistanceHor(dftdoc.ActiveSheet, topDV, inheaderline(0), finline(0), "left", onlya:=True, False)
                    End If
                End If
            End If

            If consys.InletHeaders.First.Tube.Diameter <> consys.OutletHeaders.First.Tube.Diameter Then
                If inheaderlines.Count = 0 Then
                    inheaderlines = SEDrawing.GetLinesFromOcc(topDV, "InletHeader", circuit.ConnectionSide)
                    inheaderline = SEDrawing.GetHorHeaderDimLines(inheaderlines, consys.InletHeaders.First.Tube.Diameter)
                End If
                If inheaderline(0) Is Nothing Then
                    'use SV line
                    inheaderlines = SEDrawing.GetLinesFromOcc(topDV, GNData.GetSVID(circuit.Pressure), circuit.ConnectionSide)
                    inheaderline = SEDrawing.GetHorHeaderDimLines(inheaderlines, 8)
                End If
                If circuit.ConnectionSide = "left" Then
                    headerdim = SetHeaderDistanceHor(dftdoc.ActiveSheet, topDV, inheaderline(0), finline(0), "right", False, onlycenter:=True)
                    comparer = "bigger"
                Else
                    headerdim = SetHeaderDistanceHor(dftdoc.ActiveSheet, topDV, inheaderline(0), finline(0), "left", False, onlycenter:=True)
                    comparer = "smaller"
                End If
                SEDrawing.ControlDimDistance(headerdim, {"x", comparer})
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        Finally
            topDV.ScaleFactor = scalefactor
            topDV.SetOrigin(x0, y0)
        End Try

    End Sub

    Shared Sub IsoViewHor(dftdoc As SolidEdgeDraft.DraftDocument, consys As ConSysData, circuit As CircuitData)


    End Sub

    Shared Sub IsoView(dftdoc As SolidEdgeDraft.DraftDocument, consys As ConSysData, circuit As CircuitData, coil As CoilData)
        Dim objDVs As SolidEdgeDraft.DrawingViews = dftdoc.ActiveSheet.DrawingViews
        Dim isoDV As SolidEdgeDraft.DrawingView = objDVs.Item(objDVs.Count)
        Dim x1, y1, ctd, cleft, cright, ctop, cbottom, x0, y0, offset, fpmfactor As Double
        Dim xlist, ylist As New List(Of Double)
        Dim mindex1 As Integer
        Dim mmlist As New List(Of String)
        Dim keyinfolist, klist As New List(Of KeyInfo)
        Dim DVEllipses As SolidEdgeDraft.DVEllipses2d
        Dim ellipselist As New List(Of SolidEdgeDraft.DVEllipse2d)
        Dim DVLines As SolidEdgeDraft.DVLines2d
        Dim linelist As New List(Of SolidEdgeDraft.DVLine2d)
        Dim isoelement As String

        Try
            ctd = isoDV.ScaleFactor
            If General.currentunit.ModelRangeSuffix.Substring(1, 1) = "P" Then
                ctd *= 0.5
            End If

            fpmfactor = 0
            isoDV.ScaleFactor = 1
            isoDV.SetOrigin(0, 0)

            isoDV.Update()
            isoDV.Crop = False
            isoDV.Update()
            cleft = isoDV.CropLeft
            cright = isoDV.CropRight
            ctop = isoDV.CropTop
            cbottom = isoDV.CropBottom

            If circuit.ConnectionSide = "left" Then
                'isoDV.CropRight = 0
                isoelement = "Fin1"
            Else
                'isoDV.CropLeft = 0
                If consys.HeaderAlignment = "horizontal" Then
                    isoelement = "InletHeader"
                ElseIf consys.HasFTCon Then
                    isoelement = consys.FlangeID
                ElseIf circuit.IsOnebranchEvap Then
                    isoelement = "Fin1"
                Else
                    isoelement = "OutletNipple"
                End If
            End If

            If circuit.ConnectionSide = "right" Then
                DVEllipses = isoDV.DVEllipses2d
                For i As Integer = 1 To DVEllipses.Count
                    If DVEllipses.Item(i).ModelMember.FileName.Contains(isoelement) Then
                        ellipselist.Add(DVEllipses.Item(i))
                        mindex1 += 1
                    End If
                Next

                For i As Integer = 0 To ellipselist.Count - 1
                    ellipselist(i).GetCenterPoint(x1, y1)
                    xlist.Add(Math.Round(x1, 6))
                    ylist.Add(Math.Round(y1, 6))
                Next
            Else
                DVLines = isoDV.DVLines2d
                For i As Integer = 1 To DVLines.Count
                    If DVLines.Item(i).ModelMember.FileName.Contains(isoelement) Then
                        linelist.Add(DVLines.Item(i))
                        mindex1 += 1
                    End If
                Next

                For i As Integer = 0 To linelist.Count - 1
                    linelist(i).GetStartPoint(x1, y1)
                    xlist.Add(Math.Round(x1, 6))
                    ylist.Add(Math.Round(y1, 6))
                Next
                If circuit.Pressure < 17 Then
                    fpmfactor = Math.Abs(objDVs.Item(1).CropLeft) + Math.Abs(objDVs.Item(1).CropRight)
                End If
            End If
            x0 = -xlist.Min

            y0 = cbottom
            isoDV.ScaleFactor = ctd

            If coil.FinnedHeight < 500 Then
                offset = 0.0225
            Else
                offset = 0.015
            End If
            x1 = 2 * x0 * ctd + 0.025 + fpmfactor
            Debug.Print("Iso x: " + Math.Round(x1, 6).ToString)
            isoDV.SetOrigin(2 * x0 * ctd + 0.025 + fpmfactor, y0 * ctd + offset)

        Catch ex As Exception
            isoDV.ScaleFactor = ctd
            General.CreateLogEntry(ex.ToString)
        Finally

        End Try

        isoDV.Update()

    End Sub

    Shared Sub Partlist(dftdoc As SolidEdgeDraft.DraftDocument, circuit As CircuitData, headeralignment As String)
        Dim objDVs As SolidEdgeDraft.DrawingViews = dftdoc.ActiveSheet.DrawingViews
        Dim isoDV, centerDV As SolidEdgeDraft.DrawingView
        Dim objPartLists As SolidEdgeDraft.PartsLists
        Dim partlist As SolidEdgeDraft.PartsList
        Dim objSheets As SolidEdgeDraft.Sheets
        Dim objBalloons As SolidEdgeFrameworkSupport.Balloons
        Dim x1list, y1list, x2list, y2list As New List(Of Double)
        Dim x1, y1, x2, y2, newx, newy, cx, cy, bscale, xmin, ymin, xmax, ymax As Double
        Dim proptext, postext As String
        Dim objShapes As SolidEdgeFrameworkSupport.AnnotAlignmentShapes
        Dim objColumn As SolidEdgeDraft.TableColumn

        Try
            isoDV = objDVs.Item(objDVs.Count)
            isoDV.GetOrigin(newx, newy)
            centerDV = objDVs.Item(1)
            centerDV.GetOrigin(cx, cy)

            objPartLists = dftdoc.PartsLists
            objSheets = dftdoc.Sheets

            If objPartLists.Count > 0 Then
                partlist = objPartLists.Item(1)
            Else
                partlist = objPartLists.Add(isoDV, "ISO", AutoBalloon:=1, CreatePartsList:=1)
                partlist.SetComponentSortPriority(SolidEdgeDraft.PartsListComponentType.igPartsListComponentType_FrameMembers, 6)
                partlist.RenumberAccordingToSortOrder = True
                partlist.ItemNumberIncrement = 5
                partlist.ItemNumberStart = 5
                isoDV.Update()
                partlist.Update()
                If circuit.Pressure < 17 And circuit.ConnectionSide = "left" And headeralignment = "vertical" And circuit.CircuitType <> "Defrost" Then
                    partlist.AnchorPoint = SolidEdgeDraft.TableAnchorPoint.igLowerLeft
                Else
                    partlist.AnchorPoint = SolidEdgeDraft.TableAnchorPoint.igLowerRight
                End If
                partlist.FillEndOfTableWithBlankRows = False
                partlist.ShowTopAssembly = False
                If dftdoc.ActiveSheet.Name = "A3" Then
                    bscale = 1
                    partlist.SetOrigin(0.235, 0.005)
                    isoDV.SetOrigin(newx, cy)
                    newy = cy
                Else
                    bscale = 2
                    If circuit.Pressure < 17 And circuit.ConnectionSide = "left" And headeralignment = "vertical" And circuit.CircuitType <> "Defrost" Then
                        partlist.SetOrigin(0.02, 0.005)
                    Else
                        partlist.SetOrigin(0.409, 0.005)
                    End If
                End If
                partlist.Update()
                partlist.ListType = SolidEdgeDraft.PartsListType.igExploded
            End If

            partlist.Update()
            objBalloons = dftdoc.ActiveSheet.Balloons

            For Each sBalloon As SolidEdgeFrameworkSupport.Balloon In objBalloons
                sBalloon.DisplayItemCount = False
                sBalloon.Range(x1, y1, x2, y2)
                sBalloon.TextScale = bscale
                x1list.Add(Math.Round(x1, 6))
                y1list.Add(Math.Round(y1, 6))
                x2list.Add(Math.Round(x2, 6))
                y2list.Add(Math.Round(y2, 6))
            Next
            If x1list.Min < 0.025 Then
                newx += 0.025 + Math.Abs(x1list.Min)
            End If
            If y1list.Min < 0.008 Then
                newy += 0.008 + Math.Abs(y1list.Min)
            End If
            If circuit.Pressure < 17 And circuit.ConnectionSide = "left" And x1list.Min < 0.15 And headeralignment = "vertical" And circuit.CircuitType <> "Defrost" Then
                newx += 0.15 - x1list.Min
            End If
            isoDV.SetOrigin(newx, newy)

            partlist.Columns.Item(1).DataHorizontalAlignment = SolidEdgeFrameworkSupport.TextHorizontalAlignmentConstants.igTextHzAlignCenter
            partlist.Columns.Item(1).Width = 0.01

            objColumn = partlist.Columns.Item(4)
            proptext = objColumn.PropertyText

            objColumn = partlist.Columns.Item(1)
            objColumn.Header = "Pos."
            objColumn.Width = 0.015
            objColumn.DataHorizontalAlignment = SolidEdgeFrameworkSupport.TextHorizontalAlignmentConstants.igTextHzAlignCenter

            Select Case General.userlangID
                Case 1031
                    'German
                    postext = "%{Positionsnummer|G}"
                Case 1038
                    'Hungarian
                    postext = "%{Tételszám sorrend|G}"
                Case Else
                    'English
                    postext = "%{Item Number|G}"
            End Select

            objColumn.PropertyText = postext
            objColumn.HeaderRowHorizontalAlignment = SolidEdgeFrameworkSupport.TextHorizontalAlignmentConstants.igTextHzAlignCenter

            objColumn = partlist.Columns.Item(2)
            objColumn.Header = "Quantity"
            objColumn.Width = 0.015
            objColumn.DataHorizontalAlignment = SolidEdgeFrameworkSupport.TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.PropertyText = proptext
            objColumn.HeaderRowHorizontalAlignment = SolidEdgeFrameworkSupport.TextHorizontalAlignmentConstants.igTextHzAlignCenter

            objColumn = partlist.Columns.Item(3)
            objColumn.Header = "ERP Number"
            objColumn.Width = 0.025
            objColumn.DataHorizontalAlignment = SolidEdgeFrameworkSupport.TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.PropertyText = "%{CDB_ERP_Artnr./CP|G}"
            objColumn.HeaderRowHorizontalAlignment = SolidEdgeFrameworkSupport.TextHorizontalAlignmentConstants.igTextHzAlignCenter

            objColumn = partlist.Columns.Item(4)
            objColumn.Header = "Designation"
            objColumn.Width = 0.05
            objColumn.DataHorizontalAlignment = SolidEdgeFrameworkSupport.TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.PropertyText = "%{CDB_Benennung_de/CP|G}"

            objColumn = partlist.Columns.Add(5, True)

            objColumn.Header = "PDM Number"
            objColumn.Width = 0.03
            objColumn.DataHorizontalAlignment = SolidEdgeFrameworkSupport.TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.PropertyText = "%{CDB_teilenummer/CP|G}"
            objColumn.HeaderRowHorizontalAlignment = SolidEdgeFrameworkSupport.TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.Show = True

            partlist.Update()
            isoDV.Update()

            isoDV.GetOrigin(newx, newy)
            objShapes = dftdoc.ActiveSheet.AnnotAlignmentShapes
            objShapes.Item(1).Range(xmin, ymin, xmax, ymax)
            isoDV.ScaleFactor *= 0.75
            objShapes.Item(1).Delete()
            isoDV.ScaleFactor /= 0.75
            isoDV.SetOrigin(newx, newy - 0.85 * ymin)
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function GetScaling(finnedheight As Double, finneddepth As Double, sheetname As String) As Double
        Dim scalefactor As Double = 1
        Dim sflist As New List(Of Double) From {0.1, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5}

        If sheetname.Contains("Main") Then
            Select Case finnedheight
                Case 400
                    scalefactor = 0.25
                Case 500
                    scalefactor = 0.2
                Case 600
                    scalefactor = 0.3
                Case 700
                    scalefactor = 0.25
                Case 900
                    scalefactor = 0.2
                Case 1200
                    scalefactor = 1 / 6
                Case 1600
                    scalefactor = 0.15
            End Select
            If sheetname.Contains("Rescale") Then
                Try
                    If finnedheight = 1200 Then
                        scalefactor = 0.15
                    Else
                        scalefactor = sflist(sflist.IndexOf(scalefactor) - 1)
                    End If
                Catch ex As Exception
                    scalefactor = 0.1
                End Try
            End If
        Else
            Select Case finnedheight
                Case 600
                    scalefactor = 0.4
                Case 700
                    scalefactor = 0.3
                Case 900
                    scalefactor = 0.25
                Case 1200
                    scalefactor = 0.2
                Case 1600
                    scalefactor = 0.15
                Case Else
                    scalefactor = 0.5
            End Select

            'reduce scaling by one step if it's L chamber
            If ((finneddepth > 300 And finnedheight < 1600) Or (finneddepth > 250 And finnedheight = 400)) And sheetname.Contains("Rescale") = False Then
                scalefactor = sflist(sflist.IndexOf(scalefactor) - 1)
            ElseIf sheetname.Contains("Rescale") Then
                Try
                    scalefactor = sflist(sflist.IndexOf(scalefactor) - 2)
                Catch ex As Exception
                    scalefactor = 0.1
                End Try
            End If

        End If

        Return scalefactor
    End Function

    Shared Sub DVPositions(DVlist As List(Of SolidEdgeDraft.DrawingView), ConSys As ConSysData, sheetsize As String, coil As CoilData, circuit As CircuitData)
        Dim location() As Double = {0, 0}
        Dim sheetframe() As Double
        Dim toprightDV, lowerDV, centerDV, leftDV, inletDV, outletDV, frontDV As SolidEdgeDraft.DrawingView
        Dim x, y, x0, y0, lowerdistance, movefactor, xmin1, ymin1, xmax1, ymax1, xmin2, ymin2, xmax2, ymax2, xmid, xs, ys As Double
        Dim x0out, y0out, x0in, y0in, xmin3, xmax3, ymin3, ymax3, deltay, deltax As Double

        Try
            If sheetsize = "A3" Then
                sheetframe = {0.415, 0.29}
                lowerdistance = 0.008
                movefactor = 1.5
                If coil.FinnedDepth > 299 Then
                    movefactor = 2
                End If
            Else
                movefactor = 1
                sheetframe = {0.59, 0.415}
                lowerdistance = 0.03
            End If

            If circuit.ConnectionSide = "left" Then
                toprightDV = DVlist.Last
                centerDV = DVlist(0)
            Else
                toprightDV = DVlist(0)
                centerDV = DVlist(2)
            End If

            y = sheetframe(1) - 0.02 - toprightDV.CropTop
            x = sheetframe(0) - 0.03 - toprightDV.CropRight

            toprightDV.SetOrigin(x, y)

            lowerDV = DVlist(1)
            y = 0.07 + lowerdistance + lowerDV.CropBottom

            If circuit.ConnectionSide = "left" Then
                If General.currentunit.ModelRangeSuffix.Substring(1, 1) = "P" Then
                    toprightDV.GetOrigin(x0, y0)
                    toprightDV.SetOrigin(x0 - 0.05, y0)
                    centerDV = DVlist(2)
                    leftDV = DVlist(0)

                    movefactor = toprightDV.ScaleFactor

                    x = x0 - Math.Abs(movefactor * toprightDV.CropLeft * 1.5)
                    centerDV.SetOrigin(x, y0)

                    toprightDV.Range(xmin1, ymin1, xmax1, ymax1)

                    x = coil.FinnedDepth * movefactor / 1000 + 0.065
                    leftDV.SetOrigin(x, y0)
                    leftDV.Range(xmin2, ymin2, xmax2, ymax2)

                    'where the mid of range of centerDV should be
                    xmid = (xmin1 + xmax2) / 2
                    centerDV.Range(xmin1, ymin1, xmax1, ymax1)

                    'where the mid of ranger actually is
                    xmin2 = (xmin1 + xmax1) / 2

                    centerDV.GetOrigin(xs, ys)

                    centerDV.SetOrigin(xs - (xmin2 - xmid), ys)

                    y += 0.045
                Else
                    x = sheetframe(0) / 2
                    y += 0.01
                End If
            Else
                toprightDV.GetOrigin(x0, y0)
                movefactor = toprightDV.ScaleFactor

                x = Math.Round(x0 - toprightDV.CropLeft - centerDV.CropRight - 0.05, 6)

                centerDV.SetOrigin(x, y0)

                'for FP 
                If DVlist.Count = 4 Then
                    x -= 0.65 * movefactor
                    y -= 0.0075
                    DVlist(3).SetOrigin(x, y0)
                    DVlist(3).Range(xmin1, ymin1, xmax1, ymax1)
                    If xmin1 < 0.05 Then
                        DVlist(3).SetOrigin(x + Math.Abs(0.05 - xmin1), y0)
                    End If
                End If

                x = x0
                If ConSys.HeaderAlignment = "horizontal" Then
                    frontDV = DVlist(0)
                    outletDV = DVlist(2)
                    inletDV = DVlist(3)

                    frontDV.Range(xmin1, ymin1, xmax1, ymax1)

                    outletDV.GetOrigin(x0out, y0out)
                    outletDV.Range(xmin2, ymin2, xmax2, ymax2)

                    inletDV.GetOrigin(x0in, y0in)
                    inletDV.Range(xmin3, ymin3, xmax3, ymax3)

                    deltay = ymin1 - ymin3
                    y0in += deltay
                    deltax = xmin1 - xmax3 - 0.1
                    x0in += deltax

                    inletDV.SetOrigin(x0in, y0in)

                    deltay = ymax1 - ymax2
                    y0out += deltay

                    outletDV.SetOrigin(x0in, y0out)
                End If
            End If

            If ConSys.HeaderAlignment = "vertical" Then
                lowerDV.SetOrigin(x, y)
            Else
                lowerDV.Delete()
            End If

            'center DV should be aligned correctly already

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub DVPositionsHor(DVlist As List(Of SolidEdgeDraft.DrawingView), consys As ConSysData, sheetsize As String, coil As CoilData, circuit As CircuitData)
        Dim location() As Double = {0, 0}
        Dim sheetframe() As Double
        Dim xsizes, ysizes As New List(Of Double)
        Dim x, y, xmin1, ymin1, xmax1, ymax1 As Double


        Try
            'frontDV top right for horizontal and mirrored vertical
            If sheetsize = "A3" Then
                sheetframe = {0.415, 0.29}
            Else
                sheetframe = {0.59, 0.415}
            End If

            'get the size of each DV
            For Each DV In DVlist
                DV.Range(xmin1, ymin1, xmax1, ymax1)
                xsizes.Add(Math.Abs(Math.Round(xmax1 - xmin1, 6)))
                ysizes.Add(Math.Abs(Math.Round(ymax1 - ymin1, 6)))
            Next

            y = sheetframe(1) - ysizes(0) / 2 - sheetframe(1) / 20
            If circuit.ConnectionSide = "right" And consys.HeaderAlignment = "vertical" Then
                x = sheetframe(0) - Math.Round(xsizes(0) / 2 + sheetframe(0) / 20, 6)

                'frontDV top right
                DVlist(0).SetOrigin(x, y)
                'inlet and outlet folded from front, seperated
                DVlist(1).SetOrigin(x - 2 * xsizes(1), y)
                DVlist(2).SetOrigin(x - 4 * xsizes(1), y)

            Else
                x = Math.Round(xsizes(0) / 2 + sheetframe(0) / 15, 6)

                'frontDV top left
                DVlist(0).SetOrigin(x, y)
                If consys.HeaderAlignment = "horizontal" Then
                    'inlet and outlet on top of each other
                    DVlist(1).SetOrigin(Math.Round(x + 1.5 * xsizes(1), 6), y)
                    DVlist(2).SetOrigin(Math.Round(x + 1.5 * xsizes(1), 6), y)
                Else
                    DVlist(1).SetOrigin(Math.Round(x + 2 * xsizes(1), 6), y)
                    DVlist(2).SetOrigin(Math.Round(x + 4 * xsizes(1), 6), y)
                End If
            End If

            'top DV
            DVlist(3).SetOrigin(x, y - ysizes(1) / 2 - sheetframe(1) / 20 - ysizes(3) / 2)

        Catch ex As Exception

        End Try

    End Sub

    Shared Sub CreateADBows(dftdoc As SolidEdgeDraft.DraftDocument, position As String, circuit As CircuitData)

        Try
            For Each objSheet As SolidEdgeDraft.Sheet In dftdoc.Sheets
                If objSheet.Name.Contains("Coil") Then
                    For i As Integer = 0 To objSheet.DrawingViews.Count - 1
                        If objSheet.DrawingViews.Item(i + 1).CaptionDefinitionTextPrimary.Contains("Front") Then
                            SEDrawing.AddADBlock(dftdoc, objSheet.DrawingViews.Item(i + 1), position)
                        End If
                    Next
                    For i As Integer = 0 To objSheet.DrawingViews.Count - 1
                        If objSheet.DrawingViews.Item(i + 1).CaptionDefinitionTextPrimary.Contains("Back") Then
                            SEDrawing.AddADBlock(dftdoc, objSheet.DrawingViews.Item(i + 1), circuit.ConnectionSide)
                        End If
                    Next
                    Exit For
                End If
                If objSheet.Name = "Front" Then
                    objSheet.Activate()
                    General.seapp.DoIdle()
                    For Each objDV In objSheet.DrawingViews
                        SEDrawing.AddADBlock(dftdoc, objDV, position)
                    Next
                End If
                If objSheet.Name = "Back" Then
                    objSheet.Activate()
                    General.seapp.DoIdle()
                    For Each objDV In objSheet.DrawingViews
                        SEDrawing.AddADBlock(dftdoc, objDV, circuit.ConnectionSide)
                    Next
                End If
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub SwitchLayers(dftdoc As SolidEdgeDraft.DraftDocument, plant As String)

        Try
            Dim layers As SolidEdgeFramework.Layers = dftdoc.ActiveSheet.Layers
            Dim bejilayer As SolidEdgeFramework.Layer = layers.Item("Beji")
            Dim eulayer As SolidEdgeFramework.Layer = layers.Item("Europe")

            If plant = "Beji" Then
                eulayer.HideEverywhere()
                bejilayer.ShowEverywhere()
            Else
                eulayer.ShowEverywhere()
                bejilayer.HideEverywhere()
            End If

        Catch ex As Exception

        End Try
    End Sub

    Shared Sub CreateAirDirection(dftdoc As SolidEdgeDraft.DraftDocument, conside As String, finneddepth As Double)
        Dim objDVs As SolidEdgeDraft.DrawingViews
        Dim objDV As SolidEdgeDraft.DrawingView
        Dim objBlocks As SolidEdgeDraft.Blocks
        Dim objBlock As SolidEdgeDraft.Block
        Dim objBOccs As SolidEdgeDraft.BlockOccurrences
        Dim objBOcc As SolidEdgeDraft.BlockOccurrence
        Dim delta, x0, y0 As Double
        Dim n, mmp As Integer

        Try
            'Frontview always first drawing view
            objDVs = dftdoc.ActiveSheet.DrawingViews
            objDV = objDVs(0)

            objBlocks = dftdoc.Blocks

            If conside = "left" Then
                objBlock = objBlocks(1)
                mmp = 1
            Else
                objBlock = objBlocks(0)
                mmp = -1
            End If

            objDV.GetOrigin(x0, y0)
            n = finneddepth / 50
            delta = n * 0.0065 * mmp

            Debug.Print(delta.ToString)

            'origin as function of finned depth
            objBOccs = dftdoc.ActiveSheet.BlockOccurrences
            objBOcc = objBOccs.Add(objBlock.Name, x0 + delta, y0 + 0.01)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub OneBranchFrontViewDims(dftdoc As SolidEdgeDraft.DraftDocument, consys As ConSysData, coil As CoilData, circuit As CircuitData)
        Dim objDVs As SolidEdgeDraft.DrawingViews
        Dim frontDV As SolidEdgeDraft.DrawingView
        Dim finID, SVID As String
        Dim finlines, inletlines, outletlines, svlines, flangelines As New List(Of SolidEdgeDraft.DVLine2d)
        Dim findimlines(), inletdimlines(), outletdimlines() As SolidEdgeDraft.DVLine2d
        Dim x0, y0, scalefactor As Double
        Dim searchneeded As Boolean = False


        Try
            objDVs = dftdoc.ActiveSheet.DrawingViews
            frontDV = objDVs(0)
            frontDV.GetOrigin(x0, y0)
            scalefactor = frontDV.ScaleFactor

            finID = "Fin1"

            If circuit.Pressure > 50 Then
                SVID = "0000791913"
            Else
                SVID = "0000107560"
            End If

            If frontDVLinelist.Count = 0 Then
                searchneeded = True
                finlines = SEDrawing.GetLinesFromOcc(frontDV, finID, circuit.ConnectionSide)
                inletlines = SEDrawing.GetLinesFromOcc(frontDV, consys.InletHeaders.First.StutzenDatalist.First.ID, circuit.ConnectionSide)
                outletlines = SEDrawing.GetLinesFromOcc(frontDV, consys.OutletHeaders.First.StutzenDatalist.First.ID, circuit.ConnectionSide)
                svlines = SEDrawing.GetLinesFromOcc(frontDV, SVID, circuit.ConnectionSide)
            Else
                Dim tempflist = From partiallist In frontDVLinelist Where partiallist.RefFileName.Contains(finID)

                For Each tempelement In tempflist
                    finlines.Add(tempelement.DVLine)
                Next

                Dim temphlist = From partiallist In frontDVLinelist Where partiallist.RefFileName.Contains(consys.InletHeaders.First.StutzenDatalist.First.ID)

                For Each tempelement In temphlist
                    inletlines.Add(tempelement.DVLine)
                Next
                Dim tempnlist = From partiallist In frontDVLinelist Where partiallist.RefFileName.Contains(consys.OutletHeaders.First.StutzenDatalist.First.ID)

                For Each tempelement In tempnlist
                    outletlines.Add(tempelement.DVLine)
                Next

                Dim tempsvlist = From partiallist In frontDVLinelist Where partiallist.RefFileName.Contains(SVID)

                For Each tempelement In tempsvlist
                    svlines.Add(tempelement.DVLine)
                Next
            End If
            Debug.Print(inletlines.Count.ToString)

            findimlines = GetDimLines(finlines, circuit.ConnectionSide, "fin")

            frontDV.ScaleFactor = 1
            frontDV.SetOrigin(0, 0)

            SetLengthDim(dftdoc.ActiveSheet, frontDV, findimlines(0), "fintop", circuit.ConnectionSide)
            SetLengthDim(dftdoc.ActiveSheet, frontDV, findimlines(1), "finside", circuit.ConnectionSide)

            outletdimlines = GetDimLines(outletlines, circuit.ConnectionSide, "nipple")

            If General.currentunit.ModelRangeSuffix.Contains("X") And circuit.CircuitType <> "Defrost" Then
                'DX inlet horizontal line, should only be one
                inletdimlines = GetDimLines(inletlines, circuit.ConnectionSide, "stutzen")
                DimDX(dftdoc.ActiveSheet, frontDV, findimlines(0), inletdimlines(0), circuit.ConnectionSide, "inlet")
                DimDX(dftdoc.ActiveSheet, frontDV, findimlines(0), outletdimlines(1), circuit.ConnectionSide, "outlet")

            Else
                'XP inlet vertical line
                inletdimlines = GetDimLines(inletlines, circuit.ConnectionSide, "nipple")

                DimFP(dftdoc.ActiveSheet, frontDV, findimlines(0), inletdimlines(1), circuit.ConnectionSide, "inlet")
                DimFP(dftdoc.ActiveSheet, frontDV, findimlines(0), outletdimlines(1), circuit.ConnectionSide, "outlet")
            End If

            frontDV.ScaleFactor = scalefactor
            frontDV.SetOrigin(x0, y0)

            HandleVisibility(frontDV, "front", consys)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function GetEllipsesFromOcc(objDV As SolidEdgeDraft.DrawingView, partIDList As List(Of String)) As List(Of SolidEdgeDraft.DVEllipse2d)
        Dim objEllipses As New List(Of SolidEdgeDraft.DVEllipse2d)

        Try
            For Each objEl As SolidEdgeDraft.DVEllipse2d In objDV.DVEllipses2d
                For Each pID In partIDList
                    If objEl.ModelMember.FileName.Contains(pID) Then
                        objEllipses.Add(objEl)
                        Exit For
                    End If
                Next
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return objEllipses
    End Function

    Shared Function GetSplinesFromOcc(objDV As SolidEdgeDraft.DrawingView, partIDList As List(Of String)) As List(Of SolidEdgeDraft.DVBSplineCurve2d)
        Dim objSplinelist As New List(Of SolidEdgeDraft.DVBSplineCurve2d)

        Try
            For Each objSpline As SolidEdgeDraft.DVBSplineCurve2d In objDV.DVBSplineCurves2d
                For Each pID In partIDList
                    If objSpline.ModelMember.FileName.Contains(pID) Then
                        objSplinelist.Add(objSpline)
                    End If
                Next
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return objSplinelist
    End Function

    Shared Function GetDimLines(DVLines As List(Of SolidEdgeDraft.DVLine2d), conside As String, objecttype As String) As SolidEdgeDraft.DVLine2d()
        Dim topline, sideline, lines() As SolidEdgeDraft.DVLine2d
        Dim xs, xe, ys, ye, xref, yref As Double

        Try
            If objecttype = "fin" Then
                yref = -1
            ElseIf objecttype = "nipple" Then
                yref = -4
            Else
                yref = 1
            End If

            'get frame lines
            For i As Integer = 0 To DVLines.Count - 1
                Dim DVline As SolidEdgeDraft.DVLine2d = DVLines(i)
                DVline.GetStartPoint(xs, ys)
                DVline.GetEndPoint(xe, ye)

                If Math.Abs(ys - ye) < 0.0001 Then
                    'horizontal, topline for fin, bottom line for header
                    If objecttype = "fin" Or objecttype = "nipple" Then
                        yref = Math.Round(Math.Max(ys, yref), 6)
                        If Math.Round(ye, 6) >= yref Then
                            'top line
                            topline = DVline
                        End If
                    Else
                        yref = Math.Round(Math.Min(ys, yref), 6)
                        If Math.Round(ye, 6) <= yref Then
                            'Debug.Print(Math.Round(DVline.Length, 6).ToString + " at y= " + Math.Round(ys, 6).ToString)
                            'bottom line
                            topline = DVline
                        End If
                    End If
                ElseIf Math.Abs(xs - xe) < 0.0001 Then
                    'vertical
                    If objecttype = "nipple" Then
                        If conside = "left" Then
                            xref = Math.Round(Math.Max(xe, xref), 6)
                            If Math.Round(xe, 6) >= xref Then
                                'get right side line
                                sideline = DVline
                            End If
                        Else
                            'left side needed
                            xref = Math.Round(Math.Min(xe, xref), 6)
                            If Math.Round(xe, 6) <= xref Then
                                'get left side line
                                sideline = DVline
                            End If
                        End If
                    Else
                        If conside = "left" Then
                            'left side needed
                            xref = Math.Round(Math.Min(xe, xref), 6)
                            If Math.Round(xe, 6) <= xref Then
                                'get left side line
                                sideline = DVline
                            End If
                        Else
                            xref = Math.Round(Math.Max(xe, xref), 6)
                            If Math.Round(xe, 6) >= xref Then
                                'get right side line
                                sideline = DVline
                            End If
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
        lines = {topline, sideline}

        Return lines
    End Function

    Shared Function GetCTLines(DVLines As List(Of SolidEdgeDraft.DVLine2d), headertype As String, conside As String, header As HeaderData)
        Dim leftline, rightline, lines() As SolidEdgeDraft.DVLine2d
        Dim xlist, ylist As New List(Of Double)
        Dim uniquex As List(Of Double)
        Dim xs, xe, ys, ye, xref, yref, z, targetx As Double
        Dim htype, ktype, rowcount As Integer
        Dim infolist As New List(Of KeyInfo)
        Dim lineinfo As New KeyInfo

        Try
            For i As Integer = 0 To DVLines.Count - 1
                DVLines(i).GetStartPoint(xs, ys)
                DVLines(i).GetEndPoint(xe, ye)
                If Math.Abs(ys - ye) < 0.0001 Then
                    'horizontal line
                    DVLines(i).GetKeyPoint(2, xref, yref, z, ktype, htype)
                    infolist.Add(New KeyInfo With {.Kkey = DVLines(i).Key, .X = Math.Round(xref, 6), .Y = Math.Round(yref, 6), .Kindex = i})
                    xlist.Add(Math.Round(xref, 6))
                    ylist.Add(Math.Round(yref, 6))
                End If
            Next

            Dim templist = From plist In infolist Where plist.Y = ylist.Min

            Debug.Print(templist.Count.ToString)
            If headertype = "outlet" Then
                For j As Integer = 0 To templist.Count - 1
                    lineinfo = templist(j)
                    If lineinfo.X = xlist.Min Then
                        Debug.Print(lineinfo.Kindex.ToString)
                        leftline = DVLines(lineinfo.Kindex)
                    ElseIf lineinfo.X = xlist.Max Then
                        Debug.Print(lineinfo.Kindex.ToString)
                        rightline = DVLines(lineinfo.Kindex)
                    End If
                Next
            Else
                uniquex = General.GetUniqueValues(xlist)
                uniquex.Sort()

                If conside = "right" Then
                    uniquex.Reverse()
                End If

                rowcount = General.GetUniqueValues(header.Xlist).Count
                targetx = uniquex(rowcount - 1)
                For i As Integer = 0 To infolist.Count - 1
                    If infolist(i).X = targetx And infolist(i).Y = ylist.Min Then
                        leftline = DVLines.Item(infolist(i).Kindex)
                        leftline.GetKeyPoint(2, xref, yref, z, ktype, htype)
                        Debug.Print(xref.ToString + " / " + yref.ToString)
                        Exit For
                    End If
                Next
            End If

            lines = {leftline, rightline}

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return lines
    End Function

    Shared Function GetVentLine(DVLines As List(Of SolidEdgeDraft.DVLine2d), conside As String, finline As SolidEdgeDraft.DVLine2d) As SolidEdgeDraft.DVLine2d
        Dim sideline As SolidEdgeDraft.DVLine2d = Nothing
        Dim vertlines As New List(Of SolidEdgeDraft.DVLine2d)
        Dim xlist As New List(Of Double)
        Dim xe, xs, ye, ys, xf, yf, zf As Double
        Dim ktype, htype As Integer

        Try
            finline.GetKeyPoint(2, xf, yf, zf, ktype, htype)

            For Each DVLine In DVLines
                DVLine.GetStartPoint(xs, ys)
                DVLine.GetEndPoint(xe, ye)
                If Math.Round(xs, 6) = Math.Round(xe, 6) Then
                    If (conside = "left" And xs < xf) Or (conside = "right" And xs > xf) Then
                        vertlines.Add(DVLine)
                        xlist.Add(xs)
                    End If
                End If
            Next

            If conside = "left" Then
                sideline = vertlines(xlist.IndexOf(xlist.Max))
            Else
                sideline = vertlines(xlist.IndexOf(xlist.Min))
            End If

        Catch ex As Exception

        End Try

        Return sideline
    End Function

    Shared Function SetLengthDim(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, DVLine As SolidEdgeDraft.DVLine2d, linetype As String, conside As String) As SolidEdgeFrameworkSupport.Dimension
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objSheet.Dimensions
        Dim lengthdim As SolidEdgeFrameworkSupport.Dimension
        Dim xe, ye, xs, ys, trackdistance As Double

        Try
            DVLine.GetStartPoint(xs, ys)
            DVLine.GetEndPoint(xe, ye)

            lengthdim = objdims.AddDistanceBetweenObjects(DVLine, xs, ys, 0, True, DVLine, xe, ye, 0, True)
            lengthdim.ReattachToDrawingView(DV)

            With lengthdim
                If linetype.Contains("fin") Then
                    Dim xmin, ymin, xmax, ymax As Double
                    If linetype = "finside" Then
                        .TrackDistance = -0.02
                        .Range(xmin, ymin, xmax, ymax)
                        If (Math.Round(xmax, 6) <= Math.Round(xe, 6) And conside = "right") Or (Math.Round(xmin, 6) <= Math.Round(xe, 6) And conside = "left") Then
                            .TrackDistance *= -1
                        End If
                    Else
                        .TrackDistance = -0.01
                        .Range(xmin, ymin, xmax, ymax)
                        If Math.Round(ymax, 6) <= Math.Round(ye, 6) Then
                            lengthdim.TrackDistance *= -1
                        End If
                    End If
                    trackdistance = .TrackDistance
                    .PrefixString = "finned:"
                Else
                    If linetype = "headerbottom" Then
                        .TrackDistance = 0.01
                        SEDrawing.ControlDimDistance(lengthdim, {"y", "smaller"})
                        .PrefixString = "%DI"
                        If conside = "left" Then
                            .BreakPosition = 3
                        Else
                            .BreakPosition = 1
                        End If
                        .BreakDistance = 0.003
                        .TerminatorPosition = True
                    ElseIf linetype = "nippleside" Then
                        .TrackDistance = 0.015
                        .PrefixString = "%DI"
                        If conside = "left" Then
                            .BreakPosition = 1
                            SEDrawing.ControlDimDistance(lengthdim, {"x", "bigger"})
                        Else
                            .BreakPosition = 3
                            SEDrawing.ControlDimDistance(lengthdim, {"x", "smaller"})
                        End If
                        .BreakDistance = 0.003
                        .TerminatorPosition = True
                    Else
                        .TrackDistance = -0.0136 - 0.022
                    End If
                End If
            End With

            If DVLine.ModelMember.FileName.Contains("InletHeader") Then
                lengthdim.BreakDistance = 0.008
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return lengthdim
    End Function

    Shared Sub DimDX(objsheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, finline As SolidEdgeDraft.DVLine2d, line As SolidEdgeDraft.DVLine2d, conside As String, headertype As String)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objsheet.Dimensions
        Dim distancedim As SolidEdgeFrameworkSupport.Dimension
        Dim xs, ys, xe, xc, yc, zc, xref, mp As Double
        Dim ktype, htype As Integer

        Try
            finline.GetStartPoint(xs, ys)
            finline.GetEndPoint(xe, ys)

            line.GetKeyPoint(2, xc, yc, zc, ktype, htype)

            mp = 1
            If conside = "right" Then
                'right point of fin line, center of stutzenline
                xref = Math.Max(xs, xe)
            Else
                'left point of fin line, center of stutzenline 
                xref = Math.Min(xs, xe)
                If headertype = "outlet" Then
                    mp = -1
                End If
            End If

            'horizontal distance
            distancedim = objdims.AddDistanceBetweenObjects(finline, xref, ys, 0, True, line, xc, yc, 0, True)

            distancedim.ReattachToDrawingView(DV)

            distancedim.MeasurementAxisEx = 3
            distancedim.TrackDistance = 0.01 * mp
            distancedim.Style.PrimaryDecimalRoundOff = SolidEdgeFrameworkSupport.DimDecimalRoundOffTypeConstants.igDimStyleDecimal_1

            'vertical distance
            distancedim = objdims.AddDistanceBetweenObjects(line, xc, yc, 0, True, finline, xref, ys, 0, True)

            distancedim.ReattachToDrawingView(DV)

            distancedim.MeasurementAxisDirection = True
            distancedim.MeasurementAxisEx = 1
            distancedim.TrackDistance = 0.02 * mp
            distancedim.Style.PrimaryDecimalRoundOff = SolidEdgeFrameworkSupport.DimDecimalRoundOffTypeConstants.igDimStyleDecimal_1

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try


    End Sub

    Shared Sub DimFP(objsheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, finline As SolidEdgeDraft.DVLine2d, line As SolidEdgeDraft.DVLine2d, conside As String, headertype As String)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objsheet.Dimensions
        Dim distancedim As SolidEdgeFrameworkSupport.Dimension
        Dim xs, ys, xe, xc, yc, zc, td, mp, xref, fangle As Double
        Dim ktype, htype As Integer

        Try
            'right point of fin line, center of stutzenline
            finline.GetStartPoint(xs, ys)
            finline.GetEndPoint(xe, ys)
            fangle = Math.Round(finline.Angle, 6)

            If conside = "right" Then
                xref = Math.Min(xs, xe)
                mp = 1
            Else
                xref = Math.Max(xs, xe)
                mp = -1
            End If

            line.GetKeyPoint(2, xc, yc, zc, ktype, htype)

            'horizontal distance
            If headertype = "inlet" Then
                distancedim = objdims.AddDistanceBetweenObjects(finline, xref, ys, 0, True, line, xc, yc, 0, True)
                td = 0.01 * mp * fangle / Math.Abs(fangle)
            Else
                distancedim = objdims.AddDistanceBetweenObjects(line, xc, yc, 0, True, finline, xref, ys, 0, True)
                td = 0.02
            End If

            distancedim.ReattachToDrawingView(DV)

            distancedim.MeasurementAxisEx = 3
            distancedim.TrackDistance = td
            distancedim.Style.PrimaryDecimalRoundOff = SolidEdgeFrameworkSupport.DimDecimalRoundOffTypeConstants.igDimStyleDecimal_1

            'vertical distance
            distancedim = objdims.AddDistanceBetweenObjects(line, xc, yc, 0, True, finline, xref, ys, 0, True)

            distancedim.ReattachToDrawingView(DV)

            distancedim.MeasurementAxisDirection = True
            distancedim.MeasurementAxisEx = 1
            If headertype = "inlet" Then
                distancedim.TrackDistance = -0.02 * fangle / Math.Abs(fangle)
            Else
                distancedim.TrackDistance = 0.03 * mp
            End If
            distancedim.Style.PrimaryDecimalRoundOff = SolidEdgeFrameworkSupport.DimDecimalRoundOffTypeConstants.igDimStyleDecimal_1

        Catch ex As Exception

        End Try

    End Sub

    Shared Sub HandleVisibility(objDV As SolidEdgeDraft.DrawingView, dvname As String, consys As ConSysData)
        Dim objDVLines As SolidEdgeDraft.DVLines2d = objDV.DVLines2d
        Dim objDVCircles As SolidEdgeDraft.DVCircles2d = objDV.DVCircles2d
        Dim fileID As String

        Try

            For Each objDVLine As SolidEdgeDraft.DVLine2d In objDVLines
                fileID = objDVLine.ModelMember.FileName
                If fileID.Contains("00") Then
                    fileID = objDVLine.ModelMember.FileName.Substring(0, 10)
                End If
                If dvname = "front" Then
                    'save all lines in a list
                    Dim newLineElement As New DVLineElement With {.DVLine = objDVLine, .RefFileName = objDVLine.ModelMember.FileName}
                    'frontDVLinelist.Add(newLineElement)
                End If

                If outletIDs.Contains(fileID) And dvname = "front" Then
                    objDVLine.ModelMember.ShowEdgesHiddenByOtherParts = True
                Else
                    objDVLine.ModelMember.ShowTangentEdges = True
                    objDVLine.ModelMember.ShowEdgesHiddenByOtherParts = False
                    objDVLine.ModelMember.ShowHiddenEdges = False
                End If
            Next

            If dvname = "side" Then
                For Each objDVArc As SolidEdgeDraft.DVArc2d In objDV.DVArcs2d
                    fileID = objDVArc.ModelMember.FileName
                    If fileID.Contains("letNipple") Then
                        objDVArc.ModelMember.ShowEdgesHiddenByOtherParts = True
                        objDVArc.ModelMember.ShowHiddenEdges = True
                        sideDVArclist.Add(objDVArc)
                    End If
                Next
            End If

            For Each objDVCircle As SolidEdgeDraft.DVCircle2d In objDVCircles
                fileID = objDVCircle.ModelMember.FileName
                If fileID.Contains("00") Then
                    fileID = objDVCircle.ModelMember.FileName.Substring(0, 10)
                End If
                Dim inventcheck As Boolean = False
                If consys.InletHeaders.First.VentIDs IsNot Nothing Then
                    If consys.InletHeaders.First.VentIDs.Contains(fileID) Then
                        inventcheck = True
                    End If
                End If
                Dim outventcheck As Boolean = False
                If consys.OutletHeaders.First.VentIDs IsNot Nothing Then
                    If consys.OutletHeaders.First.VentIDs.Contains(fileID) Then
                        outventcheck = True
                    End If
                End If
                If dvname <> "top" AndAlso (outletIDs.Contains(fileID) OrElse inletIDs.Contains(fileID) OrElse inventcheck OrElse outventcheck OrElse fileID.Contains("letNipple")) Then
                    objDVCircle.ModelMember.ShowEdgesHiddenByOtherParts = True
                Else
                    objDVCircle.ModelMember.ShowEdgesHiddenByOtherParts = False
                    objDVCircle.ModelMember.ShowHiddenEdges = False
                End If
            Next

            If dvname = "top" Then
                For Each objDVArc As SolidEdgeDraft.DVArc2d In objDV.DVArcs2d
                    objDVArc.ModelMember.ShowEdgesHiddenByOtherParts = False
                    objDVArc.ModelMember.ShowHiddenEdges = False
                Next
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function GetHeaderTopLine(headerlines As List(Of SolidEdgeDraft.DVLine2d), bottomline As SolidEdgeDraft.DVLine2d) As SolidEdgeDraft.DVLine2d
        Dim topline As SolidEdgeDraft.DVLine2d
        Dim xsb, ysb, xst, yst, xeb, yeb, xet, yet As Double

        Try
            bottomline.GetStartPoint(xsb, ysb)
            bottomline.GetEndPoint(xeb, yeb)

            For i As Integer = 0 To headerlines.Count - 1
                headerlines(i).GetStartPoint(xst, yst)
                headerlines(i).GetEndPoint(xet, yet)
                If Math.Abs(yet - yst) < 0.001 And yst > ysb Then
                    If (Math.Abs(xst - xsb) < 0.001 Or Math.Abs(xst - xeb) < 0.001) And (Math.Abs(xet - xsb) < 0.001 Or Math.Abs(xet - xeb) < 0.001) Then
                        topline = headerlines(i)
                        Exit For
                    End If
                End If
            Next
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        If topline Is Nothing Then
            topline = bottomline
        End If

        Return topline
    End Function

    Shared Function SetHeaderLength(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, bottomline As SolidEdgeDraft.DVLine2d, topline As SolidEdgeDraft.DVLine2d, conside As String, mrsuffix As String) As Double
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objSheet.Dimensions
        Dim distancedim As SolidEdgeFrameworkSupport.Dimension
        Dim xbref, xtref, xsb, ysb, xeb, yeb, xst, yst, xet, yet, xmin, ymin, xmax, ymax As Double

        Try
            bottomline.GetStartPoint(xsb, ysb)
            bottomline.GetEndPoint(xeb, yeb)
            Debug.Print(bottomline.Key)
            topline.GetStartPoint(xst, yst)
            topline.GetEndPoint(xet, yet)
            Debug.Print(topline.Key)
            If conside = "left" Then
                xbref = Math.Min(xsb, xeb)
                xtref = Math.Min(xst, xet)
            Else
                xbref = Math.Max(xsb, xeb)
                xtref = Math.Max(xst, xet)
            End If
            distancedim = objdims.AddDistanceBetweenObjects(bottomline, xbref, ysb, 0, True, topline, xtref, yst, 0, True)
            distancedim.ReattachToDrawingView(DV)

            If mrsuffix = "FP" Or mrsuffix = "WP" Then
                distancedim.TrackDistance = 0.04
            Else
                distancedim.TrackDistance = -0.0136 - 0.022
            End If

            If conside = "left" Then
                distancedim.TrackDistance = -distancedim.TrackDistance
                distancedim.Range(xmin, ymin, xmax, ymax)
                If Math.Round(xmax, 6) <= Math.Round(xbref, 6) Then
                    distancedim.TrackDistance = -distancedim.TrackDistance
                End If
            Else
                If mrsuffix = "FP" Or mrsuffix = "WP" Then
                    distancedim.Range(xmin, ymin, xmax, ymax)
                    If xmax > Math.Max(xsb, xeb) + 0.02 Then
                        distancedim.TrackDistance = -distancedim.TrackDistance
                    End If
                End If
            End If


        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return Math.Round(distancedim.Value * 1000)
    End Function

    Shared Sub MoveBlock(dftdoc As SolidEdgeDraft.DraftDocument, deltax As Double, deltay As Double)
        Dim objBOccs As SolidEdgeDraft.BlockOccurrences
        Dim objBOcc As SolidEdgeDraft.BlockOccurrence
        Dim x0, y0 As Double

        Try
            'Frontview always first drawing view

            objBOccs = dftdoc.ActiveSheet.BlockOccurrences
            objBOcc = objBOccs.Item(1)
            objBOcc.GetOrigin(x0, y0)
            objBOcc.SetOrigin(x0 - deltax, y0 - deltay)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub ControlDimensions(DV As SolidEdgeDraft.DrawingView, objdims As SolidEdgeFrameworkSupport.Dimensions, scalefactor As Double, consys As ConSysData, coil As CoilData, circuit As CircuitData)
        Dim x1, y1, x2, y2, x3, x4, y3, y4, z, ctd, newtd, gap, dx As Double
        Dim n As Integer = Math.Round(coil.FinnedDepth / 50)

        Try
            'position of depth dimension
            objdims.Item(1).Range(x1, y1, x2, y2)

            'position of height dimension
            objdims.Item(2).Range(x3, y3, x4, y4)

            If Math.Round(y2, 6) <= Math.Round(y4, 6) Then
                objdims.Item(1).TrackDistance = -objdims.Item(1).TrackDistance
            End If

            If (Math.Round(x3, 6) < Math.Round(x2, 6) And circuit.ConnectionSide = "right") Or (Math.Round(x4, 6) > Math.Round(x1, 6) And circuit.ConnectionSide = "left") Then
                z = x2 - x3
                ctd = objdims.Item(2).TrackDistance
                objdims.Item(2).TrackDistance = -ctd
            End If

            If inletIDs.Count = 0 Then
                'position of header length dim
                objdims.Item(4).Range(x3, y3, x4, y4)
                If (Math.Round(x3, 6) < Math.Round(x2, 6) And circuit.ConnectionSide = "right") Or (Math.Round(x3, 6) > Math.Round(x1, 6) And circuit.ConnectionSide = "left") Then
                    z = x2 - x3
                    Debug.Print("Delta = " + Math.Round(z, 5).ToString)
                    ctd = objdims.Item(4).TrackDistance
                    newtd = -n * 0.0125
                    If circuit.FinType <> "N" And circuit.FinType <> "M" Then
                        newtd -= 0.0025
                    End If
                    newtd *= scalefactor / 0.25
                    If circuit.ConnectionSide = "left" Then
                        newtd = -newtd
                    End If
                    If Math.Abs(newtd) > Math.Abs(ctd) Then
                        'check with different scale factor and finned depth
                        objdims.Item(4).TrackDistance = newtd
                    End If
                End If

                DV.ScaleFactor = scalefactor

                If coil.FinnedHeight < 500 Then
                    gap = 0.0105
                Else
                    gap = 0.009
                End If

                'position of height dimension
                objdims.Item(2).Range(x1, y1, x2, y2)

                'position of header length dim
                objdims.Item(4).Range(x3, y3, x4, y4)
                ctd = objdims.Item(4).TrackDistance

                If circuit.ConnectionSide = "left" Then
                    dx = x3 - x1

                    If Math.Abs(dx - gap) > gap / 3 Then
                        newtd = ctd - (dx - gap)
                        objdims.Item(4).TrackDistance = newtd
                        objdims.Item(4).Range(x3, y3, x4, y4)

                        If Math.Abs(gap) < Math.Round(Math.Abs(x3 + x1), 5) Then
                            'change direction of movement
                            newtd = ctd + dx - gap
                            objdims.Item(4).TrackDistance = newtd
                        End If
                    End If
                Else
                    dx = x2 - x4

                    If Math.Abs(dx - gap) > gap / 3 Then
                        newtd = ctd - (dx - gap)
                        objdims.Item(4).TrackDistance = newtd
                        objdims.Item(4).Range(x3, y3, x4, y4)

                        If Math.Abs(gap) > Math.Round(Math.Abs(x2 - x4), 5) Then
                            'change direction of movement
                            newtd = ctd + dx - gap
                            objdims.Item(4).TrackDistance = newtd
                        End If
                    End If
                End If
            End If

            DV.ScaleFactor = 1

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub SetSVDistance(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, nippleline As SolidEdgeDraft.DVLine2d, svlines As List(Of SolidEdgeDraft.DVLine2d), conside As String)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objSheet.Dimensions
        Dim distancedim As SolidEdgeFrameworkSupport.Dimension
        Dim svline As SolidEdgeDraft.DVLine2d
        Dim dimstyle As SolidEdgeFrameworkSupport.DimStyle
        Dim xsn, ysn, xen, yen, xss, yss, xes, yes, xm, ym, zm, xrmin, xrmax, yrmin, yrmax As Double
        Dim ylist As New List(Of Double)
        Dim ktype, htype As Integer

        Try
            nippleline.GetStartPoint(xsn, ysn)
            nippleline.GetEndPoint(xen, yen)

            For i As Integer = 0 To svlines.Count - 1
                svline = svlines(i)
                svline.GetStartPoint(xss, yss)
                svline.GetEndPoint(xes, yes)
                svline.ModelMember.ShowTangentEdges = True
                If Math.Abs(yes - yss) < 0.001 And svline.Length > 0.01 Then
                    ylist.Add(yes)
                Else
                    ylist.Add(1)
                End If
            Next

            svline = svlines(ylist.IndexOf(ylist.Min))
            svline.GetKeyPoint(2, xm, ym, zm, ktype, htype)

            distancedim = objdims.AddDistanceBetweenObjects(svline, xm, ym, zm, True, nippleline, xen, yen, 0, True)
            distancedim.ReattachToDrawingView(DV)
            distancedim.MeasurementAxisEx = 1
            distancedim.MeasurementAxisDirection = False
            distancedim.TrackDistance = -0.011
            dimstyle = distancedim.Style
            dimstyle.PrimaryDecimalRoundOff = SolidEdgeFrameworkSupport.DimDecimalRoundOffTypeConstants.igDimStyleDecimal1

            distancedim.Range(xrmin, yrmin, xrmax, yrmax)
            If Math.Round(yrmax, 6) <= Math.Round(ym, 6) Then
                distancedim.TrackDistance = -distancedim.TrackDistance
            End If
            If distancedim.Value < 0.025 Then
                distancedim.TerminatorPosition = True
                If conside = "left" Then
                    distancedim.BreakPosition = 3
                Else
                    distancedim.BreakPosition = 1
                End If
                distancedim.BreakDistance = 0.0045
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub SetNippleLengthFront(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headerline As SolidEdgeDraft.DVLine2d, nippleline As SolidEdgeDraft.DVLine2d, consys As ConSysData, circuit As CircuitData)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objSheet.Dimensions
        Dim lengthdim As SolidEdgeFrameworkSupport.Dimension
        Dim xc, yc, xm, ym, xmin, ymin, xmax, ymax, xs, ys, xe, ye, xrefh, yrefh, yrefn As Double
        Dim xlist, ylist As New List(Of Double)

        Try

            headerline.GetStartPoint(xs, ys)
            headerline.GetEndPoint(xe, ye)

            nippleline.GetStartPoint(xm, ym)
            nippleline.GetEndPoint(xc, yc)

            xmin = Math.Min(xs, xe)
            xmax = Math.Max(xs, xe)

            If circuit.ConnectionSide = "left" Then
                xrefh = xmax
            Else
                xrefh = xmin
            End If

            If nippleline.ModelMember.FileName.Contains("InletNipple") Or (consys.HasFTCon And ym > 0) Then
                yrefn = Math.Min(yc, ym)
            Else
                yrefn = Math.Max(yc, ym)
            End If

            yrefh = ye

            lengthdim = objdims.AddDistanceBetweenObjects(nippleline, xm, yrefn, 0, True, headerline, xrefh, yrefh, 0, True)
            lengthdim.ReattachToDrawingView(DV)

            lengthdim.MeasurementAxisEx = 1
            lengthdim.MeasurementAxisDirection = False
            lengthdim.TrackDistance = -0.025
            lengthdim.Range(xmin, ymin, xmax, ymax)
            If nippleline.ModelMember.FileName.Contains("OutletNipple") Or (consys.HasFTCon And ym < 0) Then
                If Math.Round(ymin, 6) < Math.Round(yrefn, 6) Then
                    lengthdim.TrackDistance = 0.025
                End If
            Else
                lengthdim.BreakDistance = 0.25
                lengthdim.TrackDistance = -0.02
                If Math.Round(ymax, 6) > Math.Round(yrefn, 6) Then
                    lengthdim.TrackDistance = 0.01
                End If
            End If
            If circuit.Pressure > 16 Then
                lengthdim.TrackDistance = -lengthdim.TrackDistance
            End If
            lengthdim.Style.PrimaryDecimalRoundOff = SolidEdgeFrameworkSupport.DimDecimalRoundOffTypeConstants.igDimStyleDecimal1

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub SetNippleDistance(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headerbottomline As SolidEdgeDraft.DVLine2d, headertopline As SolidEdgeDraft.DVLine2d,
                               nipplearcs As List(Of SolidEdgeDraft.DVArc2d), nipplecircs As List(Of SolidEdgeDraft.DVCircle2d), conside As String)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objSheet.Dimensions
        Dim distancedim As SolidEdgeFrameworkSupport.Dimension
        Dim refline As SolidEdgeDraft.DVLine2d
        Dim nipplearc As SolidEdgeDraft.DVArc2d
        Dim nipplecirc As SolidEdgeDraft.DVCircle2d
        Dim DimStyle As SolidEdgeFrameworkSupport.DimStyle
        Dim xc, yc, xb1, yb1, xb2, yb2, mp As Double
        Dim ylist As New List(Of Double)

        Try

            headerbottomline.GetEndPoint(xb1, yb1)
            headerbottomline.GetStartPoint(xb2, yb2)
            mp = 1

            If nipplecircs.Count > 0 Then
                'take the closest circle to the bottom line for first dimension
                For i As Integer = 0 To nipplecircs.Count - 1
                    nipplecircs(i).GetCenterPoint(xc, yc)
                    If yb1 < yc Then
                        ylist.Add(Math.Round(yc - yb1, 6))
                    Else
                        'nipple is closer to a second header
                        ylist.Add(10)
                    End If
                Next
                nipplecirc = nipplecircs(ylist.IndexOf(ylist.Min))
                nipplecirc.GetCenterPoint(xc, yc)

                If General.currentunit.ModelRangeSuffix = "FP" Or General.currentunit.ModelRangeSuffix = "WP" Then
                    mp += 1
                    If nipplecirc.ModelMember.FileName.Contains("InletNipple") Then
                        refline = headertopline
                        mp = -mp
                    Else
                        refline = headerbottomline
                    End If
                Else
                    refline = headerbottomline
                End If

                refline.GetEndPoint(xb1, yb1)
                refline.GetStartPoint(xb2, yb2)

                If (xb1 > xb2 And conside = "right") Or (xb1 < xb2 And conside = "left") Then
                    distancedim = objdims.AddDistanceBetweenObjects(refline, xb2, yb2, 0, False, nipplecirc, xc, yc, 0, False)
                Else
                    distancedim = objdims.AddDistanceBetweenObjects(refline, xb1, yb1, 0, False, nipplecirc, xc, yc, 0, False)
                End If
            Else
                'take the closest arc to the bottom line for first dimension
                For i As Integer = 0 To nipplearcs.Count - 1
                    nipplearcs(i).GetCenterPoint(xc, yc)
                    If yb1 < yc Then
                        ylist.Add(Math.Round(yc - yb1, 6))
                    Else
                        'nipple is closer to a second header
                        ylist.Add(10)
                    End If
                Next
                nipplearc = nipplearcs(ylist.IndexOf(ylist.Min))
                nipplearc.GetCenterPoint(xc, yc)
                If General.currentunit.ModelRangeSuffix = "FP" Or General.currentunit.ModelRangeSuffix = "WP" Then
                    mp += 1
                    If nipplearc.ModelMember.FileName.Contains("InletNipple") Then
                        refline = headertopline
                        mp = -mp
                    Else
                        refline = headerbottomline
                    End If
                Else
                    refline = headerbottomline
                End If

                refline.GetEndPoint(xb1, yb1)
                refline.GetStartPoint(xb2, yb2)

                If (xb1 > xb2 And conside = "right") Or (xb1 < xb2 And conside = "left") Then
                    distancedim = objdims.AddDistanceBetweenObjects(refline, xb2, yb2, 0, False, nipplearc, xc, yc, 0, False)
                Else
                    distancedim = objdims.AddDistanceBetweenObjects(refline, xb1, yb1, 0, False, nipplearc, xc, yc, 0, False)
                End If
            End If
            distancedim.ReattachToDrawingView(DV)
            distancedim.MeasurementAxisEx = 1
            distancedim.MeasurementAxisDirection = True
            distancedim.TrackDistance = -0.008 * mp
            distancedim.TerminatorPosition = True
            DimStyle = distancedim.Style
            DimStyle.PrimaryDecimalRoundOff = SolidEdgeFrameworkSupport.DimDecimalRoundOffTypeConstants.igDimStyleDecimal1

            'check if there is another nippletube


        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try


    End Sub

    Shared Sub SetVentDistance(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headerbottomline As SolidEdgeDraft.DVLine2d, headertopline As SolidEdgeDraft.DVLine2d,
                                ventellipses As List(Of SolidEdgeDraft.DVEllipse2d), ventarcs As List(Of SolidEdgeDraft.DVArc2d), ventnipples As List(Of SolidEdgeDraft.DVCircle2d), conside As String)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objSheet.Dimensions
        Dim distancedim As SolidEdgeFrameworkSupport.Dimension
        Dim refline As SolidEdgeDraft.DVLine2d
        Dim ventarc As SolidEdgeDraft.DVArc2d
        Dim ventellipse As SolidEdgeDraft.DVEllipse2d
        Dim DimStyle As SolidEdgeFrameworkSupport.DimStyle
        Dim ventobj As Object
        Dim xc, yc, xb1, yb1, xb2, yb2 As Double
        Dim ylist As New List(Of Double)

        Try
            headerbottomline.GetEndPoint(xb1, yb1)
            headerbottomline.GetStartPoint(xb2, yb2)

            If headerbottomline.ModelMember.FileName.Contains("InletHeader") Then
                refline = headertopline

                If ventellipses.Count > 0 Then
                    ventellipse = ventellipses(0)
                    ventellipse.GetCenterPoint(xc, yc)
                    ventobj = ventellipse
                ElseIf ventarcs.Count > 0 Then
                    ventarc = ventarcs(0)
                    ventarc.GetCenterPoint(xc, yc)
                    ventobj = ventarc
                Else
                    ventnipples(0).GetCenterPoint(xc, yc)
                    ventobj = ventnipples(0)
                End If
            Else
                refline = headerbottomline
                If ventellipses.Count > 0 Then
                    ventellipse = ventellipses(0)
                    ventellipse.GetCenterPoint(xc, yc)
                    ventobj = ventellipse
                Else
                    ventarc = ventarcs(0)
                    ventarc.GetCenterPoint(xc, yc)
                    ventobj = ventarc
                End If
            End If

            refline.GetEndPoint(xb1, yb1)
            refline.GetStartPoint(xb2, yb2)

            If (xb1 > xb2 And conside = "right") Or (xb1 < xb2 And conside = "left") Then
                distancedim = objdims.AddDistanceBetweenObjects(refline, xb2, yb2, 0, False, ventobj, xc, yc, 0, False)
            Else
                distancedim = objdims.AddDistanceBetweenObjects(refline, xb1, yb1, 0, False, ventobj, xc, yc, 0, False)
            End If

            distancedim.ReattachToDrawingView(DV)
            distancedim.MeasurementAxisEx = 1
            distancedim.MeasurementAxisDirection = True
            distancedim.TrackDistance = -0.008
            distancedim.TerminatorPosition = True
            distancedim.BreakPosition = SolidEdgeFrameworkSupport.DimBreakPositionConstants.igDimBreakLeft
            distancedim.BreakDistance = 0.007
            DimStyle = distancedim.Style
            DimStyle.PrimaryDecimalRoundOff = SolidEdgeFrameworkSupport.DimDecimalRoundOffTypeConstants.igDimStyleDecimal1
            SEDrawing.ControlDimDistance(distancedim, {"x", "smaller"})

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub SetBrineVentDistance(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headerbottomline As SolidEdgeDraft.DVLine2d, headertopline As SolidEdgeDraft.DVLine2d,
                                ventobj As Object, conside As String)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objSheet.Dimensions
        Dim distancedim As SolidEdgeFrameworkSupport.Dimension
        Dim refline As SolidEdgeDraft.DVLine2d
        Dim DimStyle As SolidEdgeFrameworkSupport.DimStyle
        Dim ventcircle As SolidEdgeDraft.DVCircle2d
        Dim ventarc As SolidEdgeDraft.DVArc2d
        Dim xc, yc, xb1, yb1, xb2, yb2 As Double
        Dim ylist As New List(Of Double)

        Try
            headerbottomline.GetEndPoint(xb1, yb1)
            headerbottomline.GetStartPoint(xb2, yb2)

            If headerbottomline.ModelMember.FileName.Contains("InletHeader") Then
                refline = headerbottomline
                ventarc = TryCast(ventobj, SolidEdgeDraft.DVArc2d)
                ventarc.GetCenterPoint(xc, yc)
            Else
                refline = headertopline
                ventcircle = TryCast(ventobj, SolidEdgeDraft.DVCircle2d)
                ventcircle.GetCenterPoint(xc, yc)
            End If

            refline.GetEndPoint(xb1, yb1)
            refline.GetStartPoint(xb2, yb2)

            If (xb1 < xb2 And conside = "right") Or (xb1 > xb2 And conside = "left") Then
                distancedim = objdims.AddDistanceBetweenObjects(refline, xb2, yb2, 0, False, ventobj, xc, yc, 0, False)
            Else
                distancedim = objdims.AddDistanceBetweenObjects(refline, xb1, yb1, 0, False, ventobj, xc, yc, 0, False)
            End If

            With distancedim
                .ReattachToDrawingView(DV)
                .MeasurementAxisEx = 1
                .MeasurementAxisDirection = True
                .TrackDistance = -0.008
                .TerminatorPosition = True
                .BreakPosition = SolidEdgeFrameworkSupport.DimBreakPositionConstants.igDimBreakLeft
                .BreakDistance = 0.007
            End With
            DimStyle = distancedim.Style
            DimStyle.PrimaryDecimalRoundOff = SolidEdgeFrameworkSupport.DimDecimalRoundOffTypeConstants.igDimStyleDecimal1

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub SetStutzenCircDistance(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headerbottomline As SolidEdgeDraft.DVLine2d, headertopline As SolidEdgeDraft.DVLine2d,
                                     partcircles As List(Of SolidEdgeDraft.DVCircle2d), osf As Double, conside As String, finnedheight As Double)
        Dim vertlists As New List(Of SolidEdgeDraft.DVLine2d)
        Dim xk, yk, xb, yb, xts, yts, xte, yte, yc0, zk, xref1, xref2 As Double
        Dim htype, ktype As Integer
        Dim ylist As New List(Of Double)
        Dim stutzenDVList As New List(Of StutzenDVElement)

        Try
            headertopline.GetStartPoint(xts, yts)
            headertopline.GetEndPoint(xte, yte)
            headerbottomline.GetEndPoint(xb, yb)
            xref1 = Math.Max(xte, xts)
            xref2 = Math.Min(xte, xts)

            For i As Integer = 0 To partcircles.Count - 1
                partcircles(i).GetKeyPoint(0, xk, yk, zk, ktype, htype)
                yc0 = Math.Round(yk * 1000, 3)

                If yc0 > yb * 1000 And yc0 < yts * 1000 Then
                    If ktype = 4 And ylist.IndexOf(yc0) = -1 And xk < xref1 And xk > xref2 Then
                        stutzenDVList.Add(New StutzenDVElement With {.DVElement = partcircles(i), .Ypos = yk})
                        ylist.Add(yc0)
                    End If
                End If

            Next

            If stutzenDVList.Count > 0 Then
                SetVerDims(objSheet, DV, headerbottomline, headertopline, stutzenDVList, "circle", osf, conside, finnedheight)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub SetStutzenElDistance(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headerbottomline As SolidEdgeDraft.DVLine2d, headertopline As SolidEdgeDraft.DVLine2d,
                                     partellipse As List(Of SolidEdgeDraft.DVEllipse2d), osf As Double, conside As String, finnedheight As Double)
        Dim vertlists As New List(Of SolidEdgeDraft.DVLine2d)
        Dim xk, yk, xb, yb, xts, yts, xte, yte, yc0, zk, xref1, xref2 As Double
        Dim htype, ktype As Integer
        Dim ylist As New List(Of Double)
        Dim stutzenDVList As New List(Of StutzenDVElement)

        Try
            headertopline.GetStartPoint(xts, yts)
            headertopline.GetEndPoint(xte, yte)
            headerbottomline.GetEndPoint(xb, yb)
            xref1 = Math.Max(xte, xts)
            xref2 = Math.Min(xte, xts)

            For i As Integer = 0 To partellipse.Count - 1
                partellipse(i).GetKeyPoint(0, xk, yk, zk, ktype, htype)
                yc0 = Math.Round(yk * 1000, 3)

                If yc0 > yb * 1000 And yc0 < yts * 1000 Then
                    If partellipse(i).SegmentedStyleCount > 1 And ylist.IndexOf(yc0) = -1 Then
                        stutzenDVList.Add(New StutzenDVElement With {.DVElement = partellipse(i), .Ypos = yk})
                        ylist.Add(yc0)
                    End If
                End If
            Next

            If stutzenDVList.Count > 0 Then
                SetVerDims(objSheet, DV, headerbottomline, headertopline, stutzenDVList, "ellipse", osf, conside, finnedheight)
            Else
                For i As Integer = 0 To partellipse.Count - 1
                    partellipse(i).GetKeyPoint(0, xk, yk, zk, ktype, htype)
                    yc0 = Math.Round(yk * 1000, 3)

                    If yc0 > yb * 1000 And yc0 < yts * 1000 Then
                        If partellipse(i).SegmentedStyleCount = 1 And ylist.IndexOf(yc0) = -1 And xk < xref1 And xk > xref2 Then
                            stutzenDVList.Add(New StutzenDVElement With {.DVElement = partellipse(i), .Ypos = yk})
                            ylist.Add(yc0)
                        End If
                    End If
                Next
                If stutzenDVList.Count > 0 Then
                    SetVerDims(objSheet, DV, headerbottomline, headertopline, stutzenDVList, "ellipse", osf, conside, finnedheight)
                End If
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub SetStutzenArcDistance(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headerbottomline As SolidEdgeDraft.DVLine2d, headertopline As SolidEdgeDraft.DVLine2d,
                                   headerarcs As List(Of SolidEdgeDraft.DVArc2d), osf As Double, conside As String, finnedheight As Double)
        Dim vertlists As New List(Of SolidEdgeDraft.DVLine2d)
        Dim xk, yk, xb, yb, xts, yts, xte, yte, xs, ys, yc0, xc, yc, zk, xref1, xref2 As Double
        Dim htype, ktype As Integer
        Dim ylist As New List(Of Double)
        Dim stutzenDVList As New List(Of StutzenDVElement)

        Try
            headertopline.GetStartPoint(xts, yts)
            headertopline.GetEndPoint(xte, yte)
            headerbottomline.GetEndPoint(xb, yb)
            xref1 = Math.Max(xte, xts)
            xref2 = Math.Min(xte, xts)

            For i As Integer = 0 To headerarcs.Count - 1
                headerarcs(i).GetKeyPoint(3, xk, yk, zk, ktype, htype)
                headerarcs(i).GetCenterPoint(xc, yc)
                headerarcs(i).GetStartPoint(xs, ys)
                If (Math.Abs(Math.Round(xs - xref1, 6)) < 0.001 And conside = "right") Or (conside = "left" And Math.Abs(Math.Round(xs - xref2, 6)) < 0.001) Then
                    yc0 = Math.Round(yk * 1000, 3)
                    If yk > yb And yk < yts And xk < xref1 And xk > xref2 And ktype = 32 And ylist.IndexOf(yc0) = -1 Then
                        stutzenDVList.Add(New StutzenDVElement With {.DVElement = headerarcs(i), .Ypos = yc})
                        ylist.Add(yc0)
                    End If
                End If
            Next

            If stutzenDVList.Count > 0 Then
                SetVerDims(objSheet, DV, headerbottomline, headertopline, stutzenDVList, "arc", osf, conside, finnedheight)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub SetVerDims(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headerbottomline As SolidEdgeDraft.DVLine2d, headertopline As SolidEdgeDraft.DVLine2d,
                       DVElements As List(Of StutzenDVElement), elementtype As String, osf As Double, conside As String, finnedheight As Double)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objSheet.Dimensions
        Dim distancedim, firstdim, lastdim As SolidEdgeFrameworkSupport.Dimension
        Dim objarc As SolidEdgeDraft.DVArc2d
        Dim objellipse As SolidEdgeDraft.DVEllipse2d
        Dim objcircle As SolidEdgeDraft.DVCircle2d
        Dim objarclist As New List(Of SolidEdgeDraft.DVArc2d)
        Dim objellipselist As New List(Of SolidEdgeDraft.DVEllipse2d)
        Dim objcirclist As New List(Of SolidEdgeDraft.DVCircle2d)
        Dim dimstyle As SolidEdgeFrameworkSupport.DimStyle
        Dim xb, yb, xt, yt, xc, yc, xc2, yc2, xmin1, ymin1, xmax1, ymax1, xmin2, ymin2, xmax2, ymax2, mmp, xte, yte As Double

        Try
            If conside = "left" Then
                mmp = -1
            Else
                mmp = 1
            End If
            If finnedheight < 500 And (General.currentunit.ModelRangeSuffix = "FP" Or General.currentunit.ModelRangeSuffix = "WP") Then
                mmp *= 0.5
            End If

            headertopline.GetStartPoint(xt, yt)
            headertopline.GetEndPoint(xte, yte)
            headerbottomline.GetEndPoint(xb, yb)

            Dim olist = From plist In DVElements Order By plist.Ypos

            For i As Integer = 0 To olist.Count - 1
                Select Case elementtype
                    Case "arc"
                        objarc = CType(olist(i).DVElement, SolidEdgeDraft.DVArc2d)
                        'check center point
                        objarc.GetCenterPoint(xc, yc)
                        If xc > Math.Min(xt, xte) And xc < Math.Max(xt, xte) Then
                        Else
                            objarclist.Add(objarc)
                        End If
                    Case "ellipse"
                        objellipse = CType(olist(i).DVElement, SolidEdgeDraft.DVEllipse2d)
                        objellipselist.Add(objellipse)
                    Case "circle"
                        objcircle = CType(olist(i).DVElement, SolidEdgeDraft.DVCircle2d)
                        objcirclist.Add(objcircle)
                End Select
            Next

            Select Case elementtype
                Case "arc"
                    objarclist(0).GetCenterPoint(xc, yc)
                    firstdim = objdims.AddDistanceBetweenObjects(objarclist(0), xc, yc, 0, True, headerbottomline, xb, yb, 0, False)
                    firstdim.ReattachToDrawingView(DV)
                    firstdim.MeasurementAxisEx = 1
                    firstdim.MeasurementAxisDirection = True
                    firstdim.TrackDistance = 0.25
                    firstdim.Range(xmin1, ymin1, xmax1, ymax1)
                    firstdim.TrackDistance *= osf * mmp
                    firstdim.TextScale = 0.75
                    If firstdim.Value < 0.025 Then
                        firstdim.TerminatorPosition = True
                        firstdim.BreakPosition = 1
                        firstdim.BreakDistance = 0.003
                    End If
                    dimstyle = firstdim.Style
                    dimstyle.OrigTerminatorSize = 0.5
                    dimstyle.TerminatorSize = 0.5
                    For i As Integer = 1 To objarclist.Count - 1
                        objarclist(i - 1).GetCenterPoint(xc, yc)
                        objarclist(i).GetCenterPoint(xc2, yc2)
                        distancedim = objdims.AddDistanceBetweenObjects(objarclist(i - 1), xc, yc, 0, True, objarclist(i), xc2, yc2, 0, True)
                        distancedim.ReattachToDrawingView(DV)
                        distancedim.MeasurementAxisEx = 1
                        distancedim.MeasurementAxisDirection = True
                        distancedim.TrackDistance = -0.25
                        distancedim.Range(xmin2, ymin2, xmax2, ymax2)
                        If Math.Abs(Math.Round((xmax1 - xmax2) * 1000, 3)) > 1 Then
                            distancedim.TrackDistance = 0.25
                        End If
                        distancedim.TextScale = 0.75
                        distancedim.TrackDistance *= osf * mmp
                        dimstyle = distancedim.Style
                        dimstyle.OrigTerminatorSize = 0.5
                        dimstyle.TerminatorSize = 0.5
                    Next
                    If objarclist.Count > 1 Then
                        lastdim = objdims.AddDistanceBetweenObjects(objarclist.Last, xc2, yc2, 0, True, headertopline, xt, yt, 0, False)
                        lastdim.ReattachToDrawingView(DV)
                        lastdim.MeasurementAxisEx = 1
                        lastdim.MeasurementAxisDirection = True
                        lastdim.TrackDistance = -0.25
                        lastdim.Range(xmin2, ymin2, xmax2, ymax2)
                        If Math.Abs(Math.Round((xmax1 - xmax2) * 1000, 3)) > 1 Then
                            lastdim.TrackDistance = 0.25
                        End If
                        lastdim.TrackDistance *= osf * mmp
                        lastdim.TextScale = 0.75
                        lastdim.DisplayType = 6
                        lastdim.TerminatorPosition = True
                        lastdim.BreakPosition = 1
                        lastdim.BreakDistance = 0.003
                        dimstyle = lastdim.Style
                        dimstyle.OrigTerminatorSize = 0.5
                        dimstyle.TerminatorSize = 0.5
                    End If
                Case "ellipse"
                    objellipselist(0).GetCenterPoint(xc, yc)
                    firstdim = objdims.AddDistanceBetweenObjects(objellipselist(0), xc, yc, 0, True, headerbottomline, xb, yb, 0, False)
                    firstdim.ReattachToDrawingView(DV)
                    firstdim.MeasurementAxisEx = 1
                    firstdim.MeasurementAxisDirection = True
                    firstdim.TrackDistance = 0.2
                    firstdim.Range(xmin1, ymin1, xmax1, ymax1)
                    firstdim.TrackDistance *= osf * mmp
                    firstdim.TextScale = 0.75
                    If firstdim.Value < 0.025 Then
                        firstdim.TerminatorPosition = True
                        firstdim.BreakPosition = 1
                        firstdim.BreakDistance = 0.003
                    End If
                    dimstyle = firstdim.Style
                    dimstyle.OrigTerminatorSize = 0.5
                    dimstyle.TerminatorSize = 0.5
                    For i As Integer = 1 To objellipselist.Count - 1
                        objellipselist(i - 1).GetCenterPoint(xc, yc)
                        objellipselist(i).GetCenterPoint(xc2, yc2)
                        distancedim = objdims.AddDistanceBetweenObjects(objellipselist(i - 1), xc, yc, 0, True, objellipselist(i), xc2, yc2, 0, True)
                        distancedim.ReattachToDrawingView(DV)
                        distancedim.MeasurementAxisEx = 1
                        distancedim.MeasurementAxisDirection = True
                        distancedim.TrackDistance = -0.2
                        distancedim.Range(xmin2, ymin2, xmax2, ymax2)
                        If Math.Abs(Math.Round((xmax1 - xmax2) * 1000, 3)) > 1 Then
                            distancedim.TrackDistance = 0.2
                        End If
                        distancedim.TextScale = 0.75
                        distancedim.TrackDistance *= osf * mmp
                        dimstyle = distancedim.Style
                        dimstyle.OrigTerminatorSize = 0.5
                        dimstyle.TerminatorSize = 0.5
                    Next
                    If objellipselist.Count > 1 Then
                        lastdim = objdims.AddDistanceBetweenObjects(objellipselist.Last, xc2, yc2, 0, True, headertopline, xt, yt, 0, False)
                        lastdim.ReattachToDrawingView(DV)
                        lastdim.MeasurementAxisEx = 1
                        lastdim.MeasurementAxisDirection = True
                        lastdim.TrackDistance = -0.2
                        lastdim.Range(xmin2, ymin2, xmax2, ymax2)
                        If Math.Abs(Math.Round((xmax1 - xmax2) * 1000, 3)) > 1 Then
                            lastdim.TrackDistance = 0.2
                        End If
                        lastdim.TrackDistance *= osf * mmp
                        lastdim.TextScale = 0.75
                        lastdim.DisplayType = 6
                        lastdim.TerminatorPosition = True
                        lastdim.BreakPosition = 1
                        lastdim.BreakDistance = 0.003
                        dimstyle = lastdim.Style
                        dimstyle.OrigTerminatorSize = 0.5
                        dimstyle.TerminatorSize = 0.5
                    End If
                Case "circle"
                    objcirclist(0).GetCenterPoint(xc, yc)
                    firstdim = objdims.AddDistanceBetweenObjects(objcirclist(0), xc, yc, 0, True, headerbottomline, xb, yb, 0, False)
                    firstdim.ReattachToDrawingView(DV)
                    firstdim.MeasurementAxisEx = 1
                    firstdim.MeasurementAxisDirection = True
                    firstdim.TrackDistance = 0.27
                    firstdim.Range(xmin1, ymin1, xmax1, ymax1)
                    firstdim.TrackDistance *= osf * mmp
                    firstdim.TextScale = 0.75
                    If firstdim.Value < 0.025 Then
                        firstdim.TerminatorPosition = True
                        firstdim.BreakPosition = 1
                        firstdim.BreakDistance = 0.003
                    End If
                    dimstyle = firstdim.Style
                    dimstyle.OrigTerminatorSize = 0.5
                    dimstyle.TerminatorSize = 0.5
                    For i As Integer = 1 To objcirclist.Count - 1
                        objcirclist(i - 1).GetCenterPoint(xc, yc)
                        objcirclist(i).GetCenterPoint(xc2, yc2)
                        distancedim = objdims.AddDistanceBetweenObjects(objcirclist(i - 1), xc, yc, 0, True, objcirclist(i), xc2, yc2, 0, True)
                        distancedim.ReattachToDrawingView(DV)
                        distancedim.MeasurementAxisEx = 1
                        distancedim.MeasurementAxisDirection = True
                        distancedim.TrackDistance = -0.27
                        distancedim.Range(xmin2, ymin2, xmax2, ymax2)
                        If Math.Abs(Math.Round((xmax1 - xmax2) * 1000, 3)) > 1 Then
                            distancedim.TrackDistance = 0.27
                        End If
                        distancedim.TextScale = 0.75
                        distancedim.TrackDistance *= osf * mmp
                        dimstyle = distancedim.Style
                        dimstyle.OrigTerminatorSize = 0.5
                        dimstyle.TerminatorSize = 0.5
                    Next
                    If objcirclist.Count > 1 Then
                        lastdim = objdims.AddDistanceBetweenObjects(objcirclist.Last, xc2, yc2, 0, True, headertopline, xt, yt, 0, False)
                        lastdim.ReattachToDrawingView(DV)
                        lastdim.MeasurementAxisEx = 1
                        lastdim.MeasurementAxisDirection = True
                        lastdim.TrackDistance = -0.27
                        lastdim.Range(xmin2, ymin2, xmax2, ymax2)
                        If Math.Abs(Math.Round((xmax1 - xmax2) * 1000, 3)) > 1 Then
                            lastdim.TrackDistance = 0.27
                        End If
                        lastdim.TrackDistance *= osf * mmp
                        lastdim.TextScale = 0.75
                        lastdim.DisplayType = 6
                        lastdim.TerminatorPosition = True
                        lastdim.BreakPosition = 1
                        lastdim.BreakDistance = 0.003
                        dimstyle = lastdim.Style
                        dimstyle.OrigTerminatorSize = 0.5
                        dimstyle.TerminatorSize = 0.5
                    End If
            End Select

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub SetPipeDistance(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headerbottomline As SolidEdgeDraft.DVLine2d, headertopline As SolidEdgeDraft.DVLine2d,
                             pipelines As List(Of SolidEdgeDraft.DVLine2d), conside As String)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objSheet.Dimensions
        Dim distancedim As SolidEdgeFrameworkSupport.Dimension
        Dim horpipeline, vertpipeline As SolidEdgeDraft.DVLine2d
        Dim vertlists As New List(Of SolidEdgeDraft.DVLine2d)
        Dim dimstyle As SolidEdgeFrameworkSupport.DimStyle
        Dim xc, yc, xbe, ybe, xbs, ybs, xt, yt, xs, ys, xe, ye, zc, trackdistance, xmin, ymin, xmax, ymax As Double
        Dim htype, ktype As Integer
        Dim xlist As New List(Of Double)

        Try
            headerbottomline.GetEndPoint(xbe, ybe)
            headerbottomline.GetStartPoint(xbs, ybs)
            headertopline.GetEndPoint(xt, yt)


            For i As Integer = 0 To pipelines.Count - 1
                pipelines(i).GetStartPoint(xs, ys)
                pipelines(i).GetEndPoint(xe, ye)
                pipelines(i).ModelMember.ShowTangentEdges = True
                If ys < yt And ys > ybe Then
                    'vertical for distance to header end
                    If Math.Abs(Math.Round((xs - xe) * 1000, 3)) < 0.1 Then
                        Debug.Print(Math.Round(pipelines(i).Length * 1000, 1).ToString)
                        If Math.Round(pipelines(i).Length * 1000, 1) = 6 And vertpipeline Is Nothing Then
                            vertpipeline = pipelines(i)
                        ElseIf Math.Round(pipelines(i).Length * 1000, 1) = 85 Then
                            xlist.Add(xs)
                            vertlists.Add(pipelines(i))
                        End If
                    Else
                        'horizontal line
                        If Math.Round(pipelines(i).Length * 1000, 1) = 20 Then
                            horpipeline = pipelines(i)
                        End If
                    End If
                End If
            Next

            vertpipeline.GetKeyPoint(2, xc, yc, zc, ktype, htype)

            distancedim = objdims.AddDistanceBetweenObjects(vertpipeline, xc, yc, 0, True, headerbottomline, xbe, ybe, 0, False)
            distancedim.ReattachToDrawingView(DV)
            distancedim.MeasurementAxisEx = 1
            distancedim.MeasurementAxisDirection = True
            distancedim.TrackDistance = -0.0175
            dimstyle = distancedim.Style
            dimstyle.PrimaryDecimalRoundOff = SolidEdgeFrameworkSupport.DimDecimalRoundOffTypeConstants.igDimStyleDecimal1

            vertpipeline = Nothing

            If conside = "left" Then
                vertpipeline = vertlists(xlist.IndexOf(xlist.Max))
                vertpipeline.GetEndPoint(xe, ye)
                vertpipeline.GetStartPoint(xs, ys)
                trackdistance = -0.008
                distancedim.TrackDistance = -distancedim.TrackDistance
            Else
                vertpipeline = vertlists(xlist.IndexOf(xlist.Min))
                vertpipeline.GetStartPoint(xe, ye)
                vertpipeline.GetEndPoint(xs, ys)
                trackdistance = 0.008
            End If

            distancedim = Nothing

            If (conside = "right" And xbe > xbs) Or (conside = "left" And xbs > xbe) Then
                distancedim = objdims.AddDistanceBetweenObjects(vertpipeline, xe, ye, 0, True, headerbottomline, xbs, ybs, 0, True)
            Else
                distancedim = objdims.AddDistanceBetweenObjects(vertpipeline, xe, ye, 0, True, headerbottomline, xbe, ybe, 0, True)
            End If

            distancedim.ReattachToDrawingView(DV)
            distancedim.MeasurementAxisEx = 1
            distancedim.MeasurementAxisDirection = False
            distancedim.TerminatorPosition = True
            distancedim.BreakPosition = 3
            distancedim.BreakDistance = 0.005
            distancedim.TrackDistance = trackdistance

            distancedim.Range(xmin, ymin, xmax, ymax)
            If ymax <= Math.Max(ye, ys) Then
                distancedim.TrackDistance = -trackdistance
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub SetHeaderOffset(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headercirc As SolidEdgeDraft.DVCircle2d, ctline As SolidEdgeDraft.DVLine2d, conside As String)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objSheet.Dimensions
        Dim distancedim As SolidEdgeFrameworkSupport.Dimension
        Dim xc, yc, xm, ym, z, xmin, ymin, xmax, ymax As Double
        Dim htype, ktype As Integer

        Try
            headercirc.GetCenterPoint(xc, yc)
            ctline.GetKeyPoint(2, xm, ym, z, ktype, htype)

            If Math.Abs(xm - xc) > 0.001 Then
                distancedim = objdims.AddDistanceBetweenObjects(ctline, xm, ym, 0, True, headercirc, xc, yc, 0, True)
                distancedim.ReattachToDrawingView(DV)

                If conside = "left" Then
                    distancedim.BreakPosition = SolidEdgeFrameworkSupport.DimBreakPositionConstants.igDimBreakLeft
                Else
                    distancedim.BreakPosition = SolidEdgeFrameworkSupport.DimBreakPositionConstants.igDimBreakRight
                End If
                distancedim.BreakDistance = 0.005
                distancedim.TerminatorPosition = True
                distancedim.MeasurementAxisEx = 1
                distancedim.MeasurementAxisDirection = False
                distancedim.TrackDistance = 0.026
                distancedim.Range(xmin, ymin, xmax, ymax)
                If Math.Round(ymax, 6) <= Math.Round(ym, 6) Then
                    distancedim.TrackDistance = -0.026
                End If
            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub SetCTOverhang(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, finline As SolidEdgeDraft.DVLine2d, ctline As SolidEdgeDraft.DVLine2d, conside As String)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objSheet.Dimensions
        Dim distancedim As SolidEdgeFrameworkSupport.Dimension
        Dim xc, yc, xs, ys, xe, ye, xref, yref As Double
        Dim comparer As String

        Try
            finline.GetStartPoint(xs, ys)
            finline.GetEndPoint(xe, ye)

            If (xs > xe And conside = "right") Or (xs < xe And conside = "left") Then
                xref = xs
                yref = ys
                comparer = "bigger"
                Debug.Print("startpoint for CToverhang dimension")
            Else
                xref = xe
                yref = ye
                comparer = "smaller"
                Debug.Print("endpoint for CToverhang dimension")
            End If

            ctline.GetStartPoint(xc, yc)

            distancedim = objdims.AddDistanceBetweenObjects(finline, xref, yref, 0, True, ctline, xc, yc, 0, False)

            With distancedim
                .ReattachToDrawingView(DV)
                .BreakPosition = SolidEdgeFrameworkSupport.DimBreakPositionConstants.igDimBreakLeft
                .BreakDistance = 0.01
                .TerminatorPosition = True
                .MeasurementAxisEx = 1
                .MeasurementAxisDirection = False
                .TrackDistance = 0.005
                If .Value > 0.06 OrElse .Value < 0.02 Then
                    .MeasurementAxisDirection = True
                End If
            End With

            'using wrong coretube for reference (FPTopViewDim)
            SEDrawing.ControlDimDistance(distancedim, {"x", comparer})

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub SetHeaderDistance(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headercirc As SolidEdgeDraft.DVCircle2d, finline As SolidEdgeDraft.DVLine2d, conside As String)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objSheet.Dimensions
        Dim distancedim As SolidEdgeFrameworkSupport.Dimension
        Dim xlist, ylist As New List(Of Double)
        Dim xc, yc, xm, ym, z, xmin, ymin, xmax, ymax, xs, ys, xe, ye, xref, yref As Double
        Dim htype, ktype As Integer

        Try
            headercirc.GetCenterPoint(xc, yc)
            finline.GetStartPoint(xs, ys)
            finline.GetEndPoint(xe, ye)

            If (xs > xe And conside = "right") Or (xs < xe And conside = "left") Then
                xref = xs
                yref = ys
            Else
                xref = xe
                yref = ye
            End If

            For i As Integer = 0 To headercirc.KeyPointCount - 1
                headercirc.GetKeyPoint(i, xm, ym, z, ktype, htype)
                If ym - yc > headercirc.Radius / 2 Then
                    Debug.Print("keypoint: " + i.ToString + " x: " + Math.Round(xm, 6).ToString + " / y:" + Math.Round(ym, 6).ToString)
                    Exit For
                End If
            Next

            'minimum distance between header and tube sheet
            distancedim = objdims.AddDistanceBetweenObjectsEX(finline, xref, yref, 0, False, False, headercirc, xm, ym, 0, False, True)
            distancedim.ReattachToDrawingView(DV)

            distancedim.MeasurementAxisEx = 1
            distancedim.MeasurementAxisDirection = True
            distancedim.TrackDistance = 0.0125
            distancedim.Range(xmin, ymin, xmax, ymax)
            If (Math.Round(xmin, 9) >= Math.Round(xref, 6) And conside = "left") Or (Math.Round(xmax, 6) <= Math.Round(xref, 6) And conside = "right") Then
                distancedim.TrackDistance = -0.0125
            End If

            'distance between tube sheet and center of header
            distancedim = objdims.AddDistanceBetweenObjects(finline, xref, yref, 0, True, headercirc, xc, yc, 0, True)
            distancedim.ReattachToDrawingView(DV)

            distancedim.MeasurementAxisEx = 1
            distancedim.MeasurementAxisDirection = True
            distancedim.TrackDistance = 0.02
            distancedim.Range(xmin, ymin, xmax, ymax)
            If (Math.Round(xmin, 9) >= Math.Round(xref, 6) And conside = "left") Or (Math.Round(xmax, 6) <= Math.Round(xref, 6) And conside = "right") Then
                distancedim.TrackDistance = -0.02
            End If
            distancedim.DisplayType = 6

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub SetNippleLengthTop(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headercirc As SolidEdgeDraft.DVCircle2d, nippleline As SolidEdgeDraft.DVLine2d, conside As String)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objSheet.Dimensions
        Dim distancedim As SolidEdgeFrameworkSupport.Dimension
        Dim xc, yc, xm, ym, z, xmin, ymin, xmax, ymax, xs, ys, xe, ye, xref, yref As Double
        Dim htype, ktype As Integer

        Try
            headercirc.GetCenterPoint(xc, yc)
            nippleline.GetStartPoint(xs, ys)
            nippleline.GetEndPoint(xe, ye)

            If ys > ye Then
                xref = xs
                yref = ys
            Else
                xref = xe
                yref = ye
            End If

            For i As Integer = 0 To headercirc.KeyPointCount - 1
                headercirc.GetKeyPoint(i, xm, ym, z, ktype, htype)
                If Math.Abs(xc - xm) > headercirc.Radius / 2 And ((xm < xc And conside = "right") Or (xm > xc And conside = "left")) Then
                    Exit For
                End If
            Next

            distancedim = objdims.AddDistanceBetweenObjectsEX(nippleline, xref, yref, 0, True, True, headercirc, xm, ym, 0, False, True)

            distancedim.ReattachToDrawingView(DV)

            If conside = "left" Then
                distancedim.MeasurementAxisEx = 1
            Else
                distancedim.MeasurementAxisEx = 3
            End If

            distancedim.MeasurementAxisDirection = False
            distancedim.TrackDistance = 0.01
            distancedim.Range(xmin, ymin, xmax, ymax)
            If ymin < Math.Max(ys, ye) Then
                distancedim.TrackDistance = -0.01
            End If
            distancedim.Style.PrimaryDecimalRoundOff = SolidEdgeFrameworkSupport.DimDecimalRoundOffTypeConstants.igDimStyleDecimal_1

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub FPVentDistance(objsheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, ventline As SolidEdgeDraft.DVLine2d, finline As SolidEdgeDraft.DVLine2d, conside As String)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objsheet.Dimensions
        Dim distancedim As SolidEdgeFrameworkSupport.Dimension
        Dim xc, yc, z, xsf, ysf, xef, yef, xreff, yreff, trackdistance, xvs, yvs, xve, yve As Double
        Dim ktype, htype As Integer
        Dim comparer As String

        Try
            finline.GetStartPoint(xsf, ysf)
            finline.GetEndPoint(xef, yef)

            If conside = "left" Then
                xreff = Math.Min(xsf, xef)
                trackdistance = 0.0075
                comparer = "smaller"
            Else
                xreff = Math.Max(xsf, xef)
                trackdistance = -0.0075
                comparer = "bigger"
            End If
            yreff = yef

            ventline.GetStartPoint(xvs, yvs)
            ventline.GetEndPoint(xve, yve)
            ventline.GetKeyPoint(2, xc, yc, z, ktype, htype)

            distancedim = objdims.AddDistanceBetweenObjects(finline, xreff, yreff, 0, True, ventline, xc, yc, z, True)
            With distancedim
                .ReattachToDrawingView(DV)
                .MeasurementAxisEx = 1
                .MeasurementAxisDirection = True
                .TerminatorPosition = True
                .TrackDistance = trackdistance
                .BreakPosition = SolidEdgeFrameworkSupport.DimBreakPositionConstants.igDimBreakLeft
                .BreakDistance = 0.005
            End With

            SEDrawing.ControlDimDistance(distancedim, {"x", comparer})

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub FPSetHeaderDistance(objsheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, nippleline As SolidEdgeDraft.DVLine2d, finline As SolidEdgeDraft.DVLine2d,
                                headercirc As SolidEdgeDraft.DVCircle2d, conside As String, headertype As String, fintype As String)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objsheet.Dimensions
        Dim distancedim As SolidEdgeFrameworkSupport.Dimension
        Dim xlist, ylist As New List(Of Double)
        Dim xc, yc, xsn, ysn, xen, yen, xmin, ymin, xmax, ymax, xsf, ysf, xef, yef, xreff, yreff, xrefn, yrefn, trackdistance, noffset As Double

        Try
            finline.GetStartPoint(xsf, ysf)
            finline.GetEndPoint(xef, yef)
            nippleline.GetStartPoint(xsn, ysn)
            nippleline.GetEndPoint(xen, yen)
            headercirc.GetCenterPoint(xc, yc)

            If headertype = "outlet" Then
                trackdistance = 0.0135
                If conside = "left" Then
                    xreff = Math.Max(xsf, xef)
                Else
                    xreff = Math.Min(xsf, xef)
                End If
            Else
                trackdistance = -0.0135
                If conside = "left" Then
                    xreff = Math.Min(xsf, xef)
                Else
                    xreff = Math.Max(xsf, xef)
                End If
            End If
            If fintype = "N" Then
                noffset = 0.013
            Else
                noffset = 0
            End If
            'horizontal lines → y same value for start and end
            yreff = ysf

            If conside = "left" Then
                xrefn = Math.Min(xsn, xen)
            Else
                xrefn = Math.Max(xsn, xen)
            End If
            yrefn = ysn

            'minimum distance between header and tube sheet
            distancedim = objdims.AddDistanceBetweenObjects(finline, xreff, yreff, 0, True, nippleline, xrefn, yrefn, 0, False)
            distancedim.ReattachToDrawingView(DV)
            distancedim.MeasurementAxisEx = 1
            distancedim.MeasurementAxisDirection = True
            distancedim.TerminatorPosition = False
            distancedim.TrackDistance = -trackdistance


            distancedim.Range(xmin, ymin, xmax, ymax)
            If (Math.Round(xmax, 6) > Math.Round(xreff, 6) And conside = "right" And headertype = "outlet") Or (Math.Round(xmin + 0.012, 6) < Math.Round(xreff, 6) And conside = "left" And headertype = "outlet") Then
                distancedim.TrackDistance = -distancedim.TrackDistance
            ElseIf (Math.Round(xmax, 6) <= Math.Round(xreff, 6) And conside = "right" And headertype = "inlet") Or (Math.Round(xmin, 6) >= Math.Round(xreff, 6) And conside = "left" And headertype = "inlet") Then
                distancedim.TrackDistance = -distancedim.TrackDistance
            End If

            'distance between tube sheet and center of header
            distancedim = objdims.AddDistanceBetweenObjects(finline, xreff, yreff, 0, True, headercirc, xc, yc, 0, True)
            distancedim.ReattachToDrawingView(DV)
            distancedim.MeasurementAxisEx = 1
            distancedim.MeasurementAxisDirection = True
            distancedim.TrackDistance = 2 * trackdistance
            distancedim.Range(xmin, ymin, xmax, ymax)
            distancedim.DisplayType = 6
            distancedim.Style.PrimaryDecimalRoundOff = SolidEdgeFrameworkSupport.DimDecimalRoundOffTypeConstants.igDimStyleDecimal_1

            If (Math.Round(xmax, 6) > Math.Round(xreff + 0.025, 6) And conside = "right" And headertype = "outlet") Or (Math.Round(xmin + 0.0125 + noffset, 6) < Math.Round(xreff, 6) And conside = "left" And headertype = "outlet") Then
                distancedim.TrackDistance = -distancedim.TrackDistance
            ElseIf (Math.Round(xmax, 6) <= Math.Round(xreff, 6) And conside = "right" And headertype = "inlet") Or (Math.Round(xmin, 6) >= Math.Round(xreff, 6) And conside = "left" And headertype = "inlet") Then
                distancedim.TrackDistance = -distancedim.TrackDistance
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub HorHeaderDim(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headerlines() As SolidEdgeDraft.DVLine2d, circuit As CircuitData)
        Dim dimline, notdimline As SolidEdgeDraft.DVLine2d
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objSheet.Dimensions
        Dim diadim, lengthdim As SolidEdgeFrameworkSupport.Dimension
        Dim xe, ye, xs, ys, xmin, ymin, xmax, ymax As Double
        Dim breakpos, mp As Integer

        Try
            If circuit.ConnectionSide = "left" Then
                'right line
                notdimline = headerlines(0)
                dimline = headerlines(1)
                breakpos = 3
            Else
                'left line
                dimline = headerlines(0)
                notdimline = headerlines(1)
                breakpos = 1
            End If

            dimline.GetStartPoint(xs, ys)
            dimline.GetEndPoint(xe, ye)

            diadim = objdims.AddDistanceBetweenObjects(dimline, xs, ys, 0, True, dimline, xe, ye, 0, True)

            With diadim
                .PrefixString = "%DI"
                mp = 1
                If circuit.ConnectionSide = "left" Then
                    .BreakPosition = 3
                    If circuit.CircuitType <> "Defrost" And dimline.ModelMember.FileName.Contains("InletHeader") Then
                        mp = -1
                    End If
                Else
                    diadim.BreakPosition = 1
                End If
                .BreakDistance = 0.003
                .TerminatorPosition = True
                .ReattachToDrawingView(DV)
                .TrackDistance = 0.01
                .Range(xmin, ymin, xmax, ymax)
                If (circuit.ConnectionSide = "right" And Math.Round(xmax, 6) > Math.Round(xs, 6)) Or (circuit.ConnectionSide = "left" And Math.Round(xmin, 6) < Math.Round(xs, 6)) Then
                    .TrackDistance *= -1
                    If circuit.CircuitType = "Defrost" Then
                        .TrackDistance *= -1
                    End If
                End If
            End With

            Dim x2s, x2e, y2s, y2e As Double
            notdimline.GetStartPoint(x2s, y2s)
            notdimline.GetEndPoint(x2e, y2e)

            lengthdim = objdims.AddDistanceBetweenObjects(dimline, xe, ye, 0, True, notdimline, x2e, y2e, 0, True)

            With lengthdim
                .ReattachToDrawingView(DV)
                .TrackDistance = 0.015 * mp
                If ye > 0 Then
                    .TrackDistance *= -1
                End If
                .BreakPosition = breakpos
                .BreakDistance = 0.0075
                If dimline.ModelMember.FileName.Contains("OutletHeader") And circuit.CircuitType = "Defrost" Then
                    .TrackDistance *= -1
                End If
            End With

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function HorHeaderDistance(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headerlines() As SolidEdgeDraft.DVLine2d, conside As String, finline As SolidEdgeDraft.DVLine2d, conpoint As String) As SolidEdgeFrameworkSupport.Dimension
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objSheet.Dimensions
        Dim distancedim As SolidEdgeFrameworkSupport.Dimension
        Dim headerline As SolidEdgeDraft.DVLine2d
        Dim xs, xe, ys, ye, xmin, xmax, ymin, ymax, xm, ym, zm As Double
        Dim htype, ktype, mp As Integer

        Try

            If conside = "left" Then
                headerline = headerlines(0)
                mp = 1
            Else
                headerline = headerlines(1)
                mp = -1
            End If
            headerline.GetKeyPoint(2, xm, ym, zm, ktype, htype)

            finline.GetStartPoint(xs, ys)
            finline.GetEndPoint(xe, ye)

            ymax = Math.Max(ys, ye)
            ymin = Math.Min(ys, ye)

            If conpoint = "top" Then
                distancedim = objdims.AddDistanceBetweenObjects(finline, xs, ymax, 0, True, headerline, xm, ym, 0, True)
                mp *= -1
            Else
                distancedim = objdims.AddDistanceBetweenObjects(finline, xs, ymin, 0, True, headerline, xm, ym, 0, True)
            End If

            With distancedim
                .MeasurementAxisDirection = True
                .MeasurementAxisEx = 1
                .TrackDistance = -0.01 * mp

                .Range(xmin, ymin, xmax, ymax)
                If (conside = "left" And Math.Round(xmin, 6) > Math.Round(xm, 6)) Or (conside = "right" And Math.Round(xmin, 6) <= Math.Round(xm, 6)) Then
                    .TrackDistance *= -1
                End If
                .ReattachToDrawingView(DV)

                If .Value < 0.05 Then
                    .TerminatorPosition = True
                    .BreakPosition = 3
                    .BreakDistance = 0.005
                End If
            End With

        Catch ex As Exception

        End Try

        Return distancedim
    End Function

    Shared Function StutzenHorPosition(objsheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headerdimlines() As SolidEdgeDraft.DVLine2d, headerframe() As Double,
                                       circlist As List(Of SolidEdgeDraft.DVCircle2d), circuit As CircuitData, headertype As String, rowcount As Integer,
                                       figure As Integer) As List(Of SolidEdgeDraft.DVCircle2d)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objsheet.Dimensions
        Dim firstdim, distancedim, lastdim As SolidEdgeFrameworkSupport.Dimension
        Dim dimstyle As SolidEdgeFrameworkSupport.DimStyle
        Dim sortedcirclist As New List(Of SolidEdgeDraft.DVCircle2d)
        Dim stutzenCirclist As New List(Of StutzenDVElement)
        Dim xclist As New List(Of Double)
        Dim xc, yc, xh1, yh1, xc2, yc2, xe1, ye1, xh2, yh2, xe2, ye2, xmin, ymin, xmax, ymax, xminf, yminf, xmaxf, ymaxf, xminl, yminl, xmaxl, ymaxl, trackdistance, yref As Double
        Dim breakpos As Integer = 1

        Try
            For i As Integer = 0 To circlist.Count - 1
                Dim addtolist As Boolean = False

                'if figure 4 → circ has to be inside the header frame
                If Math.Round(circlist(i).Diameter * 1000, 3) = circuit.CoreTube.Diameter Then
                    circlist(i).GetCenterPoint(xc, yc)
                    If figure = 4 Then
                        If yc < headerframe(3) And yc > headerframe(1) Then
                            addtolist = True
                        End If
                    Else
                        addtolist = True
                    End If
                End If

                If addtolist Then
                    xc = Math.Round(xc, 6)
                    If xclist.IndexOf(xc) = -1 Then
                        stutzenCirclist.Add(New StutzenDVElement With {.DVElement = circlist(i), .Ypos = xc})
                        xclist.Add(xc)
                    End If
                End If
            Next

            Dim sortedlist = From plist In stutzenCirclist Order By plist.Ypos

            If (circuit.ConnectionSide = "right" And circuit.CircuitType = "Defrost") OrElse (circuit.ConnectionSide = "left" And circuit.CircuitType <> "Defrost") Then
                sortedlist.Reverse
                'breakpos = 3
            End If

            trackdistance = 0.015
            If figure = 8 Then
                trackdistance *= -1
            ElseIf figure = 4 And rowcount = 3 Then
                trackdistance *= 1.5
            ElseIf figure = 5 And circuit.CircuitType = "Defrost" And headertype = "outlet" Then
                trackdistance *= -1
            End If

            For i As Integer = 0 To sortedlist.Count - 1
                sortedcirclist.Add(sortedlist(i).DVElement)
            Next

            'first circle and headerdimline
            sortedcirclist(0).GetCenterPoint(xc, yc)
            headerdimlines(0).GetStartPoint(xh1, yh1)
            headerdimlines(0).GetEndPoint(xe1, ye1)
            headerdimlines(1).GetStartPoint(xh2, yh2)
            headerdimlines(1).GetEndPoint(xe2, ye2)
            firstdim = objdims.AddDistanceBetweenObjects(sortedcirclist(0), xc, yc, 0, True, headerdimlines(0), xh1, yh1, 0, True)
            With firstdim
                .MeasurementAxisDirection = False
                .MeasurementAxisEx = 1
                .TrackDistance = Math.Abs(trackdistance)
                .TextScale = 0.75
                If .Value < 0.025 Then
                    .TerminatorPosition = True
                    .BreakPosition = breakpos
                    .BreakDistance = 0.003
                End If
                .Range(xminf, yminf, xmaxf, ymaxf)
                .ReattachToDrawingView(DV)
                dimstyle = .Style
            End With
            dimstyle.OrigTerminatorSize = 0.5
            dimstyle.TerminatorSize = 0.5
            ymin = yminf
            ymax = ymaxf

            For i As Integer = 1 To sortedcirclist.Count - 1
                sortedcirclist(i - 1).GetCenterPoint(xc, yc)
                sortedcirclist(i).GetCenterPoint(xc2, yc2)
                distancedim = objdims.AddDistanceBetweenObjects(sortedcirclist(i - 1), xc, yc, 0, True, sortedcirclist(i), xc2, yc2, 0, True)
                With distancedim
                    .ReattachToDrawingView(DV)
                    .MeasurementAxisEx = 1
                    .MeasurementAxisDirection = False
                    .TextScale = 0.75
                    .TrackDistance = trackdistance
                    .Range(xmin, ymin, xmax, ymax)
                    If i = 1 Then
                        yref = Math.Round(ymax, 6)
                    Else
                        If Math.Round(ymax, 6) <> yref Then
                            .TrackDistance *= -1
                        End If
                    End If
                    dimstyle = .Style
                End With
                dimstyle.OrigTerminatorSize = 0.5
                dimstyle.TerminatorSize = 0.5
            Next

            If Math.Abs(ymin - yminf) > Math.Abs(trackdistance) And Math.Round(ymax, 6) <> Math.Round(ymaxf, 6) Then
                firstdim.TrackDistance *= -1
            End If

            lastdim = objdims.AddDistanceBetweenObjects(sortedcirclist.Last, xc2, yc2, 0, True, headerdimlines(1), xe2, ye2, 0, True)
            With lastdim
                .ReattachToDrawingView(DV)
                .MeasurementAxisEx = 1
                .MeasurementAxisDirection = False
                .TextScale = 0.75
                .TrackDistance = trackdistance
                If .Value < 0.025 Then
                    .TerminatorPosition = True
                    .BreakPosition = breakpos
                    .BreakDistance = 0.003
                End If
                .Range(xminl, yminl, xmaxl, ymaxl)
                If Math.Abs(ymin - yminl) > Math.Abs(trackdistance) And Math.Round(ymax, 6) <> Math.Round(ymaxl, 6) Then
                    .TrackDistance *= -1
                End If
                dimstyle = .Style
            End With
            dimstyle.OrigTerminatorSize = 0.5
            dimstyle.TerminatorSize = 0.5

            If lastdim.Value > firstdim.Value Then
                lastdim.DisplayType = 6
            Else
                firstdim.DisplayType = 6
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return sortedcirclist
    End Function

    Shared Sub HeaderStutzenGap(objsheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headerdimlines() As SolidEdgeDraft.DVLine2d,
                                circlist As List(Of SolidEdgeDraft.DVCircle2d), circuit As CircuitData)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objsheet.Dimensions
        Dim headerdimline As SolidEdgeDraft.DVLine2d
        Dim distancedim As SolidEdgeFrameworkSupport.Dimension
        Dim dimstyle As SolidEdgeFrameworkSupport.DimStyle
        Dim xm, ym, zm, xc, yc, xmin, ymin, xmax, ymax As Double
        Dim htype, ktype, cindex As Integer
        Dim circleside As String
        Dim xlist, ylist As New List(Of Double)
        Dim clist As New List(Of SolidEdgeDraft.DVCircle2d)

        Try
            If (circuit.ConnectionSide = "right" And circuit.CircuitType <> "Defrost") OrElse (circuit.ConnectionSide = "left" And circuit.CircuitType = "Defrost") Then
                'left circle
                circleside = "left"
                headerdimline = headerdimlines(0)
            Else
                'right circle
                circleside = "right"
                headerdimline = headerdimlines(1)
            End If

            'middle point of the header line
            headerdimline.GetKeyPoint(2, xm, ym, zm, ktype, htype)

            'find the correct circle
            For Each circ In circlist
                circ.GetCenterPoint(xc, yc)
                If Math.Abs(yc - ym) > 0.048 And Math.Round(circ.Diameter * 1000, 3) = circuit.CoreTube.Diameter Then
                    clist.Add(circ)
                    xlist.Add(xc)
                End If
            Next

            If circleside = "left" Then
                cindex = xlist.IndexOf(xlist.Min)
            Else
                cindex = xlist.IndexOf(xlist.Max)
            End If

            clist(cindex).GetCenterPoint(xc, yc)

            distancedim = objdims.AddDistanceBetweenObjects(headerdimline, xm, ym, zm, True, clist(cindex), xc, yc, 0, True)
            With distancedim
                .MeasurementAxisDirection = True
                .MeasurementAxisEx = 1
                If Math.Round(.Value, 3) = 0.1 Then
                    .TrackDistance = -0.015
                Else
                    .TrackDistance = -0.01
                End If
                .TextScale = 0.75
                .BreakPosition = 2
                .BreakDistance = 0.5
                .Range(xmin, ymin, xmax, ymax)
                If Math.Round(xmin, 6) >= Math.Round(xm, 6) Then
                    .TrackDistance *= -1
                End If
                .ReattachToDrawingView(DV)
                dimstyle = .Style
            End With
            dimstyle.OrigTerminatorSize = 0.5
            dimstyle.TerminatorSize = 0.5
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub HotgasDimHor(objsheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headerdimlines() As SolidEdgeDraft.DVLine2d, angle As Double, conside As String, partID As String)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objsheet.Dimensions
        Dim splinelist As List(Of SolidEdgeDraft.DVBSplineCurve2d)
        Dim headerdimline As SolidEdgeDraft.DVLine2d
        Dim xc, yc, x, y As Double

        Try
            'only for normal circuit, never for defrost!
            If conside = "left" Then
                headerdimline = headerdimlines(1)
            Else
                headerdimline = headerdimlines(0)
            End If

            If Math.Abs(angle) > 0 Then
                'splines
                splinelist = GetSplinesFromOcc(DV, New List(Of String) From {partID})
                If splinelist.Count > 0 Then
                    splinelist(0).GetCentroid(xc, yc)
                    headerdimline.GetEndPoint(x, y)

                    Dim distancedim As SolidEdgeFrameworkSupport.Dimension = objdims.AddDistanceBetweenObjects(headerdimline, x, y, 0, True, splinelist(0), xc, yc, 0, True)
                    distancedim.MeasurementAxisEx = 1
                    distancedim.MeasurementAxis = False
                    distancedim.TrackDistance = -0.01
                    distancedim.ReattachToDrawingView(DV)
                End If
            End If


        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Function GetCTLinesHor(lines As List(Of SolidEdgeDraft.DVLine2d)) As SolidEdgeDraft.DVLine2d()
        Dim xs, ys As Double
        Dim xlist, ylist As New List(Of Double)
        Dim DVlist As New List(Of StutzenDVElement)
        Dim sortedlist As List(Of StutzenDVElement)
        Dim leftline, rightline As SolidEdgeDraft.DVLine2d
        Try
            For Each l In lines
                l.GetStartPoint(xs, ys)
                DVlist.Add(New StutzenDVElement With {.DVElement = l, .Xpos = Math.Round(xs, 6), .Ypos = Math.Round(ys, 6)})
                ylist.Add(Math.Round(ys, 6))
            Next

            Dim templist = From plist In DVlist Where plist.Ypos = ylist.Min Order By plist.Xpos

            sortedlist = templist.ToList
            leftline = sortedlist(0).DVElement

            rightline = sortedlist.Last.DVElement

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return {leftline, rightline}
    End Function

    Shared Function SetHeaderDistanceHor(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headerline As SolidEdgeDraft.DVLine2d, finline As SolidEdgeDraft.DVLine2d, conside As String,
                                  onlya As Boolean, onlycenter As Boolean) As SolidEdgeFrameworkSupport.Dimension
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objSheet.Dimensions
        Dim distancedim As SolidEdgeFrameworkSupport.Dimension
        Dim xlist, ylist As New List(Of Double)
        Dim xsh, ysh, xeh, yeh, xm, ym, z, xmin, ymin, xmax, ymax, xs, ys, xe, ye, xref, yref, yrefh As Double
        Dim htype, ktype As Integer

        Try
            headerline.GetStartPoint(xsh, ysh)
            headerline.GetEndPoint(xeh, yeh)
            headerline.GetKeyPoint(2, xm, ym, z, ktype, htype)
            finline.GetStartPoint(xs, ys)
            finline.GetEndPoint(xe, ye)

            yrefh = Math.Max(ysh, yeh)

            If (xs > xe And conside = "right") Or (xs < xe And conside = "left") Then
                xref = xs
                yref = ys
            Else
                xref = xe
                yref = ye
            End If

            If Not onlycenter Then
                'minimum distance between header and tube sheet
                distancedim = objdims.AddDistanceBetweenObjects(finline, xref, yref, 0, True, headerline, xsh, yrefh, 0, True)
                distancedim.ReattachToDrawingView(DV)

                distancedim.MeasurementAxisEx = 1
                distancedim.MeasurementAxisDirection = True
                distancedim.TrackDistance = 0.0125
                distancedim.Range(xmin, ymin, xmax, ymax)
                If (Math.Round(xmin, 9) >= Math.Round(xref, 6) And conside = "left") Or (Math.Round(xmax, 6) <= Math.Round(xref, 6) And conside = "right") Then
                    distancedim.TrackDistance = -0.0125
                End If
            End If

            If Not onlya Then
                'distance between tube sheet and center of header
                distancedim = objdims.AddDistanceBetweenObjects(finline, xref, yref, 0, True, headerline, xm, ym, 0, True)
                distancedim.ReattachToDrawingView(DV)

                distancedim.MeasurementAxisEx = 1
                distancedim.MeasurementAxisDirection = True
                distancedim.TrackDistance = 0.02
                distancedim.Range(xmin, ymin, xmax, ymax)
                If (Math.Round(xmin, 9) >= Math.Round(xref, 6) And conside = "left") Or (Math.Round(xmax, 6) <= Math.Round(xref, 6) And conside = "right") Then
                    distancedim.TrackDistance = -0.02
                End If
                distancedim.DisplayType = 6
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return distancedim
    End Function

    Shared Sub MoveIsoView(dftdoc As SolidEdgeDraft.DraftDocument, finnedheight As Double)
        Dim sheetframe(), lowerdistance As Double
        Dim xref, yref, dy, xmin, ymin, xmax, ymax As Double

        Try
            If finnedheight > 500 Then
                sheetframe = {0.59, 0.415}
            Else
                sheetframe = {0.415, 0.29}
            End If

            lowerdistance = 0.07

            dftdoc.ActiveSheet.DrawingViews.Item(3).Range(xmin, ymin, xmax, ymax)
            Debug.Print(Math.Round(sheetframe(0) - xmax, 6).ToString)
            xref = xmax + (sheetframe(0) - xmax) / 2
            dftdoc.ActiveSheet.DrawingViews.Item(5).Range(xmin, ymin, xmax, ymax)
            dy = Math.Round(ymax - ymin, 6)
            yref = (sheetframe(1) - lowerdistance) / 2 + lowerdistance

            If dy / 2 > yref Then
                'rescale
            End If

            dftdoc.ActiveSheet.DrawingViews.Item(5).SetOrigin(xref, yref)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)

        End Try

    End Sub

    Shared Sub DimHotGas(dftdoc As SolidEdgeDraft.DraftDocument, consys As ConSysData, circuit As CircuitData)
        Dim hotgasDV, objdv As SolidEdgeDraft.DrawingView
        Dim x0, y0, xmin, yInmin, xmax, yInmax, yOutmin, yOutmax, xc, yc, xl, yl, scalefactor As Double
        Dim headerlines As List(Of SolidEdgeDraft.DVLine2d)
        Dim headerdimlines() As SolidEdgeDraft.DVLine2d
        Dim circles As List(Of SolidEdgeDraft.DVCircle2d)
        Dim objdims As SolidEdgeFrameworkSupport.Dimensions = dftdoc.ActiveSheet.Dimensions
        Dim leftdim, rightdim As SolidEdgeFrameworkSupport.Dimension
        Dim dimstyle As SolidEdgeFrameworkSupport.DimStyle


        Try
            scalefactor = dftdoc.ActiveSheet.DrawingViews.Item(2).ScaleFactor
            objdv = dftdoc.ActiveSheet.DrawingViews.AddAssemblyView(dftdoc.ModelLinks.Item(1), SolidEdgeDraft.ViewOrientationConstants.igRightView, scalefactor, 0, 0, 0, "HGHeader")
            objdv.SetRotationAngle(Math.PI * (90 - consys.HotGasData.Angle) / 180)

            dftdoc.ActiveSheet.DrawingViews.Item(2).Range(xmin, yOutmin, xmax, yOutmax)
            dftdoc.ActiveSheet.DrawingViews.Item(3).Range(xmin, yInmin, xmax, yInmax)
            dftdoc.ActiveSheet.DrawingViews.Item(3).GetOrigin(x0, y0)

            hotgasDV = dftdoc.ActiveSheet.DrawingViews.AddByFold(objdv, SolidEdgeDraft.FoldTypeConstants.igFoldRight, 0, 0)
            objdv.Delete()

            hotgasDV.ScaleFactor = 1
            hotgasDV.SetOrigin(0, 0)

            hotgasDV.CaptionDefinitionTextPrimary = "Hotgas Position"
            hotgasDV.CaptionLocation = 1
            hotgasDV.DisplayCaption = True

            If consys.HotGasData.Headertype = "outlet" Then
                headerlines = SEDrawing.GetLinesFromOcc(hotgasDV, "OutletHeader", circuit.ConnectionSide)
                headerdimlines = SEDrawing.GetHorHeaderDimLines(headerlines, consys.OutletHeaders.First.Tube.Diameter)
            Else
                headerlines = SEDrawing.GetLinesFromOcc(hotgasDV, "InletHeader", circuit.ConnectionSide)
                headerdimlines = SEDrawing.GetHorHeaderDimLines(headerlines, consys.InletHeaders.First.Tube.Diameter)
            End If
            circles = SEDrawing.GetCirclesFromOcc(hotgasDV, New List(Of String) From {"Header"})

            If circles.Count > 0 Then
                circles.First.GetCenterPoint(xc, yc)
                headerdimlines(0).GetStartPoint(xl, yl)

                leftdim = objdims.AddDistanceBetweenObjects(circles.First, xc, yc, 0, True, headerdimlines(0), xl, yl, 0, True)
                With leftdim
                    .MeasurementAxisEx = 1
                    .MeasurementAxisDirection = False
                    .TrackDistance = 0.15
                    SEDrawing.ControlDimDistance(leftdim, {"y", "smaller"})
                    .TrackDistance = 0.015
                    If circuit.ConnectionSide = "left" Then
                        .TrackDistance *= -1
                    End If
                    .ReattachToDrawingView(hotgasDV)
                    dimstyle = .Style
                End With
                dimstyle.OrigTerminatorSize = 0.5
                dimstyle.TerminatorSize = 0.5

                headerdimlines(1).GetStartPoint(xl, yl)
                rightdim = objdims.AddDistanceBetweenObjects(circles.First, xc, yc, 0, True, headerdimlines(1), xl, yl, 0, True)
                With rightdim
                    .MeasurementAxisEx = 1
                    .MeasurementAxisDirection = False
                    .TrackDistance = 0.15
                    SEDrawing.ControlDimDistance(rightdim, {"y", "smaller"})
                    .TrackDistance = -0.015
                    If circuit.ConnectionSide = "right" Then
                        .TrackDistance *= -1
                    End If
                    .ReattachToDrawingView(hotgasDV)
                    dimstyle = .Style
                End With
                dimstyle.OrigTerminatorSize = 0.5
                dimstyle.TerminatorSize = 0.5

                If leftdim.Value > rightdim.Value Then
                    leftdim.DisplayType = 6
                Else
                    rightdim.DisplayType = 6
                End If
            End If

            hotgasDV.ScaleFactor = scalefactor

            hotgasDV.SetOrigin(x0, y0 - (yInmax - yInmin) / 2)
            hotgasDV.Range(xmin, yOutmin, xmax, yOutmax)

            hotgasDV.SetOrigin(x0, y0 - (yOutmax - yOutmin) / 2 - 0.01)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function FindVentArc(headerarcs As List(Of SolidEdgeDraft.DVArc2d), conside As String) As SolidEdgeDraft.DVArc2d
        Dim xs, ys, xe, ye As Double
        Dim xlist As New List(Of Double)
        Dim newDVlist As New List(Of StutzenDVElement)
        Dim newarc As SolidEdgeDraft.DVArc2d = Nothing

        Try
            For Each a In headerarcs
                a.GetStartPoint(xs, ys)
                a.GetEndPoint(xe, ye)
                If Math.Round(xs, 6) = Math.Round(xe, 6) Then
                    newDVlist.Add(New StutzenDVElement With {.DVElement = a, .Xpos = Math.Round(xs, 6)})
                    xlist.Add(Math.Round(xs, 6))
                End If
            Next

            If conside = "left" Then
                Dim templist = From plist In newDVlist Where plist.Xpos = xlist.Min

                If templist.Count > 0 Then
                    newarc = templist.ToList.First.DVElement
                End If

            Else
                Dim templist = From plist In newDVlist Where plist.Xpos = xlist.Max

                If templist.Count > 0 Then
                    newarc = templist.ToList.First.DVElement
                End If
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return newarc
    End Function

    Shared Sub SwitchBreakPosition(objdim As SolidEdgeFrameworkSupport.Dimension)
        Try
            objdim.BreakPosition = 4 - objdim.BreakPosition
        Catch ex As Exception

        End Try
    End Sub

    Shared Sub CreatePartListCoil(dftdoc As SolidEdgeDraft.DraftDocument)
        Dim isoDV, helpDV As SolidEdgeDraft.DrawingView
        Dim asmlink As SolidEdgeDraft.ModelLink = dftdoc.ModelLinks.Item(1)
        Dim objPartLists As SolidEdgeDraft.PartsLists
        Dim partlist As SolidEdgeDraft.PartsList
        Dim objBalloons As SolidEdgeFrameworkSupport.Balloons
        Dim objColumn As SolidEdgeDraft.TableColumn
        Dim proptext As String
        Dim bscale As Double

        Try
            helpDV = dftdoc.ActiveSheet.DrawingViews.AddAssemblyView(asmlink, SolidEdgeDraft.ViewOrientationConstants.igFrontView, 0.05, 0, 0, 0)
            isoDV = dftdoc.ActiveSheet.DrawingViews.AddByFold(helpDV, SolidEdgeDraft.FoldTypeConstants.igFoldDownLeft, 0, 0)

            helpDV.Delete()

            objPartLists = dftdoc.PartsLists
            partlist = objPartLists.Add(isoDV, "ISO", AutoBalloon:=1, CreatePartsList:=1)
            partlist.SetComponentSortPriority(SolidEdgeDraft.PartsListComponentType.igPartsListComponentType_FrameMembers, 6)
            partlist.RenumberAccordingToSortOrder = True
            partlist.ItemNumberIncrement = 5
            partlist.ItemNumberStart = 5
            isoDV.Update()
            partlist.Update()
            partlist.AnchorPoint = SolidEdgeDraft.TableAnchorPoint.igLowerRight
            partlist.FillEndOfTableWithBlankRows = False
            partlist.ShowTopAssembly = False
            bscale = 2
            partlist.SetOrigin(0.409, 0.005)
            partlist.Update()
            partlist.ListType = SolidEdgeDraft.PartsListType.igExploded

            partlist.Update()
            objBalloons = dftdoc.ActiveSheet.Balloons

            objColumn = partlist.Columns.Item(4)
            proptext = objColumn.PropertyText

            objColumn = partlist.Columns.Item(1)
            objColumn.Header = "Quantity"
            objColumn.Width = 0.015
            objColumn.DataHorizontalAlignment = SolidEdgeFrameworkSupport.TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.PropertyText = proptext
            objColumn.HeaderRowHorizontalAlignment = SolidEdgeFrameworkSupport.TextHorizontalAlignmentConstants.igTextHzAlignCenter

            objColumn = partlist.Columns.Item(2)
            objColumn.Header = "ERP Number"
            objColumn.Width = 0.025
            objColumn.DataHorizontalAlignment = SolidEdgeFrameworkSupport.TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.PropertyText = "%{CDB_ERP_Artnr./CP|G}"
            objColumn.HeaderRowHorizontalAlignment = SolidEdgeFrameworkSupport.TextHorizontalAlignmentConstants.igTextHzAlignCenter

            objColumn = partlist.Columns.Item(3)
            objColumn.Header = "Designation"
            objColumn.Width = 0.05
            objColumn.DataHorizontalAlignment = SolidEdgeFrameworkSupport.TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.PropertyText = "%{CDB_Benennung_de/CP|G}"

            objColumn = partlist.Columns.Item(4)
            objColumn.Header = "PDM Number"
            objColumn.Width = 0.03
            objColumn.DataHorizontalAlignment = SolidEdgeFrameworkSupport.TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.PropertyText = "%{CDB_teilenummer/CP|G}"
            objColumn.HeaderRowHorizontalAlignment = SolidEdgeFrameworkSupport.TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.Show = True

            partlist.ListType = SolidEdgeDraft.PartsListType.igTopLevel

            isoDV.Delete()
            General.seapp.DoIdle()

            For Each objb As SolidEdgeFrameworkSupport.Balloon In objBalloons
                objb.Delete()
            Next

            partlist.Update()

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

End Class
