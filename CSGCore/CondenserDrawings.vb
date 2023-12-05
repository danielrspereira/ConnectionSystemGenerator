Imports System.Runtime.Remoting.Contexts
Imports SolidEdgeFrameworkSupport

Public Class CondenserDrawings
    Shared Sub MainCoil(dftdoc As SolidEdgeDraft.DraftDocument, coil As CoilData)

        'add bow views
        AddBowViews(dftdoc, coil)

        'add air direction blocks
        CreateADBows(dftdoc, "", coil.Circuits.First)

        If coil.Circuits.First.FinType = "G" Then
            'fill out the circles in fin
            FillSupportTubes(dftdoc, coil)
        End If

        'partlist
        CreatePartListCoil(dftdoc)

    End Sub

    Shared Sub MainConsys(dftdoc As SolidEdgeDraft.DraftDocument, coil As CoilData, circuit As CircuitData, consys As ConSysData)

        Try
            FrontViewDim(dftdoc, coil, circuit, consys)

            SideViewDim(dftdoc, circuit, consys)

            IsoView(dftdoc, circuit, coil)
        Catch ex As Exception

        End Try

    End Sub

    Shared Sub EmptyConsys(dftdoc As SolidEdgeDraft.DraftDocument, coil As CoilData, consys As ConSysData)

        Try
            AddEmtpyView(dftdoc, coil, consys)

            SEDrawing.WriteCostumProps(dftdoc, consys.ConSysFile)
            SEPart.GetSetCustomProp(dftdoc, "GUE_Block", "0", "write")
            SEPart.GetSetCustomProp(dftdoc, "Z_Kategorie", "2D-Baugruppenzeichnung", "write")
            SEDraft.FitWindow()
            General.seapp.Documents.CloseDocument(dftdoc.FullName, SaveChanges:=True, DoIdle:=True)
        Catch ex As Exception

        End Try
    End Sub

    Shared Sub AddEmtpyView(dftdoc As SolidEdgeDraft.DraftDocument, coil As CoilData, consys As ConSysData)
        Dim asmlink As SolidEdgeDraft.ModelLink
        Dim frontDV As SolidEdgeDraft.DrawingView
        Dim sf As Double

        Try
            If dftdoc.ModelLinks.Count = 0 Then
                asmlink = dftdoc.ModelLinks.Add(consys.ConSysFile.Fullfilename)
            Else
                asmlink = dftdoc.ModelLinks.Item(1)
            End If

            'delete A3 sheet
            dftdoc.Sheets.Item("A3").Delete()
            General.seapp.DoIdle()

            If General.currentjob.Plant = "Beji" Then
                'switch layer for background
                GACVDrawings.SwitchLayers(dftdoc, General.currentjob.Plant)
            End If

            sf = GetScaling(coil, "consys")
            frontDV = dftdoc.ActiveSheet.DrawingViews.AddAssemblyView(asmlink, 4, sf, 0, 0, 0)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub AddBowViews(dftdoc As SolidEdgeDraft.DraftDocument, coil As CoilData)
        Dim objmls As SolidEdgeDraft.ModelLinks
        Dim asmlink As SolidEdgeDraft.ModelLink
        Dim firstfrontDV, firstbackDV As SolidEdgeDraft.DrawingView
        Dim scalefactor, xmin, ymin, xmax, ymax, sheetframe(), maxDVcount, x0, y0 As Double

        Try
            dftdoc.Sheets.Item("A3").Delete()

            sheetframe = {0.59, 0.41}
            dftdoc.ActiveSheet.Name = "Coil1"

            objmls = dftdoc.ModelLinks
            asmlink = objmls.Add(coil.CoilFile.Fullfilename)

            scalefactor = GetScaling(coil, "coil")

            If coil.Frontbowids.Count = 0 Then
                'only backbows → one row
                For i As Integer = 0 To coil.Backbowids.Count - 1
                    SEDrawing.AddBowViewtoSheet(dftdoc.ActiveSheet, asmlink, scalefactor, coil.Backbowids(i), "Back", coil)
                    If coil.Alignment = "vertical" Then
                        dftdoc.ActiveSheet.DrawingViews.Item(i + 1).SetRotationAngle(Math.PI * 3 / 2 * coil.RotationDirection)
                    End If
                Next
                firstbackDV = dftdoc.ActiveSheet.DrawingViews.Item(1)
                x0 = 0.59 / 2 + 0.02

                'calculate y location depending of DVcount
                For i As Integer = 0 To coil.Backbowids.Count - 1
                    y0 = Math.Round(0.41 - (sheetframe(1) - 0.07) / (coil.Backbowids.Count + 1) * (i + 1), 6)
                    dftdoc.ActiveSheet.DrawingViews.Item(i + 1).SetOrigin(x0, y0)
                Next

            ElseIf coil.Backbowids.Count = 0 Then
                'only frontbows → one row
                For i As Integer = 0 To coil.Frontbowids.Count - 1
                    SEDrawing.AddBowViewtoSheet(dftdoc.ActiveSheet, asmlink, scalefactor, coil.Frontbowids(i), "Front", coil)
                    If coil.Alignment = "vertical" Then
                        dftdoc.ActiveSheet.DrawingViews.Item(i + 1).SetRotationAngle(Math.PI / 2 * coil.RotationDirection)
                    End If
                Next
                firstfrontDV = dftdoc.ActiveSheet.DrawingViews.Item(1)
                x0 = 0.59 / 2 + 0.02

                'calculate x location depending of DVcount
                For i As Integer = 0 To coil.Frontbowids.Count - 1
                    y0 = Math.Round(0.41 - (sheetframe(1) - 0.07) / (coil.Frontbowids.Count + 1) * (i + 1), 6)
                    dftdoc.ActiveSheet.DrawingViews.Item(i + 1).SetOrigin(x0, y0)
                Next

            Else
                firstfrontDV = SEDrawing.AddBowViewtoSheet(dftdoc.ActiveSheet, asmlink, scalefactor, coil.Frontbowids.First, "Front", coil)
                If coil.Alignment = "vertical" Then
                    firstfrontDV.SetRotationAngle(Math.PI / 2 * coil.RotationDirection)
                End If
                'check how many drawing views fit in one line
                maxDVcount = Calculation.MaxDVcount(firstfrontDV, coil.FinnedDepth)

                If maxDVcount < coil.Frontbowids.Count + coil.Backbowids.Count Then
                    'rescale until it fits
                    Do
                        firstfrontDV.Range(xmin, ymin, xmax, ymax)
                        scalefactor = Calculation.RescaleFactor(scalefactor, ymax, "down", 0)
                        firstfrontDV.ScaleFactor = scalefactor
                        maxDVcount = Calculation.MaxDVcount(firstfrontDV, coil.FinnedDepth)
                    Loop Until maxDVcount >= coil.Frontbowids.Count + coil.Backbowids.Count
                End If
                'place them all in one line
                Dim k As Integer = 1
                Do
                    If coil.Frontbowids.Count > 1 Then
                        For i As Integer = 1 To coil.Frontbowids.Count - 1
                            SEDrawing.AddBowViewtoSheet(dftdoc.ActiveSheet, asmlink, scalefactor, coil.Frontbowids(i), "Front", coil)
                            If coil.Alignment = "vertical" Then
                                dftdoc.ActiveSheet.DrawingViews.Item(i + 1).SetRotationAngle(Math.PI / 2 * coil.RotationDirection)
                            End If
                            k += 1
                        Next
                    End If
                    For i As Integer = 0 To coil.Backbowids.Count - 1
                        SEDrawing.AddBowViewtoSheet(dftdoc.ActiveSheet, asmlink, scalefactor, coil.Backbowids(i), "Back", coil)
                        If coil.Alignment = "vertical" Then
                            dftdoc.ActiveSheet.DrawingViews.Item(coil.Frontbowids.Count + i + 1).SetRotationAngle(Math.PI * 3 / 2 * coil.RotationDirection)
                        End If
                        k += 1
                    Next
                Loop Until k >= coil.Frontbowids.Count + coil.Backbowids.Count

                x0 = 0.59 / 2 + 0.02

                'calculate x location depending of DVcount
                For i As Integer = 0 To (coil.Frontbowids.Count + coil.Backbowids.Count) - 1
                    y0 = Math.Round(0.41 - (sheetframe(1) - 0.07) / (coil.Frontbowids.Count + coil.Backbowids.Count + 1) * (i + 1), 6)
                    dftdoc.ActiveSheet.DrawingViews.Item(i + 1).SetOrigin(x0, y0)
                Next
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function GetScaling(coil As CoilData, dfttype As String) As Double
        Dim scalefactor As Double

        If dfttype = "consys" Then
            If coil.FinnedHeight < 1000 Then
                scalefactor = 0.4
            ElseIf coil.FinnedHeight > 2200 Then
                scalefactor = 0.15
            ElseIf coil.FinnedHeight > 2000 Then
                scalefactor = 0.2
            ElseIf coil.FinnedHeight < 1300 Then
                scalefactor = 0.3
            Else
                scalefactor = 0.25
            End If
        Else
            If coil.FinnedHeight < 1300 Then
                scalefactor = 0.4
            ElseIf coil.FinnedHeight > 2200 Then
                scalefactor = 0.15
            ElseIf coil.FinnedHeight > 2000 Then
                scalefactor = 0.2
            Else
                scalefactor = 0.333
            End If
        End If

        Return scalefactor
    End Function

    Shared Sub CreateADBows(dftdoc As SolidEdgeDraft.DraftDocument, position As String, circuit As CircuitData)
        Dim objSheet As SolidEdgeDraft.Sheet = dftdoc.ActiveSheet

        Try

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

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub CreatePartListCoil(dftdoc As SolidEdgeDraft.DraftDocument)
        Dim isoDV, helpDV As SolidEdgeDraft.DrawingView
        Dim asmlink As SolidEdgeDraft.ModelLink = dftdoc.ModelLinks.Item(1)
        Dim objPartLists As SolidEdgeDraft.PartsLists
        Dim partlist As SolidEdgeDraft.PartsList
        Dim objBalloons As Balloons
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
            'objColumn = partlist.Columns.Item(1)
            'objColumn.Header = "Pos."
            'objColumn.Width = 0.015
            'objColumn.DataHorizontalAlignment = TextHorizontalAlignmentConstants.igTextHzAlignCenter
            'objColumn.PropertyText = "%{Item Number|G}"
            'objColumn.HeaderRowHorizontalAlignment = TextHorizontalAlignmentConstants.igTextHzAlignCenter

            objColumn = partlist.Columns.Item(1)
            objColumn.Header = "Quantity"
            objColumn.Width = 0.015
            objColumn.DataHorizontalAlignment = TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.PropertyText = proptext
            objColumn.HeaderRowHorizontalAlignment = TextHorizontalAlignmentConstants.igTextHzAlignCenter

            objColumn = partlist.Columns.Item(2)
            objColumn.Header = "ERP Number"
            objColumn.Width = 0.025
            objColumn.DataHorizontalAlignment = TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.PropertyText = "%{CDB_ERP_Artnr./CP|G}"
            objColumn.HeaderRowHorizontalAlignment = TextHorizontalAlignmentConstants.igTextHzAlignCenter

            objColumn = partlist.Columns.Item(3)
            objColumn.Header = "Designation"
            objColumn.Width = 0.05
            objColumn.DataHorizontalAlignment = TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.PropertyText = "%{CDB_Benennung_de/CP|G}"

            'objColumn = partlist.Columns.Add(5, True)
            objColumn = partlist.Columns.Item(4)
            objColumn.Header = "PDM Number"
            objColumn.Width = 0.03
            objColumn.DataHorizontalAlignment = TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.PropertyText = "%{CDB_teilenummer/CP|G}"
            objColumn.HeaderRowHorizontalAlignment = TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.Show = True

            partlist.ListType = SolidEdgeDraft.PartsListType.igTopLevel

            isoDV.Delete()
            General.seapp.DoIdle()

            For Each objb As Balloon In objBalloons
                objb.Delete()
            Next

            partlist.Update()

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub DVLocation(objDV As SolidEdgeDraft.DrawingView)
        Dim xmin, ymin, xmax, ymax As Double
        objDV.Range(xmin, ymin, xmax, ymax)
        objDV.SetOrigin(Math.Abs(xmin - 0.032), 0.32)
    End Sub

    Shared Sub FrontViewDim(dftdoc As SolidEdgeDraft.DraftDocument, coil As CoilData, circuit As CircuitData, consys As ConSysData)
        Dim asmlink As SolidEdgeDraft.ModelLink
        Dim frontDV, backDV As SolidEdgeDraft.DrawingView
        Dim oritentationconst, row As Integer
        Dim headerframe() As Double
        Dim inletABV, outletABV As New List(Of Double)
        Dim inletIDs, outletIDs As New List(Of String)
        Dim headerlines, nipplelines As List(Of SolidEdgeDraft.DVLine2d)
        Dim headerdimlines() As SolidEdgeDraft.DVLine2d
        Dim nipplecirclist As List(Of SolidEdgeDraft.DVCircle2d)
        Dim inheaderdim, outheaderdim, outnippledimension As Dimension
        Dim x, y, sf As Double

        Try
            If dftdoc.ModelLinks.Count = 0 Then
                asmlink = dftdoc.ModelLinks.Add(consys.ConSysFile.Fullfilename)
            Else
                asmlink = dftdoc.ModelLinks.Item(1)
            End If

            'delete A3 sheet
            dftdoc.Sheets.Item("A3").Delete()
            General.seapp.DoIdle()

            If General.currentjob.Plant = "Beji" Then
                'switch layer for background
                GACVDrawings.SwitchLayers(dftdoc, General.currentjob.Plant)
            End If

#Region "Inlet"
            'inlet view
            sf = GetScaling(coil, "consys")
            frontDV = dftdoc.ActiveSheet.DrawingViews.AddAssemblyView(asmlink, 4, sf, 0, 0, 0, ConfigurationName:="Inlet")
            If coil.Alignment = "vertical" Then
                frontDV.SetRotationAngle(Math.PI / 2 * coil.RotationDirection)
            End If
            frontDV.CaptionDefinitionTextPrimary = "Inlet"
            frontDV.CaptionLocation = DimViewCaptionLocationConstants.igDimViewCaptionLocationTop
            frontDV.DisplayCaption = False
            DVLocation(frontDV)
            frontDV.GetOrigin(x, y)

            'set dimensions for each stutzen, getting higher track distance with each row, lowest row must be > upper boundary
            For Each s In consys.InletHeaders(0).StutzenDatalist
                inletIDs.Add(s.ID)
                inletABV.Add(s.ABV)
            Next

            'get vertical headerlines to create a frame
            headerlines = SEDrawing.GetLinesFromOcc(frontDV, consys.InletHeaders.First.Tube.TubeFile.Shortname, circuit.ConnectionSide)

            'left / right vertical line
            headerdimlines = SEDrawing.GetHorHeaderDimLines(headerlines, consys.InletHeaders.First.Tube.Diameter)

            headerframe = SEDrawing.GetHeaderFrame(headerdimlines)

            'set dim for header length, nipple position And vent position
            If consys.FlangeID IsNot Nothing And consys.FlangeID <> "" Then
                nipplecirclist = SEDrawing.GetCirclesFromOcc(frontDV, New List(Of String) From {consys.FlangeID})
            Else
                nipplecirclist = SEDrawing.GetCirclesFromOcc(frontDV, New List(Of String) From {"InletNipple"})
            End If
            Dim innippledimension As Dimension = NippleDimFD(dftdoc.ActiveSheet, frontDV, headerdimlines(0), nipplecirclist, "inlet", sf)
            If consys.InletHeaders.First.VentIDs IsNot Nothing Then
                'vent position
                'must create new line2d for dim, other methods have failed
                HorVentDim(dftdoc.ActiveSheet, frontDV, headerdimlines, "inlet", innippledimension, coil.Alignment)
            End If
            inheaderdim = HorHeaderDim(dftdoc.ActiveSheet, frontDV, headerdimlines, "inlet", innippledimension)

            If consys.InletHeaders.First.Tube.SVPosition(1) = "perp" Then
                SVDim(dftdoc.ActiveSheet, frontDV, headerdimlines(0), innippledimension)

                'increase trackdistance for innippledim and inheaderdim
                If innippledimension.TrackDistance < 0 Then
                    innippledimension.TrackDistance -= 0.01
                Else
                    innippledimension.TrackDistance += 0.01
                End If
                If inheaderdim.TrackDistance < 0 Then
                    inheaderdim.TrackDistance -= 0.01
                Else
                    inheaderdim.TrackDistance += 0.01
                End If
            End If

            'use circles only 
            For Each uid In General.GetUniqueStrings(inletIDs)
                'double check if they are all in one row
                row += 1
                Dim circlist As List(Of SolidEdgeDraft.DVCircle2d) = SEDrawing.GetCirclesFromOcc(frontDV, New List(Of String) From {uid})
                If General.GetUniqueStrings(inletIDs).Count < General.GetUniqueValues(inletABV).Count AndAlso inletABV.Contains(-inletABV(row - 1)) Then
                    If inletABV.IndexOf(inletABV(row - 1)) < inletABV.IndexOf(-inletABV(row - 1)) Then
                        GetPartiallist(circlist, {"y", "bigger"})
                    Else
                        GetPartiallist(circlist, {"y", "smaller"})
                    End If
                End If
                HorStutzenPos(dftdoc.ActiveSheet, frontDV, headerdimlines, headerframe, circuit, circlist, "inlet", row, sf)
            Next

            frontDV.DisplayCaption = True
            'use last dim range for relocation of the caption
            RelocateCaption(frontDV, dftdoc.ActiveSheet, "inlet")
#End Region

            If circuit.NoPasses Mod 2 = 0 Then
                oritentationconst = 4
            Else
                oritentationconst = 6
            End If

#Region "Outlet"
            'outlet view
            backDV = dftdoc.ActiveSheet.DrawingViews.AddAssemblyView(asmlink, oritentationconst, 1, 0, 0, 0, ConfigurationName:="Outlet")
            If coil.Alignment = "vertical" Then
                backDV.SetRotationAngle(Math.PI / 2 * coil.RotationDirection)
            End If
            backDV.CaptionDefinitionTextPrimary = "Outlet"
            backDV.DisplayCaption = False
            backDV.ScaleFactor = frontDV.ScaleFactor
            backDV.SetOrigin(x, 0.2)

            'set dimensions for each stutzen, getting higher track distance with each row, lowest row must be > upper boundary
            For Each s In consys.OutletHeaders(0).StutzenDatalist
                outletIDs.Add(s.ID)
                outletABV.Add(s.ABV)
            Next

            'get vertical headerlines to create a frame
            headerlines = SEDrawing.GetLinesFromOcc(backDV, consys.OutletHeaders.First.Tube.TubeFile.Shortname, circuit.ConnectionSide)

            'left / right vertical line
            headerdimlines = SEDrawing.GetHorHeaderDimLines(headerlines, consys.OutletHeaders.First.Tube.Diameter)

            headerframe = SEDrawing.GetHeaderFrame(headerdimlines)

            'set dim for header length, nipple position and vent position
            If consys.FlangeID IsNot Nothing And consys.FlangeID <> "" Then
                nipplecirclist = SEDrawing.GetCirclesFromOcc(backDV, New List(Of String) From {consys.FlangeID})
            Else
                nipplecirclist = SEDrawing.GetCirclesFromOcc(backDV, New List(Of String) From {"OutletNipple"})
            End If

            If nipplecirclist.Count = 0 Then
                'method for nippledim, when only line availbale 
                nipplelines = SEDrawing.GetLinesFromOcc(backDV, "OutletNipple", "")
                outnippledimension = NippleDimXD(dftdoc.ActiveSheet, backDV, headerdimlines(1), nipplelines, "outlet", consys.OutletNipples.First.Diameter)
            Else
                outnippledimension = NippleDimFD(dftdoc.ActiveSheet, backDV, headerdimlines(1), nipplecirclist, "outlet", sf)
                If consys.OutletHeaders.First.VentIDs IsNot Nothing Then
                    'vent position
                    HorVentDim(dftdoc.ActiveSheet, backDV, headerdimlines, "outlet", outnippledimension, coil.Alignment)
                End If
            End If

            outheaderdim = HorHeaderDim(dftdoc.ActiveSheet, backDV, headerdimlines, "outlet", outnippledimension)

            If nipplecirclist.Count = 0 Then
                'change track distance of outheaderdim
                RelocateHeaderDim(outheaderdim, backDV)
            End If

            'use circles only 
            row = 0
            For Each uid In General.GetUniqueStrings(outletIDs)
                'double check if they are all in one row
                row += 1
                Dim circlist As List(Of SolidEdgeDraft.DVCircle2d) = SEDrawing.GetCirclesFromOcc(backDV, New List(Of String) From {uid})
                If General.GetUniqueStrings(outletIDs).Count < General.GetUniqueValues(outletABV).Count AndAlso outletABV.Contains(-outletABV(row - 1)) Then
                    If outletABV.IndexOf(outletABV(row - 1)) < outletABV.IndexOf(-outletABV(row - 1)) Then
                        GetPartiallist(circlist, {"y", "smaller"})
                    Else
                        GetPartiallist(circlist, {"y", "bigger"})
                    End If
                End If
                HorStutzenPos(dftdoc.ActiveSheet, backDV, headerdimlines, headerframe, circuit, circlist, "outlet", row, sf)
            Next

            backDV.DisplayCaption = True
            'use last dim range for relocation of the caption
            RelocateCaption(backDV, dftdoc.ActiveSheet, "outlet")
#End Region
            RelocateFrontViews(frontDV, backDV, sf, inheaderdim, outheaderdim)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Function NippleDimFD(objsheet As SolidEdgeDraft.Sheet, objDV As SolidEdgeDraft.DrawingView, headerline As SolidEdgeDraft.DVLine2d, nipplecircs As List(Of SolidEdgeDraft.DVCircle2d),
                             headertype As String, scalefactor As Double) As Dimension
        Dim objdims As Dimensions = objsheet.Dimensions
        Dim firstdim, nextdim As Dimension
        Dim dimstyle As DimStyle
        Dim xe, ye, xs, ys, xc, yc, xmin, ymin, xmax, ymax, trackdistance, delta, xmin2, ymin2, xmax2, ymax2, x0, y0 As Double
        Dim firstcirc As SolidEdgeDraft.DVCircle2d
        Dim circDVlist As New List(Of StutzenDVElement)

        Try
            objDV.GetOrigin(x0, y0)
            objDV.Range(xmin2, ymin2, xmax2, ymax2)
            headerline.GetStartPoint(xs, ys)
            headerline.GetEndPoint(xe, ye)

            circDVlist = SortNippleCircs(nipplecircs)
            firstcirc = circDVlist.First.DVElement
            firstcirc.GetCenterPoint(xc, yc)

            firstdim = objdims.AddDistanceBetweenObjects(headerline.Reference, xs, ys, 0, True, nipplecircs.First.Reference, xc, yc, 0, True)
            With firstdim
                .BreakDistance = 0.5
                .BreakPosition = 2
                .TrackDistance = 0.03
                .MeasurementAxisDirection = False
                .MeasurementAxisEx = 1
                dimstyle = .Style
                dimstyle.PrimaryDecimalRoundOff = DimDecimalRoundOffTypeConstants.igDimStyleDecimal_1
                If headertype = "inlet" Then
                    SEDrawing.ControlDimDistance(firstdim, {"y", "smaller"})
                Else
                    SEDrawing.ControlDimDistance(firstdim, {"y", "bigger"})
                End If

                'use boundaries to get the correct track distance
                trackdistance = .TrackDistance
                .Range(xmin, ymin, xmax, ymax)

                If headertype = "inlet" Then
                    delta = ymin - ymin2 + 0.005
                    If delta > 0 Then
                        If trackdistance > 0 Then
                            trackdistance += delta
                        Else
                            trackdistance -= delta
                        End If
                    End If
                Else
                    delta = ymax2 - ymax + 0.0075
                    If delta > 0 Then
                        If trackdistance > 0 Then
                            trackdistance += delta
                        Else
                            trackdistance -= delta
                        End If
                    End If
                End If
                .TrackDistance = trackdistance
            End With

            If circDVlist.Count > 1 Then
                'dimension from circ to circ
                For i As Integer = 1 To circDVlist.Count - 1
                    Dim nextcirc As SolidEdgeDraft.DVCircle2d = circDVlist(i).DVElement
                    Dim xc2, yc2 As Double
                    nextcirc.GetCenterPoint(xc2, yc2)
                    firstcirc = circDVlist(i - 1).DVElement
                    firstcirc.GetCenterPoint(xc, yc)
                    nextdim = objdims.AddDistanceBetweenObjects(firstcirc.Reference, xc, yc, 0, True, nextcirc.Reference, xc2, yc2, 0, True)

                    'since dimension has different first ref object, the trackdistance is completely different for the dim to align with firstdim 
                    With nextdim
                        .TrackDistance = trackdistance
                        .MeasurementAxisDirection = False
                        .MeasurementAxisEx = 1
                        .BreakDistance = 0.5
                        .BreakPosition = 2
                        If headertype = "inlet" Then
                            SEDrawing.ControlDimDistance(nextdim, {"y", "smaller"})
                        Else
                            SEDrawing.ControlDimDistance(nextdim, {"y", "bigger"})
                        End If
                        trackdistance = .TrackDistance
                        firstdim.Range(xmin, ymin, xmax, ymax)
                        .Range(xmin2, ymin2, xmax2, ymax2)

                        If headertype = "inlet" Then
                            delta = ymin2 - ymin
                            If trackdistance > 0 Then
                                trackdistance += delta
                            Else
                                trackdistance -= delta
                            End If
                        Else
                            delta = ymax2 - ymax
                            If trackdistance > 0 Then
                                trackdistance -= delta
                            Else
                                trackdistance += delta
                            End If
                        End If
                        .TrackDistance = trackdistance
                    End With
                Next

            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return firstdim
    End Function

    Shared Function NippleDimXD(objsheet As SolidEdgeDraft.Sheet, objDV As SolidEdgeDraft.DrawingView, headerline As SolidEdgeDraft.DVLine2d, nipplelines As List(Of SolidEdgeDraft.DVLine2d),
                              headertype As String, diameter As Double) As Dimension
        Dim objdims As Dimensions = objsheet.Dimensions
        Dim firstdim, nextdim As Dimension
        Dim nippleline As SolidEdgeDraft.DVLine2d
        Dim dimstyle As DimStyle
        Dim xe, ye, xs, ys, xc, yc, z, xmin, ymin, xmax, ymax, trackdistance, delta, x0, y0, xmin2, ymin2, xmax2, ymax2, scalefactor As Double
        Dim htype, ktype As Integer
        Dim nippleDVlist As New List(Of StutzenDVElement)

        Try
            objDV.GetOrigin(x0, y0)
            scalefactor = objDV.ScaleFactor

            objDV.ScaleFactor = 1
            objDV.SetOrigin(0, 0)

            headerline.GetStartPoint(xs, ys)
            headerline.GetEndPoint(xe, ye)

            nippleDVlist = SortNippleLines(nipplelines, diameter)
            nippleline = nippleDVlist.First.DVElement
            nippleline.GetKeyPoint(2, xc, yc, z, ktype, htype)

            firstdim = objdims.AddDistanceBetweenObjects(nippleline, xc, yc, 0, True, headerline, xs, ys, 0, True)
            With firstdim
                .TrackDistance = 0.01
                .MeasurementAxisDirection = False
                .MeasurementAxisEx = 1
                .ReattachToDrawingView(objDV)
                If headertype = "inlet" Then
                    SEDrawing.ControlDimDistance(firstdim, {"y", "bigger"})
                Else
                    SEDrawing.ControlDimDistance(firstdim, {"y", "smaller"})
                End If
                dimstyle = .Style
                dimstyle.PrimaryDecimalRoundOff = DimDecimalRoundOffTypeConstants.igDimStyleDecimal1
                If .Value < 0.05 Then
                    .TerminatorPosition = True
                    .BreakPosition = 3
                    .BreakDistance = 0.005
                End If
            End With

            If nippleDVlist.Count > 1 Then
                'dimension from circ to circ
                For i As Integer = 1 To nippleDVlist.Count - 1
                    objDV.GetOrigin(x0, y0)
                    scalefactor = objDV.ScaleFactor
                    Dim nextline As SolidEdgeDraft.DVLine2d = nippleDVlist(i).DVElement
                    Dim xc2, yc2 As Double
                    nextline.GetKeyPoint(2, xc2, yc2, z, ktype, htype)
                    nippleline = nippleDVlist(i - 1).DVElement
                    nippleline.GetKeyPoint(2, xc, yc, z, ktype, htype)
                    nextdim = objdims.AddDistanceBetweenObjects(nippleline, xc, yc, 0, True, nextline, xc2, yc2, 0, True)

                    'since dimension has different first ref object, the trackdistance is completely different for the dim to align with firstdim 
                    With nextdim
                        .ReattachToDrawingView(objDV)
                        .TrackDistance = 0.01
                        .MeasurementAxisDirection = False
                        .MeasurementAxisEx = 1
                        If headertype = "inlet" Then
                            SEDrawing.ControlDimDistance(nextdim, {"y", "smaller"})
                        Else
                            SEDrawing.ControlDimDistance(nextdim, {"y", "bigger"})
                        End If
                        'use boundaries to get the correct track distance
                        objDV.ScaleFactor = scalefactor
                        objDV.SetOrigin(0, 0)
                        trackdistance = .TrackDistance
                        firstdim.Range(xmin, ymin, xmax, ymax)
                        .Range(xmin2, ymin2, xmax2, ymax2)

                        If headertype = "inlet" Then
                            delta = ymin2 - ymin
                            If trackdistance > 0 Then
                                trackdistance += delta
                            Else
                                trackdistance -= delta
                            End If
                        Else
                            delta = ymax2 - ymax
                            If trackdistance > 0 Then
                                trackdistance -= delta
                            Else
                                trackdistance += delta
                            End If
                        End If
                        .TrackDistance = trackdistance
                    End With
                Next

            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        Finally
            objDV.SetOrigin(x0, y0)
            objDV.ScaleFactor = scalefactor
        End Try

        Return firstdim
    End Function

    Shared Function SortNippleLines(linelist As List(Of SolidEdgeDraft.DVLine2d), length As Double) As List(Of StutzenDVElement)
        Dim x, y, z, xs, ys, xe, ye As Double
        Dim ktype, htype As Integer
        Dim dvelementlist As New List(Of StutzenDVElement)
        Dim xlist As New List(Of Double)

        For Each l In linelist
            l.GetStartPoint(xs, ys)
            l.GetEndPoint(xe, ye)
            If Math.Round(ys, 6) = Math.Round(ye, 6) And Math.Round(length / 1000, 6) = Math.Round(l.Length, 6) Then
                l.GetKeyPoint(2, x, y, z, ktype, htype)
                dvelementlist.Add(New StutzenDVElement With {.DVElement = l, .Xpos = Math.Round(x, 6)})
            End If
        Next

        If dvelementlist.Count > 1 Then
            Dim templist = From plist In dvelementlist Order By plist.Xpos

            Return templist.ToList
        Else
            Return dvelementlist
        End If
    End Function

    Shared Function SortNippleCircs(circlist As List(Of SolidEdgeDraft.DVCircle2d)) As List(Of StutzenDVElement)
        Dim diameterlist As New List(Of Double)
        Dim dvelementlist As New List(Of StutzenDVElement)
        Dim xc, yc As Double

        'find biggest circles
        For Each c In circlist
            diameterlist.Add(Math.Round(c.Diameter, 6))
        Next
        For Each c In circlist
            If Math.Round(c.Diameter, 6) = diameterlist.Max Then
                c.GetCenterPoint(xc, yc)
                dvelementlist.Add(New StutzenDVElement With {.DVElement = c, .Xpos = Math.Round(xc, 6)})
            End If
        Next

        If dvelementlist.Count > 1 Then
            Dim templist = From plist In dvelementlist Order By plist.Xpos

            Return templist.ToList
        Else
            Return dvelementlist
        End If
    End Function

    Shared Sub GetPartiallist(ByRef objlist As List(Of SolidEdgeDraft.DVCircle2d), comparer As String())
        Dim smallerlist, biggerlist As New List(Of SolidEdgeDraft.DVCircle2d)
        Dim xlist, ylist As New List(Of Double)
        Dim x, y As Double

        Try

            If objlist.Count > 0 Then
                For Each c In objlist
                    c.GetCenterPoint(x, y)
                    xlist.Add(Math.Round(x, 6))
                    ylist.Add(Math.Round(y, 6))
                Next

                If comparer(0) = "y" Then
                    If Math.Abs(ylist.Max - ylist.Min) * 1000 > 1 Then
                        For Each c In objlist
                            c.GetCenterPoint(x, y)
                            If Math.Abs(ylist.Max - Math.Round(y, 6)) > 1 Then
                                smallerlist.Add(c)
                            Else
                                biggerlist.Add(c)
                            End If
                        Next
                    Else
                        smallerlist = objlist
                        biggerlist = objlist
                    End If
                Else
                    If Math.Abs(xlist.Max - xlist.Min) * 1000 > 1 Then
                        For Each c In objlist
                            c.GetCenterPoint(x, y)
                            If Math.Abs(xlist.Max - Math.Round(y, 6)) > 1 Then
                                smallerlist.Add(c)
                            Else
                                biggerlist.Add(c)
                            End If
                        Next
                    Else
                        smallerlist = objlist
                        biggerlist = objlist
                    End If
                End If

            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        If comparer(1) = "bigger" Then
            objlist = biggerlist
        Else
            objlist = smallerlist
        End If
    End Sub

    Shared Sub HorVentDim(objsheet As SolidEdgeDraft.Sheet, objDV As SolidEdgeDraft.DrawingView, headerlines() As SolidEdgeDraft.DVLine2d, headertype As String, refdim As Dimension, alignment As String)
        Dim objdims As Dimensions = objsheet.Dimensions
        Dim headerline As SolidEdgeDraft.DVLine2d
        Dim distancedim As Dimension
        Dim helpline As Line2d
        Dim dvelementlist, sortedlist As New List(Of StutzenDVElement)
        Dim xs, ys, xe, ye, xm, ym, xmin, ymin, xmax, ymax, xmin2, ymin2, xmax2, ymax2, trackdistance, delta, x0, y0, scalefactor As Double
        Dim ylist As New List(Of Double)
        Dim breakpos As Integer

        Try
            objDV.GetOrigin(x0, y0)
            scalefactor = objDV.ScaleFactor

            For Each l In SEDrawing.GetLinesFromOcc(objDV, "letHeader", "")
                l.GetStartPoint(xs, ys)
                l.GetEndPoint(xe, ye)
                If alignment = "horizontal" Then
                    If Math.Round(ys, 6) = Math.Round(ye, 6) AndAlso ((Math.Max(xs, xe) > 0 And headertype = "inlet") OrElse (Math.Min(xs, xe) < 0 And headertype = "outlet")) Then
                        ylist.Add(Math.Round(ys, 6))
                        dvelementlist.Add(New StutzenDVElement With {.DVElement = l, .Ypos = Math.Round(ys, 6), .Xpos = l.Length})
                    End If
                Else
                    If Math.Round(ys, 6) = Math.Round(ye, 6) AndAlso ((Math.Min(xs, xe) < 0 And headertype = "inlet") OrElse (Math.Max(xs, xe) > 0 And headertype = "outlet")) Then
                        ylist.Add(Math.Round(ys, 6))
                        dvelementlist.Add(New StutzenDVElement With {.DVElement = l, .Ypos = Math.Round(ys, 6), .Xpos = l.Length})
                    End If
                End If
            Next

            'dvelementlist should contain 6 lines, only the 2 highest (inlet) or lowest (outlet) are relevant
            If headertype = "inlet" Then
                Dim templist = From plist In dvelementlist Where plist.Ypos = ylist.Max Order By plist.Xpos

                sortedlist = templist.ToList
                If alignment = "horizontal" Then
                    headerline = headerlines(1)
                Else
                    headerline = headerlines(0)
                End If
            Else
                Dim templist = From plist In dvelementlist Where plist.Ypos = ylist.Min Order By plist.Xpos

                sortedlist = templist.ToList
                If alignment = "horizontal" Then
                    headerline = headerlines(0)
                Else
                    headerline = headerlines(1)
                End If
            End If

            xm = GetMiddlePoint(sortedlist, headertype, alignment)

            'create the line2d
            helpline = DrawDimLine2(objDV.Sheet, objsheet, xm, sortedlist.First.Ypos)

            If alignment = "horizontal" Then
                breakpos = 1
            Else
                breakpos = 3
            End If

            helpline.GetStartPoint(xm, ym)
            headerline.GetStartPoint(xs, ys)
            headerline.GetEndPoint(xe, ye)

            objDV.ScaleFactor = 1
            objDV.SetOrigin(0, 0)

            distancedim = objdims.AddDistanceBetweenObjects(headerline, xs, ys, 0, True, helpline, xm, ym, 0, True)

            With distancedim
                .ReattachToDrawingView(objDV)
                .TrackDistance = 0.01
                .MeasurementAxisDirection = False
                .MeasurementAxisEx = 1
                .TerminatorPosition = True
                .BreakPosition = 1
                .BreakDistance = 0.0055
                .Style.PrimaryDecimalRoundOff = DimDecimalRoundOffTypeConstants.igDimStyleDecimal_1
                .Range(xmin, ymin, xmax, ymax)
                If headertype = "inlet" Then
                    SEDrawing.ControlDimDistance(distancedim, {"y", "smaller"})
                Else
                    SEDrawing.ControlDimDistance(distancedim, {"y", "bigger"})
                    .Range(xmin2, ymin2, xmax2, ymax2)
                    If Math.Abs(ymax2 - ymax) < 0.001 Then
                        .TrackDistance *= -1
                    End If
                End If

                'use boundaries to get the correct track distance
                objDV.ScaleFactor = scalefactor
                objDV.SetOrigin(0, 0)
                trackdistance = .TrackDistance
                .Range(xmin, ymin, xmax, ymax)

                If headertype = "inlet" Then
                    refdim.Range(xmin2, ymin2, xmax2, ymax2)

                    delta = ymin2 - ymin
                    .TrackDistance -= delta
                    .Range(xmin, ymin, xmax, ymax)

                    If Math.Abs(ymin2 - ymin) > delta Then
                        .TrackDistance = trackdistance + delta
                    End If
                Else
                    refdim.Range(xmin2, ymin2, xmax2, ymax2)

                    delta = ymax2 - ymax
                    .TrackDistance -= delta
                    .Range(xmin, ymin, xmax, ymax)

                    If Math.Abs(ymax2 - ymax) > delta Then
                        .TrackDistance = trackdistance - delta
                    End If
                End If
            End With

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        Finally
            objDV.ScaleFactor = scalefactor
            objDV.SetOrigin(x0, y0)
        End Try

    End Sub

    Shared Function HorHeaderDim(objSheet As SolidEdgeDraft.Sheet, objDV As SolidEdgeDraft.DrawingView, headerlines() As SolidEdgeDraft.DVLine2d,
                                  headertype As String, nippledim As Dimension) As Dimension
        Dim objdims As Dimensions = objSheet.Dimensions
        Dim diadim, lengthdim As Dimension
        Dim leftline As SolidEdgeDraft.DVLine2d = headerlines(0)
        Dim rightline As SolidEdgeDraft.DVLine2d = headerlines(1)
        Dim xe, ye, xs, ys, xmin, ymin, xmax, ymax, trackdistance As Double

        Try

            leftline.GetStartPoint(xs, ys)
            leftline.GetEndPoint(xe, ye)
            diadim = objdims.AddLength(leftline.Reference)

            With diadim
                .PrefixString = "%DI"
                .BreakDistance = 0.003
                .TerminatorPosition = True
                .TrackDistance = 0.01
                .Range(xmin, ymin, xmax, ymax)
                SEDrawing.ControlDimDistance(diadim, {"x", "smaller"})
            End With

            Dim x2s, x2e, y2s, y2e As Double
            rightline.GetStartPoint(x2s, y2s)
            rightline.GetEndPoint(x2e, y2e)

            lengthdim = objdims.AddDistanceBetweenObjects(leftline.Reference, xe, ye, 0, True, rightline.Reference, x2e, y2e, 0, True)

            With lengthdim
                .BreakDistance = 0.5
                .BreakPosition = 2
                .TrackDistance = 0.03
                .MeasurementAxisDirection = False
                .MeasurementAxisEx = 1

                trackdistance = nippledim.TrackDistance
                If headertype = "inlet" Then
                    If trackdistance > 0 Then
                        trackdistance += 0.01
                    Else
                        trackdistance -= 0.01
                    End If
                Else
                    If trackdistance > 0 Then
                        trackdistance += 0.01
                    Else
                        trackdistance -= 0.01
                    End If
                End If
                .TrackDistance = trackdistance
                If headertype = "inlet" Then
                    SEDrawing.ControlDimDistance(lengthdim, {"y", "smaller"})
                    SEDrawing.ControlDimDistance(diadim, {"y", "smaller"}, breakpos:=3)
                Else
                    SEDrawing.ControlDimDistance(lengthdim, {"y", "bigger"})
                    SEDrawing.ControlDimDistance(diadim, {"y", "smaller"}, breakpos:=1)
                End If
            End With
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return lengthdim
    End Function

    Shared Sub SVDim(objSheet As SolidEdgeDraft.Sheet, objDV As SolidEdgeDraft.DrawingView, headerline As SolidEdgeDraft.DVLine2d, nippledim As Dimension)
        Dim xs, ys, xe, ye, xc, yc, scalef, x0, y0 As Double
        Dim arcDVList, sortedlist As New List(Of StutzenDVElement)
        Dim svarc As SolidEdgeDraft.DVArc2d
        Dim objdims As Dimensions = objSheet.Dimensions
        Dim distancedim As Dimension

        Try
            headerline.GetStartPoint(xs, ys)
            headerline.GetEndPoint(xe, ye)
            scalef = objDV.ScaleFactor
            objDV.GetOrigin(x0, y0)

            For Each a As SolidEdgeDraft.DVArc2d In objDV.DVArcs2d
                If a.ModelMember.FileName.Contains("InletHeader") Then
                    If a.Radius * 1000 > 7 Then
                        a.GetCenterPoint(xc, yc)
                        arcDVList.Add(New StutzenDVElement With {.DVElement = a, .Xpos = Math.Round(xc, 6), .Ypos = Math.Round(yc, 6)})
                    End If
                End If
            Next

            If arcDVList.Count > 0 Then
                Dim templist = From plist In arcDVList Order By plist.Ypos Descending

                sortedlist = templist.ToList
                svarc = sortedlist.First.DVElement
                svarc.GetCenterPoint(xc, yc)

                objDV.ScaleFactor = 1
                objDV.SetOrigin(0, 0)

                distancedim = objdims.AddDistanceBetweenObjects(headerline, xe, ye, 0, True, svarc, xc, yc, 0, True)

                With distancedim
                    .ReattachToDrawingView(objDV)
                    .TrackDistance = nippledim.TrackDistance
                    .MeasurementAxisDirection = False
                    .MeasurementAxisEx = 1
                    SEDrawing.ControlDimDistance(distancedim, {"y", "smaller"})
                End With

            End If

        Catch ex As Exception

        Finally
            objDV.ScaleFactor = scalef
            objDV.SetOrigin(x0, y0)
        End Try

    End Sub

    Shared Function HorStutzenPos(objsheet As SolidEdgeDraft.Sheet, objDV As SolidEdgeDraft.DrawingView, headerdimlines() As SolidEdgeDraft.DVLine2d, headerframe() As Double, circuit As CircuitData,
                                  circlist As List(Of SolidEdgeDraft.DVCircle2d), headertype As String, rowcount As Integer, scalefactor As Double) As List(Of SolidEdgeDraft.DVCircle2d)
        Dim objdims As Dimensions = objsheet.Dimensions
        Dim firstdim, distancedim As Dimension
        Dim dimstyle As DimStyle
        Dim sortedcirclist As New List(Of SolidEdgeDraft.DVCircle2d)
        Dim stutzenCirclist As New List(Of StutzenDVElement)
        Dim xclist As New List(Of Double)
        Dim xc, yc, xh1, yh1, xc2, yc2, xe1, ye1, xh2, yh2, xe2, ye2, xmin, ymin, xmax, ymax, trackdistance, delta, x0, y0, xmin2, ymin2, xmax2, ymax2 As Double
        Dim breakpos As Integer = 1

        Try
            objDV.GetOrigin(x0, y0)
            objDV.Range(xmin2, ymin2, xmax2, ymax2)

            For i As Integer = 0 To circlist.Count - 1
                circlist(i).GetCenterPoint(xc, yc)

                If xc > headerframe(0) And xc < headerframe(2) And Math.Round(circlist(i).Diameter * 1000, 3) = circuit.CoreTube.Diameter Then
                    stutzenCirclist.Add(New StutzenDVElement With {.DVElement = circlist(i), .Ypos = Math.Round(xc, 6)})
                End If
            Next

            Dim sortedlist = From plist In stutzenCirclist Order By plist.Ypos

            For i As Integer = 0 To sortedlist.Count - 1
                sortedcirclist.Add(sortedlist(i).DVElement)
            Next

            'first circle and headerdimline
            sortedcirclist(0).GetCenterPoint(xc, yc)
            headerdimlines(0).GetStartPoint(xh1, yh1)
            headerdimlines(0).GetEndPoint(xe1, ye1)
            headerdimlines(1).GetStartPoint(xh2, yh2)
            headerdimlines(1).GetEndPoint(xe2, ye2)
            firstdim = objdims.AddDistanceBetweenObjects(sortedcirclist(0).Reference, xc, yc, 0, True, headerdimlines(0).Reference, xh1, yh1, 0, True)
            With firstdim
                .MeasurementAxisDirection = False
                .MeasurementAxisEx = 1
                .TrackDistance = 0.02
                .TextScale = 0.75
                If .Value < 0.025 Then
                    .TerminatorPosition = True
                    .BreakPosition = breakpos
                    .BreakDistance = 0.003
                Else
                    .BreakPosition = 2
                    .BreakDistance = 0.5
                End If
                If headertype = "inlet" Then
                    SEDrawing.ControlDimDistance(firstdim, {"y", "bigger"})
                Else
                    SEDrawing.ControlDimDistance(firstdim, {"y", "smaller"})
                End If
                dimstyle = .Style
            End With
            dimstyle.OrigTerminatorSize = 0.5
            dimstyle.TerminatorSize = 0.5

            trackdistance = firstdim.TrackDistance
            firstdim.Range(xmin, ymin, xmax, ymax)

            If headertype = "inlet" Then
                If rowcount = 1 Then
                    delta = ymax2 - ymax + 0.005
                    If delta > 0 Then
                        If trackdistance > 0 Then
                            trackdistance += delta
                        Else
                            trackdistance -= delta
                        End If
                    End If
                Else
                    'get range of previous row
                    objdims.Item(objdims.Count - 1).Range(xmin2, ymin2, xmax2, ymax2)
                    delta = Math.Abs(ymax2 - ymax)
                    If trackdistance > 0 Then
                        trackdistance += delta + 0.01
                    Else
                        trackdistance -= delta + 0.01
                    End If
                End If
            Else
                If rowcount = 1 Then
                    delta = ymin - ymin2 + 0.0075
                    If delta > 0 Then
                        If trackdistance > 0 Then
                            trackdistance += delta
                        Else
                            trackdistance -= delta
                        End If
                    End If
                Else
                    'get range of previous row
                    objdims.Item(objdims.Count - 1).Range(xmin2, ymin2, xmax2, ymax2)
                    delta = Math.Abs(ymin2 - ymin)
                    If trackdistance > 0 Then
                        trackdistance += delta + 0.01
                    Else
                        trackdistance -= delta + 0.01
                    End If
                End If
            End If
            firstdim.TrackDistance = trackdistance

            For i As Integer = 1 To sortedcirclist.Count - 1
                sortedcirclist(i - 1).GetCenterPoint(xc, yc)
                sortedcirclist(i).GetCenterPoint(xc2, yc2)
                distancedim = objdims.AddDistanceBetweenObjects(sortedcirclist(i - 1).Reference, xc, yc, 0, True, sortedcirclist(i).Reference, xc2, yc2, 0, True)
                With distancedim
                    .MeasurementAxisEx = 1
                    .MeasurementAxisDirection = False
                    .TextScale = 0.75
                    .BreakPosition = 2
                    .BreakDistance = 0.5
                    .TrackDistance = trackdistance
                    If headertype = "inlet" Then
                        SEDrawing.ControlDimDistance(distancedim, {"y", "bigger"})
                    Else
                        SEDrawing.ControlDimDistance(distancedim, {"y", "smaller"})
                    End If
                    dimstyle = .Style
                End With
                dimstyle.OrigTerminatorSize = 0.5
                dimstyle.TerminatorSize = 0.5
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Function

    Shared Sub RelocateCaption(objDV As SolidEdgeDraft.DrawingView, objsheet As SolidEdgeDraft.Sheet, headertype As String)
        Dim objdims As Dimensions = objsheet.Dimensions
        Dim x, y, xmin, ymin, xmax, ymax As Double

        Try
            objdims.Item(objdims.Count).Range(xmin, ymin, xmax, ymax)
            objDV.GetCaptionPosition(x, y)
            If headertype = "inlet" Then
                objDV.SetCaptionPosition(x, ymax + 0.0075)
            Else
                objDV.SetCaptionPosition(x, ymin - 0.01)
            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub RelocateFrontViews(frontDV As SolidEdgeDraft.DrawingView, backDV As SolidEdgeDraft.DrawingView, scalefactor As Double, inheaderdim As Dimension,
                                  outheaderdim As Dimension)
        Dim xcap, ycap, xin, yin, yout, xmin, ymin, xmax, ymax, xmin2, ymin2, xmax2, ymax2 As Double

        Try
            frontDV.ScaleFactor = 1
            frontDV.SetOrigin(0, 0)
            frontDV.Range(xmin, ymin, xmax, ymax)

            frontDV.GetCaptionPosition(xcap, ycap)
            yin = 0.41 - ycap - 0.005
            xin = scalefactor * (Math.Abs(xmin) + Math.Abs(xmax)) / 2 + 0.035
            frontDV.ScaleFactor = scalefactor
            frontDV.SetOrigin(xin, yin)

            backDV.ScaleFactor = scalefactor
            backDV.SetOrigin(0, 0)

            inheaderdim.Range(xmin, ymin, xmax, ymax)
            outheaderdim.Range(xmin2, ymin2, xmax2, ymax2)
            yout = ymin - ymax2 - 0.01

            backDV.SetOrigin(xin, yout)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub FillSupportTubes(dftdoc As SolidEdgeDraft.DraftDocument, coil As CoilData)
        Dim x, y As Double

        Try
            For Each objsheet As SolidEdgeDraft.Sheet In dftdoc.Sheets
                Dim finlines As List(Of SolidEdgeDraft.DVLine2d)
                Dim finframelines() As SolidEdgeDraft.DVLine2d
                Dim finframe(), dx, dy As Double

                If objsheet.Name.Contains("Coil") Then
                    For i As Integer = 0 To objsheet.DrawingViews.Count - 1
                        finlines = SEDrawing.GetLinesFromOcc(objsheet.DrawingViews.Item(i + 1), "Fin", "")
                        finframelines = SEDrawing.GetHorHeaderDimLines(finlines, coil.FinnedDepth)
                        finframe = SEDrawing.GetHeaderFrame(finframelines)

                        dx = Math.Min(finframe(0), finframe(2))
                        dy = Math.Min(finframe(1), finframe(3))

                        If coil.SupportTubesPosition(0).Count > 0 Then
                            Dim subsheet As SolidEdgeDraft.Sheet = objsheet.DrawingViews.Item(i + 1).Sheet
                            subsheet.Activate()
                            General.seapp.DoIdle()

                            For j As Integer = 0 To coil.SupportTubesPosition(0).Count - 1
                                x = coil.SupportTubesPosition(0)(j) + dx
                                y = coil.SupportTubesPosition(1)(j) + dy
                                If objsheet.DrawingViews.Item(i + 1).CaptionDefinitionTextPrimary.Contains("Back") Then
                                    x = coil.FinnedHeight / 1000 - coil.SupportTubesPosition(0)(j) + dx
                                End If
                                SEDrawing.CreateBoundary(objsheet.DrawingViews.Item(i + 1).Sheet, x, y, 15)
                            Next

                            objsheet.Activate()
                            General.seapp.DoIdle()
                        End If
                    Next

                End If

            Next

        Catch ex As Exception

        End Try

    End Sub

    Shared Sub SideViewDim(dftdoc As SolidEdgeDraft.DraftDocument, circuit As CircuitData, consys As ConSysData)
        Dim inletDV, outletDV As SolidEdgeDraft.DrawingView
        Dim finlines, headerlines As List(Of SolidEdgeDraft.DVLine2d)
        Dim finline, nippleline As SolidEdgeDraft.DVLine2d
        Dim headercircs As List(Of SolidEdgeDraft.DVCircle2d)
        Dim headercirc As SolidEdgeDraft.DVCircle2d
        Dim outnippledim, innippledim As Dimension
        Dim x, y, xmin, ymin, xmax, ymax, xmin2, ymin2, xmax2, ymax2, scalefactor As Double
        Dim norientation As String

        Try

#Region "Inlet"
            'place the DV
            inletDV = PlaceSideDV(dftdoc, 1)
            inletDV.GetOrigin(x, y)
            scalefactor = inletDV.ScaleFactor
            inletDV.ScaleFactor = 1
            inletDV.SetOrigin(0, 0)

            headerlines = SEDrawing.GetLinesFromOcc(inletDV, "InletHeader", "")
            'right vertical fin line
            finlines = SEDrawing.GetLinesFromOcc(inletDV, "Fin" + circuit.Coilnumber.ToString + circuit.CircuitNumber.ToString + ".par", "")

            finline = GetSpecificLine(finlines, {"x", "max"})
            headercircs = SEDrawing.GetCirclesFromOcc(inletDV, New List(Of String) From {"InletHeader"})
            If headercircs.Count = 0 Then
                'use cap
                headercircs = SEDrawing.GetCirclesFromOcc(inletDV, New List(Of String) From {consys.InletHeaders.First.Tube.TopCapID})
            End If

            headercirc = GetHeaderCircle(headercircs, consys.InletHeaders.First.Tube.Diameter)
            SetHeaderDistance(dftdoc.ActiveSheet, inletDV, headercirc, finline, "inlet", consys.InletHeaders.First.Dim_a)

            If consys.FlangeID <> "" Then
                nippleline = GetSpecificLine(SEDrawing.GetLinesFromOcc(inletDV, consys.FlangeID, ""), {"x", "max"})
            Else
                nippleline = GetSpecificLine(SEDrawing.GetLinesFromOcc(inletDV, "InletNipple", ""), {"x", "max"})
            End If

            innippledim = SideNippleDim(dftdoc.ActiveSheet, inletDV, finline, nippleline, "inlet")

            If consys.InletHeaders.First.Ventsize <> "" Then
                VentDiameter(dftdoc.ActiveSheet, inletDV, SEDrawing.GetLinesFromOcc(inletDV, "InletHeader", ""), headercirc, GNData.GetVentDiameter(consys.HeaderMaterial, consys.InletHeaders.First.Ventsize), "inlet")
            End If

            If consys.FlangeID = "" Then
                innippledim.TrackDistance = -0.025
            End If

            inletDV.ScaleFactor = scalefactor
            inletDV.SetOrigin(x, y)
            innippledim.Range(xmin, ymin, xmax, ymax)
            If xmax > 0.585 Then
                inletDV.SetOrigin(x - (xmax - 0.585), y)
            End If
            If ymax > 0.41 Then
                innippledim.TrackDistance += ymax - 0.41
            End If
#End Region

#Region "Outlet"
            'place the DV
            outletDV = PlaceSideDV(dftdoc, 2)
            outletDV.GetOrigin(x, y)
            outletDV.ScaleFactor = 1
            outletDV.SetOrigin(0, 0)

            headerlines = SEDrawing.GetLinesFromOcc(outletDV, "OutletHeader", "")
            'right vertical fin line
            finlines = SEDrawing.GetLinesFromOcc(outletDV, "Fin" + circuit.Coilnumber.ToString + circuit.CircuitNumber.ToString + ".par", "")

            finline = GetSpecificLine(finlines, {"x", "max"})
            headercircs = SEDrawing.GetCirclesFromOcc(outletDV, New List(Of String) From {"OutletHeader"})
            If headercircs.Count = 0 Then
                'use cap
                headercircs = SEDrawing.GetCirclesFromOcc(outletDV, New List(Of String) From {consys.OutletHeaders.First.Tube.TopCapID})
            End If

            headercirc = GetHeaderCircle(headercircs, consys.OutletHeaders.First.Tube.Diameter)
            SetHeaderDistance(dftdoc.ActiveSheet, outletDV, headercirc, finline, "outlet", consys.OutletHeaders.First.Dim_a)

            norientation = NippleOrientation(outletDV)

            If norientation = "horizontal" Then
                If consys.FlangeID <> "" Then
                    nippleline = GetSpecificLine(SEDrawing.GetLinesFromOcc(outletDV, consys.FlangeID, ""), {"x", "max"})
                Else
                    nippleline = GetSpecificLine(SEDrawing.GetLinesFromOcc(outletDV, "OutletNipple", ""), {"x", "max"})
                End If
                outnippledim = SideNippleDim(dftdoc.ActiveSheet, outletDV, finline, nippleline, "outlet")
            Else
                nippleline = GetSpecificLine(SEDrawing.GetLinesFromOcc(outletDV, "OutletNipple", ""), {"y", "min"})
                outnippledim = SideNippleDim2(dftdoc.ActiveSheet, outletDV, headercirc, nippleline)
            End If

            If consys.InletHeaders.First.Ventsize <> "" Then
                VentDiameter(dftdoc.ActiveSheet, outletDV, SEDrawing.GetLinesFromOcc(outletDV, "OutletHeader", ""), headercirc, GNData.GetVentDiameter(consys.HeaderMaterial, consys.OutletHeaders.First.Ventsize), "outlet")
            End If

            outletDV.ScaleFactor = scalefactor
            outletDV.SetOrigin(x, y)
            inletDV.Range(xmin, ymin, xmax, ymax)
            outletDV.Range(xmin2, ymin2, xmax2, ymax2)

            outletDV.SetOrigin(x - (xmin2 - xmin), y)

#End Region

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function PlaceSideDV(dftdoc As SolidEdgeDraft.DraftDocument, refno As Integer) As SolidEdgeDraft.DrawingView
        Dim refDV, sideDV As SolidEdgeDraft.DrawingView
        Dim xref, yref, xmin, ymin, xmax, ymax, xside As Double

        Try
            refDV = dftdoc.ActiveSheet.DrawingViews.Item(refno)
            refDV.GetOrigin(xref, yref)

            sideDV = dftdoc.ActiveSheet.DrawingViews.AddByFold(refDV, SolidEdgeDraft.FoldTypeConstants.igFoldRight, 0, 0)

            If refno = 1 Then
                sideDV.Configuration = "SideInlet"
            Else
                sideDV.Configuration = "SideOutlet"
            End If

            sideDV.MatchConfiguration = True
            sideDV.Update()

            refDV.SetOrigin(xref, yref)
            refDV.Range(xmin, ymin, xmax, ymax)
            'later relocation based on final dimension
            xside = xmax + (0.59 - xmax) / 2 + 0.0415 - sideDV.CropRight

            sideDV.SetOrigin(xside, yref)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return sideDV
    End Function

    Shared Function GetSpecificLine(objlines As List(Of SolidEdgeDraft.DVLine2d), comparer() As String) As SolidEdgeDraft.DVLine2d
        Dim xslist, yslist, xelist, yelist As New List(Of Double)
        Dim xs, ys, xe, ye As Double
        Dim refline As SolidEdgeDraft.DVLine2d

        Try
            For Each l In objlines
                l.GetStartPoint(xs, ys)
                l.GetEndPoint(xe, ye)
                xslist.Add(Math.Round(xs, 6))
                yslist.Add(Math.Round(ys, 6))
                xelist.Add(Math.Round(xe, 6))
                yelist.Add(Math.Round(ye, 6))
            Next

            For i As Integer = 0 To xslist.Count - 1
                If comparer(0) = "y" Then
                    If yslist(i) = yelist(i) Then
                        If comparer(1) = "max" AndAlso yslist(i) = yslist.Max Then
                            refline = objlines(i)
                            Exit For
                        ElseIf comparer(1) = "min" AndAlso yslist(i) = yslist.Min Then
                            refline = objlines(i)
                            Exit For
                        End If
                    End If
                Else
                    If xslist(i) = xelist(i) Then
                        If comparer(1) = "max" AndAlso xslist(i) = xslist.Max Then
                            refline = objlines(i)
                            Exit For
                        ElseIf comparer(1) = "min" AndAlso xslist(i) = xslist.Min Then
                            refline = objlines(i)
                            Exit For
                        End If
                    End If
                End If
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
        Return refline
    End Function

    Shared Function GetHeaderCircle(headercircs As List(Of SolidEdgeDraft.DVCircle2d), diameter As Double) As SolidEdgeDraft.DVCircle2d

        Try
            For Each c In headercircs
                If Math.Round(c.Diameter * 1000, 6) = diameter Then
                    Return c
                End If
            Next
        Catch ex As Exception

        End Try

    End Function

    Shared Sub SetHeaderDistance(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headercirc As SolidEdgeDraft.DVCircle2d, finline As SolidEdgeDraft.DVLine2d,
                                 headertype As String, distance_a As Double)
        Dim objdims As Dimensions = objSheet.Dimensions
        Dim hordistance, dim_a, dim_a_, vertdistance As Dimension
        Dim xlist, ylist As New List(Of Double)
        Dim xc, yc, xm, ym, z, xs, ys, xe, ye, xref1, yref1 As Double
        Dim htype, ktype As Integer
        Dim comparer1, comparer2 As String

        Try
            If headertype = "inlet" Then
                comparer1 = "bigger"
                comparer2 = "smaller"
            Else
                comparer1 = "smaller"
                comparer2 = "bigger"
            End If
            'use different trackdistance for inlet & outlet

            finline.GetStartPoint(xs, ys)
            finline.GetEndPoint(xe, ye)

            If (ys > ye And headertype = "inlet") Or (ys < ye And headertype = "outlet") Then
                xref1 = xs
                yref1 = ys
            Else
                xref1 = xe
                yref1 = ye
            End If

            headercirc.GetCenterPoint(xc, yc)
            For i As Integer = 0 To headercirc.KeyPointCount - 1
                headercirc.GetKeyPoint(i, xm, ym, z, ktype, htype)
                If Math.Round(xc - xm, 5) > headercirc.Radius / 2 Then
                    Debug.Print("keypoint: " + i.ToString + " x: " + Math.Round(xm, 6).ToString + " / y:" + Math.Round(ym, 6).ToString)
                    Exit For
                End If
            Next

            'minimum distance between header and tube sheet
            dim_a = objdims.AddDistanceBetweenObjectsEX(finline, xref1, yref1, 0, True, False, headercirc, xm, ym, 0, False, True)

            With dim_a
                .MeasurementAxisEx = 3
                .MeasurementAxisDirection = False
                .TrackDistance = 0.0125
                .ReattachToDrawingView(DV)
                SEDrawing.ControlDimDistance(dim_a, {"y", comparer1})

                If Math.Round(.Value * 1000, 3) > distance_a Then
                    dim_a.Delete()
                    'draw small line in the 
                    Dim objline As Line2d = DrawDimLine(headercirc, DV.Sheet, objSheet)
                    If objline IsNot Nothing Then
                        objline.GetStartPoint(xm, ym)
                        dim_a_ = objdims.AddDistanceBetweenObjects(finline, xref1, yref1, 0, True, objline, xm, ym, 0, False)
                        With dim_a_
                            .MeasurementAxisEx = 3
                            .MeasurementAxisDirection = True
                            .TrackDistance = 0.0125
                            .ReattachToDrawingView(DV)
                            SEDrawing.ControlDimDistance(dim_a_, {"y", comparer1})
                        End With
                    End If
                End If
            End With

            'distance between tube sheet and center of header
            hordistance = objdims.AddDistanceBetweenObjects(finline, xref1, yref1, 0, True, headercirc, xc, yc, 0, True)

            With hordistance
                .ReattachToDrawingView(DV)
                .MeasurementAxisEx = 1
                .MeasurementAxisDirection = False
                .TrackDistance = 0.0345
                SEDrawing.ControlDimDistance(hordistance, {"y", comparer2})
            End With

            vertdistance = objdims.AddDistanceBetweenObjects(finline, xref1, yref1, 0, True, headercirc, xc, yc, 0, True)

            With vertdistance
                .ReattachToDrawingView(DV)
                .MeasurementAxisEx = 1
                .MeasurementAxisDirection = True
                .TerminatorPosition = True
                .BreakPosition = 3
                .BreakDistance = 0.003
                .TrackDistance = 0.008
                SEDrawing.ControlDimDistance(vertdistance, {"x", "smaller"})
            End With

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function DrawDimLine(headercirc As SolidEdgeDraft.DVCircle2d, subsheet As SolidEdgeDraft.Sheet, mainsheet As SolidEdgeDraft.Sheet) As Line2d
        Dim xc, yc As Double
        Dim objline As Line2d

        Try
            subsheet.Activate()
            General.seapp.DoIdle()


            headercirc.GetCenterPoint(xc, yc)

            objline = subsheet.Lines2d.AddBy2Points(xc - headercirc.Radius, yc, xc - headercirc.Radius, yc - 0.001)
            objline.Style.Width = 0.0002

            mainsheet.Activate()
            General.seapp.DoIdle()
        Catch ex As Exception

        End Try

#Disable Warning BC42104 ' Die objline-Variable wird verwendet, bevor ihr ein Wert zugewiesen wird. Zur Laufzeit kann eine Nullverweisausnahme auftreten.
        Return objline
#Enable Warning BC42104 ' Die objline-Variable wird verwendet, bevor ihr ein Wert zugewiesen wird. Zur Laufzeit kann eine Nullverweisausnahme auftreten.
    End Function

    Shared Function DrawDimLine2(subsheet As SolidEdgeDraft.Sheet, mainsheet As SolidEdgeDraft.Sheet, x As Double, y As Double) As Line2d
        Dim objline As Line2d
        Try
            subsheet.Activate()
            General.seapp.DoIdle()

            objline = subsheet.Lines2d.AddBy2Points(x, y, x, y - 0.001)
            objline.Style.Width = 0.0002

            mainsheet.Activate()
            General.seapp.DoIdle()
        Catch ex As Exception

        End Try

#Disable Warning BC42104 ' Die objline-Variable wird verwendet, bevor ihr ein Wert zugewiesen wird. Zur Laufzeit kann eine Nullverweisausnahme auftreten.
        Return objline
#Enable Warning BC42104 ' Die objline-Variable wird verwendet, bevor ihr ein Wert zugewiesen wird. Zur Laufzeit kann eine Nullverweisausnahme auftreten.
    End Function

    Shared Function GetMiddlePoint(dvelementlist As List(Of StutzenDVElement), headertype As String, alignment As String) As Double
        Dim xs1, xe1, ys1, ye1, xs2, xe2, ys2, ye2, xc As Double
        Dim shortline, longline As SolidEdgeDraft.DVLine2d

        Try
            shortline = dvelementlist(0).DVElement
            shortline.GetStartPoint(xs1, ys1)
            shortline.GetEndPoint(xe1, ye1)

            longline = dvelementlist(1).DVElement
            longline.GetStartPoint(xs2, ys2)
            longline.GetEndPoint(xe2, ye2)

            If (headertype = "inlet" And alignment = "horizontal") Or (headertype = "outlet" And alignment = "vertical") Then
                xc = (Math.Min(xs1, xe1) + Math.Max(xs2, xe2)) / 2
            Else
                xc = (Math.Min(xs2, xe2) + Math.Max(xs1, xe1)) / 2
            End If

        Catch ex As Exception

        End Try
        Return Math.Round(xc, 6)
    End Function

    Shared Function SideNippleDim(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, finline As SolidEdgeDraft.DVLine2d, nippleline As SolidEdgeDraft.DVLine2d,
                                  headertype As String) As Dimension
        Dim objdims As Dimensions = objSheet.Dimensions
        Dim hordistance, diadim As Dimension
        Dim xlist, ylist As New List(Of Double)
        Dim xc, yc, xs, ys, xe, ye, xref1, yref1 As Double

        Try

            finline.GetStartPoint(xs, ys)
            finline.GetEndPoint(xe, ye)
            nippleline.GetStartPoint(xc, yc)

            If (ys > ye And headertype = "inlet") Or (ys < ye And headertype = "outlet") Then
                xref1 = xs
                yref1 = ys
            Else
                xref1 = xe
                yref1 = ye
            End If

            'distance tube sheet
            hordistance = objdims.AddDistanceBetweenObjects(finline, xref1, yref1, 0, False, nippleline, xc, yc, 0, False)

            With hordistance
                .ReattachToDrawingView(DV)
                .MeasurementAxisEx = 1
                .MeasurementAxisDirection = True
                .TrackDistance = 0.04
                SEDrawing.ControlDimDistance(hordistance, {"y", "smaller"})
            End With

            'diameter
            nippleline.GetStartPoint(xs, ys)
            nippleline.GetEndPoint(xe, ye)

            diadim = objdims.AddDistanceBetweenObjects(nippleline, xs, ys, 0, True, nippleline, xe, ye, 0, True)

            With diadim
                .ReattachToDrawingView(DV)
                .PrefixString = "%DI"
                .BreakPosition = 1
                .BreakDistance = 0.005
                .TerminatorPosition = True
                .TrackDistance = 0.01
                SEDrawing.ControlDimDistance(diadim, {"x", "bigger"})
            End With

        Catch ex As Exception

        End Try

#Disable Warning BC42104 ' Die hordistance-Variable wird verwendet, bevor ihr ein Wert zugewiesen wird. Zur Laufzeit kann eine Nullverweisausnahme auftreten.
        Return hordistance
#Enable Warning BC42104 ' Die hordistance-Variable wird verwendet, bevor ihr ein Wert zugewiesen wird. Zur Laufzeit kann eine Nullverweisausnahme auftreten.
    End Function

    Shared Function SideNippleDim2(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headercirc As SolidEdgeDraft.DVCircle2d, nippleline As SolidEdgeDraft.DVLine2d) As Dimension
        Dim objdims As Dimensions = objSheet.Dimensions
        Dim hordistance, diadim As Dimension
        Dim xlist, ylist As New List(Of Double)
        Dim xc, yc, xs, ys, xe, ye, xref1, yref1 As Double

        Try

            headercirc.GetCenterPoint(xref1, yref1)
            nippleline.GetStartPoint(xc, yc)

            'vertical tube sheet
            hordistance = objdims.AddDistanceBetweenObjects(headercirc, xref1, yref1, 0, False, nippleline, xc, yc, 0, False)

            With hordistance
                .ReattachToDrawingView(DV)
                .MeasurementAxisEx = 1
                .MeasurementAxisDirection = True
                .TrackDistance = 0.025
                SEDrawing.ControlDimDistance(hordistance, {"x", "bigger"})
            End With

            'diameter
            nippleline.GetStartPoint(xs, ys)
            nippleline.GetEndPoint(xe, ye)

            diadim = objdims.AddDistanceBetweenObjects(nippleline, xs, ys, 0, True, nippleline, xe, ye, 0, True)

            With diadim
                .ReattachToDrawingView(DV)
                .PrefixString = "%DI"
                .BreakPosition = 1
                .BreakDistance = 0.005
                .TerminatorPosition = True
                .TrackDistance = 0.01
                SEDrawing.ControlDimDistance(diadim, {"y", "smaller"})
            End With

        Catch ex As Exception

        End Try

#Disable Warning BC42104 ' Die hordistance-Variable wird verwendet, bevor ihr ein Wert zugewiesen wird. Zur Laufzeit kann eine Nullverweisausnahme auftreten.
        Return hordistance
#Enable Warning BC42104 ' Die hordistance-Variable wird verwendet, bevor ihr ein Wert zugewiesen wird. Zur Laufzeit kann eine Nullverweisausnahme auftreten.
    End Function

    Shared Sub VentDiameter(objSheet As SolidEdgeDraft.Sheet, DV As SolidEdgeDraft.DrawingView, headerlines As List(Of SolidEdgeDraft.DVLine2d), headercirc As SolidEdgeDraft.DVCircle2d,
                            ventdiameter As Double, headertype As String)
        Dim objdims As Dimensions = objSheet.Dimensions
        Dim diadim As Dimension
        Dim dimlines As New List(Of SolidEdgeDraft.DVLine2d)
        Dim xs, ys, xe, ye, xc, yc As Double
        Dim xslist, yslist, xelist, yelist, lengthlist As New List(Of Double)
        Dim vertlines As New List(Of StutzenDVElement)
        Dim comparer As String
        Dim breakpos As Integer

        Try
            If headertype = "inlet" Then
                comparer = "bigger"
                breakpos = 3
            Else
                comparer = "smaller"
                breakpos = 1
            End If

            'center of header
            headercirc.GetCenterPoint(xc, yc)

            For Each l In headerlines
                l.GetStartPoint(xs, ys)
                l.GetEndPoint(xe, ye)
                xslist.Add(Math.Round(xs * 1000, 3))
                yslist.Add(Math.Round(ys * 1000, 3))
                xelist.Add(Math.Round(xe * 1000, 3))
                yelist.Add(Math.Round(ye * 1000, 3))
                lengthlist.Add(Math.Round(l.Length * 1000, 3))
                If Math.Round(xs * 1000, 3) = Math.Round(xe * 1000, 3) Then
                    vertlines.Add(New StutzenDVElement With {.DVElement = l, .Xpos = Math.Round(xe * 1000, 3)})
                End If
            Next

            For Each objE In vertlines
                If Math.Round(Math.Abs(Math.Round(xc * 1000, 3) - objE.Xpos), 2) = ventdiameter / 2 Then
                    dimlines.Add(objE.DVElement)
                End If
            Next

            If dimlines.Count = 2 Then
                dimlines(0).GetStartPoint(xs, ys)
                dimlines(1).GetStartPoint(xe, ye)

                diadim = objdims.AddDistanceBetweenObjects(dimlines(0), xs, ys, 0, True, dimlines(1), xe, ye, 0, True)
                With diadim
                    .ReattachToDrawingView(DV)
                    .MeasurementAxisDirection = False
                    .MeasurementAxisEx = 1
                    .PrefixString = "%DI"
                    .BreakPosition = breakpos
                    .BreakDistance = 0.005
                    .TerminatorPosition = True
                    .TrackDistance = 0.016
                    SEDrawing.ControlDimDistance(diadim, {"y", comparer})
                End With
            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub RelocateHeaderDim(headerdim As Dimension, objDV As SolidEdgeDraft.DrawingView)
        Dim delta, xmin1, ymin1, xmax1, ymax1, xmin2, ymin2, xmax2, ymax2 As Double

        headerdim.Range(xmin1, ymin1, xmax1, ymax1)
        objDV.Range(xmin2, ymin2, xmax2, ymax2)

        delta = ymax2 - ymax1

        If headerdim.TrackDistance > 0 Then
            headerdim.TrackDistance += delta + 0.01
        Else
            headerdim.TrackDistance -= delta + 0.01
        End If
    End Sub

    Shared Function NippleOrientation(objDV As SolidEdgeDraft.DrawingView) As String
        Dim orientation As String
        Dim nippleDVelements As New List(Of StutzenDVElement)
        Dim xs, ys, xe, ye As Double

        For Each l As SolidEdgeDraft.DVLine2d In objDV.DVLines2d
            If l.ModelMember.FileName.Contains("Nipple") Then
                nippleDVelements.Add(New StutzenDVElement With {.DVElement = l, .Xpos = Math.Round(l.Length, 6)})
            End If
        Next

        Dim templist = From plist In nippleDVelements Order By plist.Xpos Descending

        Dim longestline As SolidEdgeDraft.DVLine2d = templist.ToList.First.DVElement

        longestline.GetStartPoint(xs, ys)
        longestline.GetEndPoint(xe, ye)

        If Math.Round(xs, 6) = Math.Round(xe, 6) Then
            orientation = "vertical"
        Else
            orientation = "horizontal"
        End If

        Return orientation
    End Function

    Shared Sub IsoView(dftdoc As SolidEdgeDraft.DraftDocument, circuit As CircuitData, coil As CoilData)
        Dim isoDV, helpDV As SolidEdgeDraft.DrawingView
        Dim asmlink As SolidEdgeDraft.ModelLink = dftdoc.ModelLinks.Item(1)
        Dim objShapes As AnnotAlignmentShapes
        Dim objPartLists As SolidEdgeDraft.PartsLists
        Dim partlist As SolidEdgeDraft.PartsList
        Dim objBalloons As Balloons
        Dim bscale, x1, y1, x2, y2, x0, y0 As Double
        Dim x1list, y1list, x2list, y2list As New List(Of Double)
        Dim objColumn As SolidEdgeDraft.TableColumn
        Dim proptext, postext As String

        Try
            helpDV = dftdoc.ActiveSheet.DrawingViews.AddAssemblyView(asmlink, SolidEdgeDraft.ViewOrientationConstants.igFrontView, 0.05, 0, 0, 0)
            isoDV = dftdoc.ActiveSheet.DrawingViews.AddByFold(helpDV, SolidEdgeDraft.FoldTypeConstants.igFoldDownLeft, 0, 0)

            helpDV.Delete()

            'place partlist
            objPartLists = dftdoc.PartsLists
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
                partlist.AnchorPoint = SolidEdgeDraft.TableAnchorPoint.igLowerRight
                partlist.FillEndOfTableWithBlankRows = False
                partlist.ShowTopAssembly = False
                bscale = 2
                partlist.SetOrigin(0.409, 0.005)
                partlist.Update()
                partlist.ListType = SolidEdgeDraft.PartsListType.igExploded
            End If

            partlist.Update()
            objBalloons = dftdoc.ActiveSheet.Balloons

            For Each sBalloon As Balloon In objBalloons
                sBalloon.DisplayItemCount = False
                sBalloon.Range(x1, y1, x2, y2)
                sBalloon.TextScale = bscale
            Next

            partlist.Columns.Item(1).DataHorizontalAlignment = TextHorizontalAlignmentConstants.igTextHzAlignCenter
            partlist.Columns.Item(1).Width = 0.01

            objColumn = partlist.Columns.Item(4)
            proptext = objColumn.PropertyText
            objColumn = partlist.Columns.Item(1)
            objColumn.Header = "Pos."
            objColumn.Width = 0.015
            objColumn.DataHorizontalAlignment = TextHorizontalAlignmentConstants.igTextHzAlignCenter

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
            objColumn.HeaderRowHorizontalAlignment = TextHorizontalAlignmentConstants.igTextHzAlignCenter

            objColumn = partlist.Columns.Item(2)
            objColumn.Header = "Quantity"
            objColumn.Width = 0.015
            objColumn.DataHorizontalAlignment = TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.PropertyText = proptext
            objColumn.HeaderRowHorizontalAlignment = TextHorizontalAlignmentConstants.igTextHzAlignCenter

            objColumn = partlist.Columns.Item(3)
            objColumn.Header = "ERP Number"
            objColumn.Width = 0.025
            objColumn.DataHorizontalAlignment = TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.PropertyText = "%{CDB_ERP_Artnr./CP|G}"
            objColumn.HeaderRowHorizontalAlignment = TextHorizontalAlignmentConstants.igTextHzAlignCenter

            objColumn = partlist.Columns.Item(4)
            objColumn.Header = "Designation"
            objColumn.Width = 0.05
            objColumn.DataHorizontalAlignment = TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.PropertyText = "%{CDB_Benennung_de/CP|G}"

            objColumn = partlist.Columns.Add(5, True)

            objColumn.Header = "PDM Number"
            objColumn.Width = 0.03
            objColumn.DataHorizontalAlignment = TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.PropertyText = "%{CDB_teilenummer/CP|G}"
            objColumn.HeaderRowHorizontalAlignment = TextHorizontalAlignmentConstants.igTextHzAlignCenter
            objColumn.Show = True

            partlist.Update()

            'if odd passnumber, move balloons and add break view
            If circuit.NoPasses Mod 2 > 0 Then
                IsoOdd(dftdoc, isoDV)
            Else
                'rescale and position
                objShapes = dftdoc.ActiveSheet.AnnotAlignmentShapes
                objShapes.Item(1).Delete()
                If coil.FinnedHeight < 1000 Then
                    isoDV.ScaleFactor = 0.15
                ElseIf coil.FinnedHeight < 1200 Then
                    isoDV.ScaleFactor = 0.1
                Else
                    isoDV.ScaleFactor = 1 / 15
                End If

                isoDV.GetOrigin(x0, y0)

                For Each sBalloon As Balloon In dftdoc.ActiveSheet.Balloons
                    sBalloon.Range(x1, y1, x2, y2)
                    y1list.Add(Math.Round(y1, 6))
                    y2list.Add(Math.Round(y2, 6))
                Next
                y0 -= (Math.Min(y1list.Min, y2list.Min) - 0.015)
                isoDV.SetOrigin(0.147, y0)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub IsoOdd(dftdoc As SolidEdgeDraft.DraftDocument, isoDV As SolidEdgeDraft.DrawingView)
        Dim xmin1, ymin1, xmax1, ymax1, xmin2, ymin2, delta, xcap, ycap, sf As Double
        Dim xmin1list, ymin1list, xmax1list, ymax1list, xmin2list, ymin2list, xmax2list, ymax2list As New List(Of Double)
        Dim objBalloons As Balloons
        Dim objShapes As AnnotAlignmentShapes

        Try
            objShapes = dftdoc.ActiveSheet.AnnotAlignmentShapes
            objShapes.Item(1).Delete()
            sf = isoDV.ScaleFactor
            objBalloons = dftdoc.ActiveSheet.Balloons

            For Each sBalloon As Balloon In objBalloons
                sBalloon.DisplayItemCount = False
                sBalloon.Range(xmin1, ymin1, xmax1, ymax1)
                sBalloon.TextScale = 2
                xmin1list.Add(Math.Round(xmin1, 6))
                ymin1list.Add(Math.Round(ymin1, 6))
                xmax1list.Add(Math.Round(xmax1, 6))
                ymax1list.Add(Math.Round(ymax1, 6))
            Next

            'gap = f ( finned length, scale factor, angle )
            'get vertical finlines and compare the position with balloons to find horizontal limit
            Dim finlines As List(Of SolidEdgeDraft.DVLine2d) = SEDrawing.GetLinesFromOcc(isoDV, "Fin", "")

            For Each f In finlines
                f.GetStartPoint(xmin1, ymin1)
                f.GetEndPoint(xmin2, ymin2)
                xmin1list.Add(Math.Round(Math.Min(xmin1, xmin2) * sf, 6))
                ymin1list.Add(Math.Round(Math.Min(ymin1, ymin2) * sf, 6))
                xmax1list.Add(Math.Round(Math.Max(xmin1, xmin2) * sf, 6))
                ymax1list.Add(Math.Round(Math.Max(ymin1, ymin2) * sf, 6))
            Next

            'lowest balloon with ymin > 0 and highest ballon with ymax < 0
            Dim tupperlist = From plist In ymin1list Where plist > 0

            Dim upperlist As List(Of Double) = tupperlist.ToList

            Dim tlowerlist = From plist In ymax1list Where plist < 0

            Dim lowerlist As List(Of Double) = tlowerlist.ToList

            Dim tleftlist = From plist In xmax1list Where plist < 0

            Dim leftlist As List(Of Double) = tleftlist.ToList

            Dim trightlist = From plist In xmin1list Where plist > 0

            Dim rightlist As List(Of Double) = trightlist.ToList

            Debug.Print("ymax boundary: " + upperlist.Min.ToString)
            Debug.Print("ymin boundary: " + lowerlist.Max.ToString)
            Debug.Print("xmax boundary: " + rightlist.Min.ToString)
            Debug.Print("xmin boundary: " + leftlist.Max.ToString)

            'add break lines
            isoDV.SetOrigin(0, 0)
            Dim vertlines As SolidEdgeDraft.BreakLinePair = isoDV.BreakLinePairs.Add(SolidEdgeDraft.BreakLinePairDirConstants.igBreakLinePairDirConstants_Vertical, leftlist.Max + 0.005, rightlist.Min - 0.005, False)
            vertlines.Gap = 0.005

            General.seapp.DoIdle()
            isoDV.Update()

            Dim horlines As SolidEdgeDraft.BreakLinePair = isoDV.BreakLinePairs.Add(SolidEdgeDraft.BreakLinePairDirConstants.igBreakLinePairDirConstants_Horizontal, lowerlist.Max + 0.005, upperlist.Min - 0.005, False)
            horlines.Gap = 0.005

            isoDV.Range(xmin1, ymin1, xmax1, ymax1)
            delta = Math.Abs(Math.Round(ymax1 - ymin1, 6))

            dftdoc.ActiveSheet.DrawingViews.Item(2).GetCaptionPosition(xcap, ycap)
            If ycap < delta + 0.01 Then
                'rescaling needed
                Do
                    Dim newscalefactor As Double = RescaleIso(Math.Round(isoDV.ScaleFactor, 4))
                    isoDV.ScaleFactor = newscalefactor
                    isoDV.Range(xmin1, ymin1, xmax1, ymax1)
                    delta = Math.Abs(Math.Round(ymax1 - ymin1, 6))
                Loop Until ycap >= delta + 0.01
            End If

            xmin1list.Clear()
            ymin1list.Clear()
            xmax1list.Clear()
            ymax1list.Clear()

            'get range of isoDV
            xmin1list.Add(Math.Round(xmin1, 6))
            ymin1list.Add(Math.Round(ymin1, 6))
            xmax1list.Add(Math.Round(xmax1, 6))
            ymax1list.Add(Math.Round(ymax1, 6))

            'get balloon positions again. move DV so left balloon → xmin = 0.023, bottom balloon → ymin = 0.007
            For Each sBalloon As Balloon In objBalloons
                sBalloon.DisplayItemCount = False
                sBalloon.Range(xmin1, ymin1, xmax1, ymax1)
                sBalloon.TextScale = 2
                xmin1list.Add(Math.Round(xmin1, 6))
                ymin1list.Add(Math.Round(ymin1, 6))
                xmax1list.Add(Math.Round(xmax1, 6))
                ymax1list.Add(Math.Round(ymax1, 6))
            Next

            isoDV.SetOrigin(Math.Abs(xmin1list.Min) + 0.023, Math.Abs(ymin1list.Min) + 0.007)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Function RescaleIso(scalefactor As Double) As Double
        Dim scalefactorlist As New List(Of Double) From {0.03, 0.03333, 0.04, 0.05, 0.1}
        Dim newsf As Double

        If scalefactor <= scalefactorlist.Min OrElse scalefactorlist.IndexOf(scalefactor) = -1 Then
            newsf = 0.9 * scalefactor
        Else
            newsf = scalefactorlist(scalefactorlist.IndexOf(scalefactor) - 1)
        End If

        Return newsf
    End Function

End Class
