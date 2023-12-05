Public Class CSData
    Public Shared skeys, newStutzen As New List(Of String)
    Public Shared newAngles As New List(Of Double)

    Shared Sub GetCSData(ByRef consys As ConSysData, circuit As CircuitData, ByRef headertube As HeaderData, coil As CoilData)
        Dim neutralmrs, neutralmrn, RR, passes, pressure As String
        Dim csdata(), newoverhang As Double

        Try
            neutralmrs = GetNeutralModelRangeSuffix(General.currentunit.ModelRangeSuffix)
            neutralmrn = GetNeutralModelRangeName(General.currentunit.ModelRangeName)
            passes = "NULL"

            If circuit.IsOnebranchEvap Then
                csdata = {110, 0, 0, 0, 0}
                If circuit.FinType = "F" And circuit.Pressure < 17 And consys.HeaderMaterial = "C" Then
                    csdata(0) = 105
                End If
            Else
                pressure = circuit.Pressure
                If General.currentunit.ApplicationType.Contains("Condenser") Then
                    If Not neutralmrs.Contains("A") Then
                        neutralmrs = "XX"
                    End If
                    If circuit.Pressure > 100 Then
                        pressure = "130"
                    ElseIf General.currentunit.UnitDescription <> "VShape" Then
                        pressure = "46"
                    End If
                    RR = coil.NoRows
                    If General.currentunit.UnitDescription.Contains("VShape") And General.currentunit.UnitSize <> "Compact" Then
                        passes = circuit.NoPasses
                        neutralmrs = "XX"
                        If circuit.Pressure = 32 And consys.HeaderMaterial = "C" Then
                            pressure = "46"
                        End If
                    End If
                Else
                    RR = "NULL"
                End If

                csdata = Database.GetCSData(neutralmrn, neutralmrs, consys.HeaderAlignment, pressure, circuit.CoreTube.Material.Substring(0, 1), circuit.FinType, headertube.Tube.Diameter, headertube.Tube.HeaderType, RR, passes)

                'change some values for horizontal headers
                If consys.HeaderAlignment = "horizontal" And General.currentunit.ApplicationType = "Evaporator" Then
                    newoverhang = Calculation.CreateNewOverhang(headertube.Tube.HeaderType, headertube.Tube.Materialcodeletter, circuit.NoDistributions)

                    If circuit.NoDistributions = 1 Then
                        csdata(4) = 13.5
                        csdata(0) = Math.Round(110 - headertube.Tube.Diameter / 2, 2)
                        If (circuit.FinType = "N" Or circuit.FinType = "M") And consys.HeaderMaterial = "C" Then
                            csdata(0) += 5
                        End If
                        If headertube.Tube.HeaderType = "outlet" Then
                            csdata(2) = -100
                        End If
                    End If
                    If headertube.Tube.HeaderType = "outlet" And neutralmrs.Substring(0, 1) = "A" And headertube.Tube.Diameter = 26.9 Then
                        csdata(4) = 41
                    End If
                    'if it is mirrored, swap the values for overhang top and bottom
                    If circuit.ConnectionSide = "left" Then
                        csdata(3) = csdata(4)
                        csdata(4) = newoverhang
                    Else
                        csdata(3) = newoverhang
                    End If
                End If
                If circuit.ConnectionSide = "left" Then
                    csdata(1) = -csdata(1)
                End If
                If circuit.CoreTubeOverhang = 30 And circuit.NoDistributions > 1 And neutralmrs <> "CP" And General.currentunit.ApplicationType = "Evaporator" Then
                    csdata(0) += 5
                End If
                If circuit.Orbitalwelding Then
                    csdata(0) += 20
                End If
            End If

            With headertube
                .Dim_a = csdata(0)
                .Displacehor = csdata(1)
                .Displacever = csdata(2)
                .Overhangtop = csdata(3)
                .Overhangbottom = csdata(4)
            End With

        Catch ex As Exception

        End Try

    End Sub

    Shared Sub GetBrineCSData(ByRef consys As ConSysData, circuit As CircuitData)
        If circuit.NoDistributions = 1 Then
            consys.InletHeaders.First.Dim_a = 110
            consys.OutletHeaders.First.Dim_a = consys.InletHeaders.First.Dim_a
        Else
            consys.InletHeaders.First.Dim_a = 130
            If consys.HeaderAlignment = "horizontal" Then
                consys.OutletHeaders.First.Dim_a = 130
                consys.OutletHeaders.First.Overhangbottom = 15
            Else
                consys.OutletHeaders.First.Dim_a = 215
                consys.OutletHeaders.First.Overhangbottom = 18
            End If
            consys.OutletHeaders.First.Overhangtop = consys.OutletHeaders.First.Overhangbottom
            consys.InletHeaders.First.Overhangbottom = consys.OutletHeaders.First.Overhangbottom
            consys.InletHeaders.First.Overhangtop = consys.OutletHeaders.First.Overhangtop
            If circuit.CoreTubeOverhang = 30 Then
                consys.InletHeaders.First.Dim_a += 5
                consys.OutletHeaders.First.Dim_a += 5
            End If
        End If
    End Sub

    Shared Sub GetGADCCSData(ByRef consys As ConSysData, circuit As CircuitData, ByRef headertube As HeaderData)
        Dim neutralmrs, neutralmrn, RR, passes, pressure As String
        Dim csdata(4), headera, headeroverhangb As Double

        Try
            neutralmrs = GetNeutralModelRangeSuffix(General.currentunit.ModelRangeSuffix)

            If neutralmrs = "FP" Then
                If circuit.NoDistributions = 1 Then
                    If headertube.Tube.HeaderType = "inlet" Then
                        csdata = {145, 0, 0, 12.5, 12.5}
                    Else
                        csdata = {90, 25, 0, 12.5, 12.5}
                    End If
                Else
                    If headertube.Tube.HeaderType = "inlet" Then
                        csdata = {127.5, 0, 0, 9.5, 9.5}
                    Else
                        csdata = {72.5, 0, 0, 9.5, 9.5}
                    End If
                End If
            ElseIf neutralmrs = "RX" Then
                If circuit.IsOnebranchEvap Then
                    If headertube.Tube.HeaderType = "inlet" Then
                        csdata = {65, 0, 0, 0, 0}
                    Else
                        csdata = {85, 0, 0, 0, 0}
                    End If
                ElseIf circuit.NoDistributions = 1 Then
                    If headertube.Tube.HeaderType = "pot" Then
                        csdata = {90, 0, 0, 24.7, 10.3}
                    Else
                        csdata = {90, 0, 0, 0, 0}
                    End If
                Else
                    If headertube.Tube.HeaderType = "pot" Then
                        Select Case headertube.Tube.Diameter
                            Case 42
                                If consys.OutletHeaders.First.Tube.Diameter = 28 Then
                                    csdata = {70, 0, 0, 18.5, 16.5}
                                Else
                                    csdata = {70, 0, 0, 17, 18}
                                End If
                            Case 54
                                csdata = {70, 0, 0, 22.7, 22.8}
                            Case 64
                                csdata = {70, 0, 0, 25.5, 24.5}
                        End Select
                    Else
                        headera = Math.Round(consys.OilSifons.First.Tube.Diameter / 2 + 70 - headertube.Tube.Diameter / 2, 1)
                        If headertube.Tube.Diameter = 28 Then
                            headeroverhangb = 18.5
                        Else
                            headeroverhangb = 22
                        End If
                        csdata = {headera, 0, 0, 10, headeroverhangb}
                    End If
                End If
            Else
                If circuit.IsOnebranchEvap Then
                    If headertube.Tube.HeaderType = "inlet" Then
                        csdata = {65, 0, 0, 0, 0}
                    Else
                        csdata = {85, 0, 0, 0, 0}
                    End If
                ElseIf circuit.NoDistributions = 1 Then
                    If headertube.Tube.HeaderType = "pot" Then
                        csdata = {90, 0, 0, 24.7, 10.3}
                    Else
                        csdata = {90, 0, 0, 0, 0}
                    End If
                Else
                    If headertube.Tube.HeaderType = "pot" Then
                        csdata = {70, 0, 0, 18.8, 16.2}
                    Else
                        headera = Math.Round(consys.OilSifons.First.Tube.Diameter / 2 + 70 - headertube.Tube.Diameter / 2, 1)
                        csdata = {headera, 0, -37.5, 0, 25}
                    End If
                End If
            End If

            With headertube
                .Dim_a = csdata(0)
                .Displacehor = csdata(1)
                .Displacever = csdata(2)
                .Overhangtop = csdata(3)
                .Overhangbottom = csdata(4)
            End With
        Catch ex As Exception

        End Try

    End Sub

    Shared Function GetNeutralModelRangeSuffix(mrsuffix As String) As String
        Dim neutralmrs As String

        If mrsuffix = "WP" Then
            neutralmrs = "FP"
        ElseIf mrsuffix = "AG" Then
            neutralmrs = "AP"
        ElseIf mrsuffix = "CG" Then
            neutralmrs = "CP"
        ElseIf mrsuffix = "PX" Then
            neutralmrs = "RX"
        Else
            neutralmrs = mrsuffix
        End If

        Return neutralmrs
    End Function

    Shared Function GetNeutralModelRangeName(mrname As String) As String
        Dim neutralmrn As String = mrname

        If General.currentunit.ApplicationType = "Condenser" And mrname <> "GCO" Then
            neutralmrn = "Gx" + mrname.Substring(2, 2)
        End If

        Return neutralmrn
    End Function

    Shared Function GetOneBranch(headeralign As String, vtype As String, pressure As Integer, nodistr As Integer) As Boolean
        Dim onebranch As Boolean = True

        If nodistr > 1 Then
            onebranch = False
        Else
            If headeralign = "horizontal" Or (vtype = "P" And pressure > 16) Then
                onebranch = False
            End If
        End If

        Return onebranch
    End Function

    Shared Sub FlangeData(ByRef consys As ConSysData, cbtext As String)
        Dim flangetype, flangeID, flangeerp, necklength As String

        Try
            If General.isProdavit Then
                flangeerp = PCFData.GetValue("ThreadFlangeConnectionFC" + consys.Circnumber.ToString, "ERPCode")

                If flangeerp = "" Then
                    flangetype = cbtext
                    flangeID = ""
                Else
                    flangetype = GetFlangeTypebyERP(flangeerp, consys.HeaderMaterial)
                    flangeID = Database.GetFlangeIDByERP(flangeerp)
                End If
            Else
                flangetype = cbtext
                flangeID = ""
            End If

            If flangeID = "" Then
                flangeID = Database.GetFlangeID(flangetype, consys.OutletNipples.First.Diameter, consys.HeaderMaterial)
            End If

            If General.isProdavit And General.currentunit.ApplicationType = "Condenser" Then
                If flangeerp.Contains(".") And flangeerp <> "BGFPS065.1" And flangeerp <> "BGFPS065.3" Then
                    necklength = "long"
                Else
                    necklength = "short"
                End If
            Else
                necklength = "short"
            End If

            If (Not flangetype.Contains("Loose") Or consys.OutletNipples.First.Diameter < 50) And General.currentunit.ApplicationType = "Condenser" And necklength = "long" Then
                consys.InletNipples.First.Length = Calculation.GetNewNippleLength(consys.InletNipples.First.Diameter, flangetype, consys.InletHeaders.First.Dim_a)
                consys.OutletNipples.First.Length = Calculation.GetNewNippleLength(consys.OutletNipples.First.Diameter, flangetype, consys.OutletHeaders.First.Dim_a)
            ElseIf (General.currentunit.ApplicationType = "Evaporator" Or necklength = "short") And consys.OutletNipples.First.Diameter > 50 And flangetype.Contains("Loose") Then
                'change flangeID for big diameters
                Select Case flangeID
                    Case "0000812286"
                        flangeID = "0000535943"
                    Case "0000812291"
                        flangeID = "0000535944"
                    Case "0000812294"
                        flangeID = "0000535951"
                    Case "0000812297"
                        flangeID = "0000507114"
                    Case "000812297"
                        flangeID = "0000507114"
                    Case "0000812298"
                        flangeID = "0000536001"
                End Select
            End If

            WSM.CheckoutPart(flangeID, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)

            If flangeID = "0000001745" Then
                flangeID = "0000391044"
            End If

            If flangeID = "" Then
                consys.HasFTCon = False
            End If
            consys.FlangeID = flangeID
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function GetFlangeTypebyERP(ERPCode As String, material As String) As String
        Dim suberp, ftypename As String

        If ERPCode.Length > 4 Then
            suberp = ERPCode.Substring(0, 5)
            If suberp = "BGFPS" Then
                ftypename = "Loose / F / 1"
            Else
                If suberp.Contains("V") Then
                    If suberp.Substring(0, 2) = "AZ" Then
                        ftypename = "Thread / S / 1"
                    Else
                        ftypename = "Welded / V / 1"
                    End If
                Else
                    If suberp.Contains("-") Then
                        ftypename = "Welded / F / 4"
                    Else
                        If suberp.Substring(0, 2) = "AZ" Then
                            ftypename = "Thread / S / 1"
                        Else
                            ftypename = "Welded / F / 1"
                        End If
                    End If
                End If
            End If
        ElseIf ERPCode <> "" Then
            If material = "C" Then
                'most likely a threaded connection
                ftypename = "Thread / M / 1"
            Else
                ftypename = "Thread / V / 1"
            End If
        Else
            ftypename = ""
        End If

        Return ftypename
    End Function

    Shared Function GetValveSizeByERP(erpcode As String) As String

    End Function

    Shared Function GetCoords(ByRef coil As CoilData, ByRef circuit As CircuitData, circfile As String) As List(Of Double)()
        Dim dftdoc As SolidEdgeDraft.DraftDocument
        Dim objsheet, mainsheet, objsheet2 As SolidEdgeDraft.Sheet
        Dim stutzencoords, stutzencoords2 As List(Of Double)()
        Dim circframe(), circsize, offset(), circfinpitch() As Double
        Dim changefin As Boolean = False
        Dim circfin, ADpos1 As String

        Try
            dftdoc = SEDraft.OpenDFT(circfile)
            If Not circuit.CustomCirc Then
                'check if mirrored
                If Not General.currentunit.UnitDescription = "VShape" And ((circuit.ConnectionSide = "left" And Not circuit.CircuitType.Contains("Defrost")) Or (circuit.ConnectionSide = "right" And circuit.CircuitType.Contains("Defrost"))) Then
                    'move all objects by frame width in x direction
                    SEDraft.CreateMirroredDFT(dftdoc, False, circuit.ConnectionSide, coil.Alignment, circuit.CustomCirc)
                End If

                If General.currentunit.UnitDescription = "VShape" And circuit.Quantity = 1 And Not circuit.CircuitType.Contains("Subcooler") Then
                    'search correct drawing view - B = right side // A = left side
                    objsheet = If(SEDraft.FindSheetX(circuit.ConnectionSide, circuit.PitchX, circuit.PitchY), SEDraft.FindSheet(dftdoc))
                Else
                    'Find the sheet of the main circuiting view
                    objsheet = SEDraft.FindSheet(dftdoc)
                    If General.currentunit.UnitDescription = "VShape" And circuit.Quantity > 1 Then
                        'check position of first coretube (must be top right)
                        SEDraft.VShapeMirrorDV(objsheet, circuit.ConnectionSide)
                    End If
                End If
            Else
                objsheet = SEDraft.FindSheet(dftdoc)
            End If

            If objsheet IsNot Nothing Then
                'consider blind tubes
                If coil.NoBlindTubes > 0 Then
                    'no of blind tube layers at the beginning of the circuit
                    coil.NoBlindTubeLayers = CircProps.CheckBTLayer(coil.NoBlindTubes, coil.FinnedDepth, {circuit.PitchX, circuit.PitchY}, coil.Alignment)
                End If

                circframe = SEDraft.GetCoilFrame(objsheet, coil.Alignment, General.currentunit.MultiCircuitDesign, coil.Number)
                If circuit.CircuitSize Is Nothing Then
                    Calculation.GetCircsize(circuit, coil.Alignment, circframe, objsheet)
                End If

                offset = CircProps.GetCircOffset(circuit, coil)

                'check, if fin matches the ct arrangement from the circuiting
                If circuit.FinType <> "N" And circuit.FinType <> "M" And circuit.CircuitType <> "Defrost" And Not circuit.CustomCirc Then
                    circfin = CircProps.CheckFin(coil.Alignment, circuit.FinType, objsheet)
                    circfinpitch = GNData.GetFinPitch(coil.Alignment, circfin)
                    If circfinpitch(0) <> circuit.PitchX Or circfinpitch(1) <> circuit.PitchY Then
                        'Get new frame
                        circframe = CircProps.ChangeFrame(circframe, circfinpitch, GNData.GetFinPitch(coil.Alignment, circuit.FinType))
                        Calculation.GetCircsize(circuit, coil.Alignment, circframe, objsheet)
                        circuit.PitchX = circfinpitch(0)
                        circuit.PitchY = circfinpitch(1)
                        changefin = True
                    End If
                End If

                stutzencoords = SEDraft.GetInOutCoords(objsheet, circuit.PitchX, circuit.PitchY, coil.Alignment, circuit.CircuitType, circuit.NoPasses, coil.Number)

                If changefin Then
                    stutzencoords = CircProps.ChangeBowprops(stutzencoords, circfinpitch, GNData.GetFinPitch(coil.Alignment, circuit.FinType), "stutzen")

                    circfinpitch = GNData.GetFinPitch(coil.Alignment, circuit.FinType)
                    circuit.PitchX = circfinpitch(0)
                    circuit.PitchY = circfinpitch(1)
                End If

                If Unit.rotatefin Then
                    If coil.Alignment = "horizontal" Then
                        circsize = circframe(0)
                    Else
                        circsize = circframe(1)
                    End If
                    stutzencoords = Unit.SwitchCoordLists(stutzencoords, circsize, circuit.ConnectionSide)
                    If coil.Alignment = "vertical" Then
                        stutzencoords(0) = CircProps.MoveCoordsbyRotation(coil.FinnedDepth, stutzencoords(0), 0)
                        stutzencoords(1) = CircProps.MoveCoordsbyRotation(coil.FinnedHeight, stutzencoords(1), offset(1))
                        stutzencoords(2) = CircProps.MoveCoordsbyRotation(coil.FinnedDepth, stutzencoords(2), 0)
                        stutzencoords(3) = CircProps.MoveCoordsbyRotation(coil.FinnedHeight, stutzencoords(3), offset(1))
                    End If
                ElseIf circuit.CircuitType.Contains("Defrost") Then
                    stutzencoords = Unit.TransformCoords(stutzencoords, coil.FinnedDepth)
                ElseIf offset.Max > 0 Then
                    For i As Integer = 0 To stutzencoords(0).Count - 1
                        stutzencoords(0)(i) += offset(0)
                        stutzencoords(1)(i) += offset(1)
                        stutzencoords(2)(i) += offset(0)
                        stutzencoords(3)(i) += offset(1)
                    Next
                End If

                If General.currentunit.UnitDescription = "Dual" Then
                    mainsheet = dftdoc.Sheets.Item(1)
                    mainsheet.Activate()
                    General.seapp.DoIdle()

                    ADpos1 = SEDraft.GetADPosition(objsheet)
                    If ADpos1 = "left" Then
                        'add gap + fd to all inlet and outlet coords
                        For i As Integer = 0 To stutzencoords(0).Count - 1
                            stutzencoords(0)(i) = Math.Round(stutzencoords(0)(i) + coil.Gap + coil.FinnedDepth, 3)
                        Next
                        For i As Integer = 0 To stutzencoords(2).Count - 1
                            stutzencoords(2)(i) = Math.Round(stutzencoords(2)(i) + coil.Gap + coil.FinnedDepth, 3)
                        Next
                    End If

                    objsheet2 = SEDraft.FindSheet2(circfile, objsheet.Name)
                    If objsheet2 IsNot Nothing Then
                        Dim ADpos2 As String = SEDraft.GetADPosition(objsheet2)
                        If ADpos1 = ADpos2 Then
                            stutzencoords2 = Calculation.MirrorCoordsFromList(stutzencoords, coil.FinnedDepth, New List(Of Integer) From {0, 2})
                        Else
                            stutzencoords2 = SEDraft.GetInOutCoords(objsheet2, circuit.PitchX, circuit.PitchY, coil.Alignment, circuit.CircuitType, circuit.NoPasses, coil.Number)

                            If ADpos1 = "right" Then
                                'add gap + fd to all inlet and outlet coords
                                For i As Integer = 0 To stutzencoords2(0).Count - 1
                                    stutzencoords2(0)(i) = Math.Round(stutzencoords2(0)(i) + coil.Gap + coil.FinnedDepth, 3)
                                Next
                                For i As Integer = 0 To stutzencoords2(2).Count - 1
                                    stutzencoords2(2)(i) = Math.Round(stutzencoords2(2)(i) + coil.Gap + coil.FinnedDepth, 3)
                                Next
                            End If
                        End If
                    End If
                End If

                'resolve the pattern for quantity > 1 into definite positions?
                If circuit.Quantity > 1 And General.currentunit.UnitDescription <> "Dual" Then
                    For i As Integer = 1 To circuit.Quantity - 1
                        If coil.Alignment = "horizontal" Then
                            For j As Integer = 0 To (circuit.NoDistributions / circuit.Quantity) - 1
                                stutzencoords(0).Add(stutzencoords(0)(j) + i * circuit.CircuitSize(0) / circuit.Quantity)
                                stutzencoords(1).Add(stutzencoords(1)(j))
                                stutzencoords(2).Add(stutzencoords(2)(j) + i * circuit.CircuitSize(0) / circuit.Quantity)
                                stutzencoords(3).Add(stutzencoords(3)(j))
                            Next
                        Else
                            For j As Integer = 0 To (circuit.NoDistributions / circuit.Quantity) - 1
                                stutzencoords(0).Add(stutzencoords(0)(j))
                                stutzencoords(2).Add(stutzencoords(2)(j))
                                If stutzencoords(1)(0) > circuit.CircuitSize(1) / circuit.Quantity Then
                                    stutzencoords(1).Add(stutzencoords(1)(j) - i * circuit.CircuitSize(1) / circuit.Quantity)
                                    stutzencoords(3).Add(stutzencoords(3)(j) - i * circuit.CircuitSize(1) / circuit.Quantity)
                                Else
                                    stutzencoords(1).Add(stutzencoords(1)(j) + i * circuit.CircuitSize(1) / circuit.Quantity)
                                    stutzencoords(3).Add(stutzencoords(3)(j) + i * circuit.CircuitSize(1) / circuit.Quantity)
                                End If
                            Next
                        End If
                    Next
                End If

                General.seapp.Documents.CloseDocument(circfile, SaveChanges:=False, DoIdle:=True)
            End If


        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return stutzencoords
    End Function

    Shared Sub GetStutzen(ByRef consys As ConSysData, coil As CoilData, circuit As CircuitData, headerpairno As Integer)
        Dim overridefig As Boolean = False
        Dim abvlist, uniqueabsabvlist As List(Of Double)
        Dim invalues, outvalues As New Dictionary(Of Double, String)       'abv - ID
        Dim abvangle As New Dictionary(Of Double, Double)       'abv - angle
        Dim stutzenkeys, templateids, IDList As New List(Of String)
        Dim anglelist As New List(Of Double)
        Dim figurelist As New List(Of Integer)

        'no more iterations, only one time calculation with given input - consider previous angles if figure 5
        Try
            If General.currentunit.ModelRangeName = "GACV" Then
                'override certain headerdata for GACV horizontal
                overridefig = ControlCSDataGACV(consys, {consys.InletHeaders.First.Xlist, consys.InletHeaders.First.Ylist, consys.OutletHeaders.First.Xlist, consys.OutletHeaders.First.Ylist}, circuit, coil.FinnedHeight, coil.FinnedDepth)
            End If

            'get header origin
            If consys.InletHeaders.First.Tube.Quantity > 0 Or circuit.NoDistributions = 1 Then
                If headerpairno = 1 Then
                    consys.InletHeaders.First.Origin = GetHeaderOrigin(consys.InletHeaders.First, coil, circuit, consys)
                Else
                    consys.InletHeaders.Last.Origin = GetHeaderOrigin(consys.InletHeaders.Last, coil, circuit, consys)
                End If
                'each stutzencoord has its own ABV → same length
                abvlist = GetABVList(consys.InletHeaders.First, consys.HeaderAlignment)

                'sort by distance? → how to handle +/-25? → have to be own entry as the holes sit on different planes in the header
                uniqueabsabvlist = SortABV(abvlist)
                uniqueabsabvlist.Sort()

                If uniqueabsabvlist.Max <= 0 Then
                    uniqueabsabvlist.Reverse()
                End If

                For i As Integer = 0 To uniqueabsabvlist.Count - 1
                    Dim sID As String
                    Dim angle As Double
                    Dim figure As Integer = GNData.DefineFigure(consys.HeaderAlignment, circuit, consys.InletHeaders.First, uniqueabsabvlist, uniqueabsabvlist(i), overridefig, coil.NoRows)
                    Dim sentry() As String = GNData.GetStutzenID(uniqueabsabvlist(i), consys.InletHeaders.First, circuit, consys, figure, 0, abvangle)
                    IDList.Add(sentry(0))
                    anglelist.Add(sentry(1))
                    sID = sentry(0)
                    angle = sentry(1)
                    abvangle.Add(uniqueabsabvlist(i), angle)
                    invalues.Add(uniqueabsabvlist(i), sID)
                    figurelist.Add(figure)
                Next

                'for every template, create a model
                CheckForTemplate(IDList, anglelist, stutzenkeys, circuit, consys.InletHeaders.First, uniqueabsabvlist)

                For i As Integer = 0 To consys.InletHeaders.First.Xlist.Count - 1
                    Dim specialtag As String = ""
                    If abvlist(i) = -55.9 Then
                        specialtag = "s2star"
                    End If
                    For j As Integer = 0 To uniqueabsabvlist.Count - 1
                        If abvlist(i) = uniqueabsabvlist(j) Then
                            If headerpairno = 1 Then
                                consys.InletHeaders.First.StutzenDatalist.Add(New StutzenData With {.ID = IDList(j), .Angle = anglelist(j), .ABV = abvlist(i), .SpecialTag = specialtag, .Figure = figurelist(j),
                                .XPos = consys.InletHeaders.First.Xlist(i), .YPos = -circuit.CoreTubeOverhang + 5, .ZPos = consys.InletHeaders.First.Ylist(i)})
                            Else
                                consys.InletHeaders.Last.StutzenDatalist.Add(New StutzenData With {.ID = IDList(j), .Angle = anglelist(j), .ABV = abvlist(i), .SpecialTag = specialtag, .Figure = figurelist(j),
                                .XPos = consys.InletHeaders.Last.Xlist(i), .YPos = -circuit.CoreTubeOverhang + 5, .ZPos = consys.InletHeaders.Last.Ylist(i)})
                            End If
                        End If
                    Next
                Next

                IDList.Clear()
                anglelist.Clear()
                abvangle.Clear()
                figurelist.Clear()
            End If

            If headerpairno = 1 Then
                consys.OutletHeaders.First.Origin = GetHeaderOrigin(consys.OutletHeaders.First, coil, circuit, consys)
            Else
                consys.OutletHeaders.Last.Origin = GetHeaderOrigin(consys.OutletHeaders.Last, coil, circuit, consys)
            End If
            abvlist = GetABVList(consys.OutletHeaders.First, consys.HeaderAlignment)

            'sort by distance? → how to handle +/-25? → use only unique values and assign by abv
            uniqueabsabvlist = SortABV(abvlist)
            uniqueabsabvlist.Sort()

            If uniqueabsabvlist.Max <= 0 Then
                uniqueabsabvlist.Reverse()
            End If

            For i As Integer = 0 To uniqueabsabvlist.Count - 1
                Dim sID As String
                Dim angle As Double
                Dim figure As Integer = GNData.DefineFigure(consys.HeaderAlignment, circuit, consys.OutletHeaders.First, uniqueabsabvlist, uniqueabsabvlist(i), overridefig, coil.NoRows)
                Dim sentry() As String = GNData.GetStutzenID(uniqueabsabvlist(i), consys.OutletHeaders.First, circuit, consys, figure, 0, abvangle)
                sID = sentry(0)
                angle = sentry(1)
                IDList.Add(sentry(0))
                anglelist.Add(sentry(1))
                abvangle.Add(uniqueabsabvlist(i), angle)
                outvalues.Add(uniqueabsabvlist(i), sID)
                figurelist.Add(figure)
            Next

            'for every template, create a model
            CheckForTemplate(IDList, anglelist, stutzenkeys, circuit, consys.OutletHeaders.First, uniqueabsabvlist)

            For i As Integer = 0 To consys.OutletHeaders.First.Xlist.Count - 1
                Dim specialtag As String = ""
                If abvlist(i) = -55.9 Then
                    specialtag = "s2star"
                End If
                For j As Integer = 0 To uniqueabsabvlist.Count - 1
                    If abvlist(i) = uniqueabsabvlist(j) Then
                        If headerpairno = 1 Then
                            consys.OutletHeaders.First.StutzenDatalist.Add(New StutzenData With {.ID = IDList(j), .Angle = anglelist(j), .ABV = abvlist(i), .SpecialTag = specialtag, .Figure = figurelist(j),
                        .XPos = consys.OutletHeaders.First.Xlist(i), .YPos = -circuit.CoreTubeOverhang + 5, .ZPos = consys.OutletHeaders.First.Ylist(i)})
                        Else
                            consys.OutletHeaders.Last.StutzenDatalist.Add(New StutzenData With {.ID = IDList(j), .Angle = anglelist(j), .ABV = abvlist(i), .SpecialTag = specialtag, .Figure = figurelist(j),
                        .XPos = consys.OutletHeaders.Last.Xlist(i), .YPos = -circuit.CoreTubeOverhang + 5, .ZPos = consys.OutletHeaders.Last.Ylist(i)})
                        End If
                    End If
                Next
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub CheckForTemplate(ByRef IDList As List(Of String), ByRef anglelist As List(Of Double), ByRef stutzenkeys As List(Of String), circuit As CircuitData, header As HeaderData, uniqueabsabvlist As List(Of Double))

        For i As Integer = 0 To IDList.Count - 1
            Dim tempfig As Integer = GNData.CheckTemplate(IDList(i))
            If tempfig > 0 Then
                Dim minangle As Double = Calculation.MinimumAngle(uniqueabsabvlist, uniqueabsabvlist(i), tempfig)
                WSM.CheckoutCircs(IDList(i), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                General.WaitForFile(General.currentjob.Workspace, IDList(i), "par", 100)
                Dim stutzenkey As String = CircProps.CreateStutzenkey(tempfig, header, circuit.CoreTubeOverhang, Math.Abs(uniqueabsabvlist(i)), minangle, header.Tube.IsBrine)
                If stutzenkey <> "0\0\0" Then
                    stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                    Dim createkey As Boolean
                    If Unit.stutzendict.Count = 0 Then
                        createkey = True
                    Else
                        For Each s In Unit.stutzendict
                            If s.Key = stutzenkey Then
                                createkey = False
                                IDList(i) = s.Value
                                anglelist(i) = GetAnglefromKey(stutzenkey)
                                Exit For
                            Else
                                createkey = True
                            End If
                        Next
                    End If
                    If createkey Then
                        stutzenkeys.Add(stutzenkey)
                        Dim templateid = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, IDList(i), circuit.CoreTube.Material, circuit.FinType, tempfig, circuit.Pressure, uniqueabsabvlist(i))
                        Unit.stutzendict.Add(stutzenkey, templateid)
                        IDList(i) = templateid
                        If tempfig = 5 Then
                            anglelist(i) = GetAnglefromKey(stutzenkey)
                        Else
                            anglelist(i) = 0
                        End If
                    End If
                End If
            End If
        Next

    End Sub

    Shared Function ControlCSDataGACV(ByRef consys As ConSysData, stutzencoords() As List(Of Double), circuit As CircuitData, finnedheight As Double, finneddepth As Double) As Boolean
        Dim uniqueinlets As List(Of Double)
        Dim overridefig As Boolean

        If consys.HeaderAlignment = "vertical" Then
            If consys.InletHeaders.First.Tube.Quantity > 0 Then
                If consys.InletHeaders(0).Tube.Diameter > 60 And consys.HeaderMaterial = "C" And circuit.FinType = "F" Then
                    uniqueinlets = General.GetUniqueValues(stutzencoords(0))
                    If uniqueinlets.Count < 3 Then
                        overridefig = True
                        If circuit.ConnectionSide = "right" Then
                            consys.InletHeaders(0).Displacehor -= 25
                        Else
                            consys.InletHeaders(0).Displacehor += 25
                        End If
                    End If

                End If
                'set vertical displacement to 0, if highest inlet is not in top row
                If consys.InletHeaders(0).Displacever <> 0 Then
                    If Math.Round(stutzencoords(1).Max + circuit.PitchY / 2, 2) < finnedheight Then
                        consys.InletHeaders(0).Displacever = 0
                    End If
                End If
                If consys.OutletHeaders.First.Tube.Diameter = 26.9 And General.currentunit.ModelRangeSuffix = "CP" Then
                    uniqueinlets = General.GetUniqueValues(stutzencoords(2))
                    If uniqueinlets.Count = 2 Then
                        consys.OutletHeaders.First.Overhangbottom += 25
                    End If
                End If
            Else
                'DX Evaporator
                If circuit.FinType = "N" And circuit.NoPasses > 36 And consys.OutletHeaders.First.Tube.Materialcodeletter = "C" Then
                    If consys.OutletHeaders.First.Tube.Diameter = 54 Then
                        consys.OutletHeaders.First.Overhangbottom = 18.5
                        consys.OutletHeaders.First.Overhangtop = 18.5
                    ElseIf consys.OutletHeaders.First.Tube.Diameter = 64 Then
                        consys.OutletHeaders.First.Overhangbottom = 20
                        consys.OutletHeaders.First.Overhangtop = 20
                        consys.OutletHeaders.First.Displacever = -20
                    End If
                End If

            End If
            If consys.OutletHeaders(0).Displacever <> 0 Then
                If Math.Round(stutzencoords(3).Max + circuit.PitchY / 2, 2) < finnedheight Then
                    consys.OutletHeaders(0).Displacever = 0
                End If
            End If
        Else
            uniqueinlets = General.GetUniqueValues(stutzencoords(3))
            If circuit.ConnectionSide = "left" Then
                consys.OutletHeaders(0).Overhangbottom = consys.OutletHeaders(0).Overhangtop
                consys.InletHeaders(0).Overhangbottom = consys.InletHeaders(0).Overhangtop
                consys.OutletHeaders(0).Overhangtop = finneddepth - stutzencoords(2).Max + General.currentunit.TubeSheet.Dim_d + 62
                consys.InletHeaders(0).Overhangtop = finneddepth - stutzencoords(0).Max + General.currentunit.TubeSheet.Dim_d + 62
            Else
                consys.OutletHeaders(0).Overhangtop = consys.OutletHeaders(0).Overhangbottom
                consys.InletHeaders(0).Overhangtop = consys.InletHeaders(0).Overhangbottom
                consys.OutletHeaders(0).Overhangbottom = General.currentunit.TubeSheet.Dim_d + 62 + stutzencoords(2).Min
                consys.InletHeaders(0).Overhangbottom = General.currentunit.TubeSheet.Dim_d + 62 + stutzencoords(0).Min
            End If

            If consys.InletHeaders(0).Tube.Diameter = 26.9 And uniqueinlets.Count > 1 Then
                If circuit.ConnectionSide = "left" Then
                    consys.InletHeaders(0).Overhangbottom = 41
                Else
                    consys.InletHeaders(0).Overhangtop = 41
                End If
            End If

            If stutzencoords(0).Count > 1 And uniqueinlets.Count < 3 And consys.HeaderMaterial <> "C" Then
                OverwriteAPPosition(consys, uniqueinlets.Count, circuit.Orbitalwelding)
            End If

            If uniqueinlets.Count = 2 And General.currentunit.ModelRangeSuffix.Substring(0, 1) = "A" And consys.OutletHeaders(0).Tube.Diameter > 60 Then
                consys.OutletHeaders(0).Displacever = -100
            End If
        End If

        Return overridefig
    End Function

    Shared Function ControlDXGACV(ByRef header As HeaderData, circuit As CircuitData) As Boolean
        Dim changedoh As Boolean = False

        'VA headers, including 80b with CU coretubes
        'only for multiple headers → should exclude CP

        Select Case circuit.FinType
            Case "F"
                'F fin → 32b RX VA header
                Select Case header.Tube.Diameter
                    Case 60.3
                        header.Displacever = -48
                    Case 76.1
                        header.Displacever = -45
                    Case 88.9
                        header.Displacever = -42
                End Select
            Case "E"
                Select Case circuit.Pressure
                    Case 54
                        If header.Tube.Diameter = 60.3 Then
                            'E-54-CX-VA // 60.3
                            header.Displacever = -50
                            changedoh = True
                        End If
                    Case 80
                        'E-80-CX-CU // 48.3 , 60.3
                        'E-80-CX-VA // 48.3 , 60.3
                        If header.Tube.Diameter > 48 Then
                            header.Displacever = -50
                            changedoh = True
                        End If
                End Select
            Case Else
                Select Case circuit.Pressure
                    Case 54
                        'N-54-CX-VA // 42.4 , 48.3
                        If header.Tube.Diameter = 42.4 Then
                            header.Overhangtop = 20
                            header.Displacever = -22
                            changedoh = True
                        ElseIf header.Tube.Diameter = 48.3 Then
                            header.Displacever = -19.8
                            changedoh = True
                        End If
                    Case 80
                        If circuit.CoreTube.Materialcodeletter = "C" Then
                            'N-80-CX-CU // 33.7 , 42.4
                            If header.Tube.Diameter = 33.7 Then
                                header.Displacever = -20
                                changedoh = True
                                'for first column -23.6
                            ElseIf header.Tube.Diameter = 42.4 Then
                                header.Displacever = -20
                                changedoh = True
                                'for first colum -22
                            End If
                        Else
                            'N-80-CX-VA // 42.4
                            If header.Tube.Diameter = 42.4 Then
                                header.Displacever = -22
                                changedoh = True
                            End If
                        End If
                End Select
        End Select

        Return changedoh
    End Function

    Shared Sub OverwriteAPPosition(ByRef consys As ConSysData, inletcount As Integer, orbital As Boolean)

        If inletcount = 1 Then
            consys.InletHeaders(0).Dim_a = 110
            consys.OutletHeaders(0).Dim_a = 110
        Else
            consys.InletHeaders(0).Dim_a = 130
            consys.OutletHeaders(0).Dim_a = 130
        End If
        consys.OutletHeaders(0).Displacever = -50

        If orbital Then
            With consys
                .InletHeaders.First.Dim_a += 20
                .OutletHeaders.First.Dim_a += 20
            End With
        End If

    End Sub

    Shared Function GetHeaderOrigin(header As HeaderData, coil As CoilData, circuit As CircuitData, consys As ConSysData) As Double()
        Dim xorigin, yorigin, zorigin, origin() As Double

        yorigin = header.Dim_a + header.Tube.Diameter / 2

        If (circuit.NoPasses = 1 Or circuit.NoPasses = 3) And header.OddLocation = "back" Then
            yorigin = -(coil.FinnedLength + General.currentunit.TubeSheet.Thickness * 2 + yorigin)
        End If

        'get the origin position depending of the alignment
        If General.currentunit.UnitDescription = "VShape" Then
            'always vertical alignment
            xorigin = GetXZOrigin(header.Xlist, General.currentunit.ModelRangeName, header.Tube.HeaderType, consys.HeaderAlignment, coil.FinnedDepth, coil.FinnedHeight, circuit.ConnectionSide)
            xorigin += header.Displacehor
            zorigin = header.Ylist.Min - header.Overhangbottom
            If header.Tube.HeaderType = "inlet" Then
                zorigin += header.Displacever
            End If
        ElseIf General.currentunit.UnitDescription = "Dual" Then
            If header.Xlist.Max < coil.FinnedDepth Then
                xorigin = header.Xlist.Max + Math.Abs(header.Displacehor)
            Else
                xorigin = header.Xlist.Min - Math.Abs(header.Displacehor)
            End If
            zorigin = header.Ylist.Min - header.Overhangbottom
        Else
            If consys.HeaderAlignment = "vertical" Then
                xorigin = GetXZOrigin(header.Xlist, General.currentunit.ModelRangeName, header.Tube.HeaderType, consys.HeaderAlignment, coil.FinnedDepth, coil.FinnedHeight, circuit.ConnectionSide)
                xorigin += header.Displacehor
                zorigin = header.Ylist.Min - header.Overhangbottom
                'shorten the inlet header by 25mm because of collision with heating plate
                If General.currentunit.ModelRangeName = "GACV" And header.Tube.HeaderType = "inlet" And header.Tube.Diameter > 60 And circuit.FinType = "F" And header.Ylist.Min = 12.5 Then
                    zorigin += 25
                End If
            Else
                If consys.HeaderAlignment = coil.Alignment Then
                    zorigin = GetXZOrigin(header.Ylist, General.currentunit.ModelRangeName, header.Tube.HeaderType, consys.HeaderAlignment, coil.FinnedDepth, coil.FinnedHeight, circuit.ConnectionSide)
                Else
                    zorigin = GetXZOrigin(header.Ylist, General.currentunit.ModelRangeName, header.Tube.HeaderType, consys.HeaderAlignment, coil.FinnedHeight, coil.FinnedDepth, circuit.ConnectionSide)
                End If
                zorigin += header.Displacever
                xorigin = header.Xlist.Min - header.Overhangbottom + header.Displacehor
            End If
        End If

        origin = {xorigin, yorigin, zorigin}

        Return origin
    End Function

    Shared Function GetXZOrigin(xzlist As List(Of Double), modelrange As String, headertype As String, alignment As String, finneddepth As Double, finnedheight As Double, conside As String) As Double
        Dim delta As Double
        Dim xzorigin As Double

        delta = Math.Abs(finneddepth - xzlist.Max)
        If delta < finneddepth / 2 Then
            xzorigin = xzlist.Max
        ElseIf delta = finneddepth / 2 Then
            delta = Math.Abs(finneddepth - xzlist.Min)
            If delta < finneddepth / 2 Then
                xzorigin = xzlist.Max
            Else
                xzorigin = xzlist.Min
            End If
        Else
            xzorigin = xzlist.Min
        End If

        If conside = "right" Then
            If modelrange = "GACV" And headertype = "inlet" And alignment = "vertical" Then
                xzorigin = xzlist.Min
            End If
        Else
            If modelrange = "GACV" And headertype = "inlet" And alignment = "vertical" Then
                xzorigin = xzlist.Max
            End If
        End If

        Return xzorigin
    End Function

    Shared Function GetABVList(header As HeaderData, alignment As String) As List(Of Double)
        Dim abvlist As New List(Of Double)
        Dim valuelist As List(Of Double)
        Dim originpoint, abv As Double

        If alignment = "vertical" Then
            valuelist = header.Xlist
            originpoint = header.Origin(0)
        Else
            valuelist = header.Ylist
            originpoint = header.Origin(2)
        End If

        'get the abv list
        For i As Integer = 0 To valuelist.Count - 1
            abv = Math.Round(originpoint - valuelist(i), 3)
            If alignment = "horizontal" And General.currentunit.ModelRangeName = "GACV" And Not header.Tube.IsBrine Then
                If Math.Abs(abv) = 50 Then
                    If General.currentunit.ModelRangeSuffix.Substring(0, 1) = "A" Then
                        'outlet and 3 rows, outlet and 2 rows and diameter < 60
                        If header.Tube.HeaderType = "outlet" And General.GetUniqueValues(valuelist).Count > 1 Then
                            abv = -55.9
                            If General.GetUniqueValues(valuelist).Count = 2 And header.Tube.Diameter > 60 And header.Displacever = -100 Then
                                abv = -50
                            End If
                        ElseIf header.Tube.HeaderType = "inlet" And header.Tube.Diameter < 60 Then
                            abv = -55.9
                        End If
                    Else
                        'can only be CP
                        If header.Tube.Materialcodeletter = "C" Then
                            If abv = -16 Then
                                abv = -29.7
                            End If
                        Else
                            If (header.Tube.HeaderType = "outlet" Or header.Tube.Diameter < 60) And General.GetUniqueValues(valuelist).Count > 1 Then
                                abv = -55.9
                            End If
                        End If
                    End If
                ElseIf header.Tube.HeaderType = "inlet" And header.Tube.Diameter = 60.3 Then
                    'tangential
                    abv = 0
                End If
            End If
            abvlist.Add(abv)
        Next

        Return abvlist
    End Function

    Shared Function SortABV(abvlist As List(Of Double)) As List(Of Double)
        Dim templist As New List(Of Double)

        For i As Integer = 0 To abvlist.Count - 1
            templist.Add(abvlist(i))
        Next

        Return General.GetUniqueValues(templist)
    End Function

    Shared Function GetAnglefromKey(stutzenkey As String) As Double
        Dim values() As String
        Dim angle As Double

        values = stutzenkey.Split({"\"}, 0)
        angle = CDbl(values(3))

        Return angle
    End Function

    Shared Sub GACVSpecialInlet(ByRef header As HeaderData, circuit As CircuitData, consys As ConSysData)
        Dim abvs, displacever, gap As Double
        Dim sdata(), stutzenkey As String

        For i As Integer = 0 To header.Ylist.Count - 1
            Dim specialtag As String = ""
            Dim loopcount, mp, tempfigure As Integer
            loopcount = 0
            gap = 0
            If header.StutzenDatalist(i).ABV < 0 Then
                mp = -1
            Else
                mp = 1
            End If
            If header.Ylist(i) = header.Ylist.Max Then
                loopcount = 2
                specialtag = "s1star"
                If circuit.ConnectionSide = "right" Then
                    abvs = Math.Round(Math.Abs(header.Displacehor) + header.Xlist(i) - header.Xlist.Min, 3)
                Else
                    abvs = Math.Round(Math.Abs(header.Displacehor) + Math.Abs(header.Xlist(i) - header.Xlist.Max), 3)
                End If
                'if abvs < pitchx then only 2 rows of inlet → change to fig 4, for big F FP VA headers, displace > pitchx
                If abvs < circuit.PitchX Or (abvs = Math.Abs(header.Displacehor) And General.currentunit.ModelRangeSuffix <> "CP" And (circuit.FinType <> "F" Or header.Tube.Materialcodeletter <> "C")) Then
                    loopcount = 1
                    'diagonal
                    abvs = Math.Round(Math.Sqrt(header.Displacehor ^ 2 + header.Displacever ^ 2), 3)
                End If
                If header.Tube.Materialcodeletter = "V" And circuit.Pressure < 17 And circuit.FinType = "F" Then
                    gap = 25
                    specialtag = "s1star*"
                End If
            ElseIf header.Ylist(i) = circuit.PitchY / 2 And circuit.FinType <> "N" And circuit.FinType <> "M" Then
                If (header.Xlist(i) = header.Xlist.Max And circuit.ConnectionSide = "right") Or (circuit.ConnectionSide = "left" And header.Xlist(i) = header.Xlist.Min) Then
                    'last row → figure 45 
                    If header.Tube.Diameter > 55 Then
                        loopcount = 3
                        specialtag = "s4star"
                        'horizontal but identical with figure 5 of that row
                        If circuit.ConnectionSide = "right" Then
                            abvs = Math.Round(Math.Abs(header.Displacehor) + header.Xlist(i) - header.Xlist.Min, 3)
                        Else
                            abvs = Math.Round(Math.Abs(header.Displacehor) + Math.Abs(header.Xlist(i) - header.Xlist.Max), 3)
                        End If
                    End If
                Else
                    'first row → figure 4
                    loopcount = 1
                    specialtag = "s6star"
                    'diagonal
                    abvs = Math.Round(Math.Sqrt(header.Displacehor ^ 2 + 25 ^ 2), 3)
                End If
            End If

            If loopcount > 0 Then
                tempfigure = GNData.GetCurrentFigure(abvs, loopcount)
                If tempfigure = 45 Then
                    If header.Tube.Diameter > 60 And header.Ylist(i) = circuit.PitchY / 2 And circuit.FinType = "F" Then
                        displacever = 25
                        header.StutzenDatalist(i).HoleOffset = 0
                    Else
                        displacever = header.Displacever
                    End If
                End If

                If header.Ylist(i) = header.Ylist.Max Then
                    header.StutzenDatalist(i).HoleOffset = header.Tube.Length - header.Overhangbottom - header.Overhangtop
                End If
                sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, displacever, New Dictionary(Of Double, Double), gap)
                header.StutzenDatalist(i).ID = sdata(0)
                'if a template is selected, then it needs to be adjusted first and the correct angle should be added to the list
                If GNData.CheckIfTemplate(sdata(0), "Stutzen") Then
                    'create key
                    WSM.CheckoutPart(sdata(0), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                    General.WaitForFile(General.currentjob.Workspace, sdata(0), "par", 100)
                    stutzenkey = CircProps.CreateStutzenkey(tempfigure, header, circuit.CoreTubeOverhang, Math.Abs(abvs), 0, False)
                    stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                    If skeys.IndexOf(stutzenkey) = -1 Then
                        header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, tempfigure, circuit.Pressure, abvs)
                        newStutzen.Add(header.StutzenDatalist(i).ID)
                        newAngles.Add(sdata(1))
                        skeys.Add(stutzenkey)
                    Else
                        header.StutzenDatalist(i).ID = newStutzen(skeys.IndexOf(stutzenkey))
                        sdata(1) = newAngles(skeys.IndexOf(stutzenkey))
                    End If
                End If
                header.StutzenDatalist(i).Angle = sdata(1)
                header.StutzenDatalist(i).ABV = abvs * mp
                If specialtag <> "" Then
                    header.StutzenDatalist(i).SpecialTag = specialtag
                End If
                header.StutzenDatalist(i).Figure = tempfigure
            End If
        Next

    End Sub

    Shared Sub GACVSpecialInletN(ByRef header As HeaderData, circuit As CircuitData, consys As ConSysData)
        Dim sdata(), stutzenkey As String

        'only N AP and CP // so far only top row affected
        For i As Integer = 0 To header.Ylist.Count - 1
            Dim tempfigure As Integer
            Dim abvs As Double
            Dim specialtag As String

            If header.Ylist(i) = header.Ylist.Max Then
                If Math.Abs(header.StutzenDatalist(i).ABV) = 19.2 Or Math.Abs(header.StutzenDatalist(i).ABV) = 75 Then
                    tempfigure = 4
                    specialtag = "s1star"
                    abvs = Math.Round(Math.Sqrt(header.Displacehor ^ 2 + header.Displacever ^ 2), 3)
                Else
                    tempfigure = 45
                    specialtag = "s4star"
                    abvs = Math.Abs(header.StutzenDatalist(i).ABV)
                End If

                sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, header.Displacever, New Dictionary(Of Double, Double))
                header.StutzenDatalist(i).ID = sdata(0)

                If GNData.CheckIfTemplate(sdata(0), "Stutzen") Then
                    'create key
                    WSM.CheckoutPart(sdata(0), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                    General.WaitForFile(General.currentjob.Workspace, sdata(0), "par", 100)
                    stutzenkey = CircProps.CreateStutzenkey(tempfigure, header, circuit.CoreTubeOverhang, Math.Abs(abvs), 0, False)
                    stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                    If skeys.IndexOf(stutzenkey) = -1 Then
                        header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, tempfigure, circuit.Pressure, abvs)
                        newStutzen.Add(header.StutzenDatalist(i).ID)
                        newAngles.Add(sdata(1))
                        skeys.Add(stutzenkey)
                    Else
                        header.StutzenDatalist(i).ID = newStutzen(skeys.IndexOf(stutzenkey))
                        sdata(1) = newAngles(skeys.IndexOf(stutzenkey))
                    End If
                End If
                header.StutzenDatalist(i).Angle = sdata(1)
                header.StutzenDatalist(i).SpecialTag = specialtag
                header.StutzenDatalist(i).Figure = tempfigure
            End If
        Next

    End Sub

    Shared Sub GACVSpecialOutlet(ByRef header As HeaderData, circuit As CircuitData, consys As ConSysData)
        Dim abvs As Double
        Dim sdata(), stutzenkey As String

        For i As Integer = 0 To header.Ylist.Count - 1
            Dim specialtag As String = ""
            Dim loopcount, tempfigure, mp As Integer
            If header.StutzenDatalist(i).ABV < 0 Then
                mp = -1
            Else
                mp = 1
            End If
            If header.Ylist(i) = header.Ylist.Max Then
                specialtag = "s1star"
                If (header.Xlist(i) = header.Xlist.Min And circuit.ConnectionSide = "right") Or (header.Xlist(i) = header.Xlist.Max And circuit.ConnectionSide = "left") Then
                    loopcount = 1
                    abvs = Math.Round(Math.Sqrt(header.Displacehor ^ 2 + header.Displacever ^ 2), 3)
                Else
                    loopcount = 3
                    If circuit.ConnectionSide = "right" Then
                        abvs = Math.Round(Math.Abs(header.Displacehor) + header.Xlist(i) - header.Xlist.Min, 3)
                    Else
                        abvs = Math.Round(Math.Abs(header.Displacehor) + Math.Abs(header.Xlist(i) - header.Xlist.Max), 3)
                    End If
                End If
                tempfigure = GNData.GetCurrentFigure(abvs, loopcount)
                sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, header.Displacever, New Dictionary(Of Double, Double))
                header.StutzenDatalist(i).ID = sdata(0)
                'if a template is selected, then it needs to be adjusted first and the correct angle should be added to the list
                If GNData.CheckIfTemplate(sdata(0), "Stutzen") Then
                    'create key
                    WSM.CheckoutPart(sdata(0), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                    General.WaitForFile(General.currentjob.Workspace, sdata(0), "par", 100)
                    stutzenkey = CircProps.CreateStutzenkey(tempfigure, header, circuit.CoreTubeOverhang, Math.Abs(abvs), 0, False)
                    stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                    If skeys.IndexOf(stutzenkey) = -1 Then
                        header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, tempfigure, circuit.Pressure, abvs)
                        newStutzen.Add(header.StutzenDatalist(i).ID)
                        newAngles.Add(sdata(1))
                        skeys.Add(stutzenkey)
                    Else
                        header.StutzenDatalist(i).ID = newStutzen(skeys.IndexOf(stutzenkey))
                        sdata(1) = newAngles(skeys.IndexOf(stutzenkey))
                    End If
                End If
                header.StutzenDatalist(i).Angle = sdata(1)
                header.StutzenDatalist(i).ABV = abvs * mp
            End If
            If specialtag <> "" Then
                header.StutzenDatalist(i).SpecialTag = specialtag
                header.StutzenDatalist(i).Figure = tempfigure
            End If
        Next

    End Sub

    Shared Sub GACVSpecialOutletN(ByRef header As HeaderData, circuit As CircuitData, consys As ConSysData)
        Dim sdata(), stutzenkey As String

        For i As Integer = 0 To header.Ylist.Count - 1
            Dim tempfigure As Integer = 0
            Dim abvs, displacever As Double
            Dim specialtag As String = ""

            If header.Ylist(i) = header.Ylist.Max Then
                If Math.Abs(header.StutzenDatalist(i).ABV) < 50 Then
                    tempfigure = 4
                    specialtag = "s1star"
                    displacever = header.Displacever
                    If consys.SpecialCX And circuit.CoreTube.Materialcodeletter = "C" Then
                        If header.Tube.Diameter = 33.7 Then
                            abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 23.6 ^ 2), 3)
                        Else
                            abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 22 ^ 2), 3)
                        End If
                    Else
                        abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + header.Displacever ^ 2), 3)
                    End If
                Else
                    tempfigure = 45
                    If consys.VType = "P" And header.Tube.Diameter = 88.9 And circuit.Pressure = 32 Then
                        specialtag = "s4starAP"
                        displacever = 15
                    Else
                        specialtag = "s4star"
                        displacever = header.Displacever
                    End If
                    abvs = Math.Abs(header.StutzenDatalist(i).ABV)
                End If
                sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, displacever, New Dictionary(Of Double, Double))
            ElseIf header.Ylist(i) = header.Ylist.Max - 50 And General.currentunit.ModelRangeSuffix = "CP" And header.Tube.Diameter = 88.9 Then
                If Math.Abs(header.StutzenDatalist(i).ABV) < 50 Then
                    tempfigure = 4
                    specialtag = "s1star"
                    abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 15 ^ 2), 3)
                Else
                    tempfigure = 45
                    specialtag = "s4star"
                    abvs = Math.Abs(header.StutzenDatalist(i).ABV)
                End If
                sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 15, New Dictionary(Of Double, Double))

            End If

            If tempfigure > 0 Then
                header.StutzenDatalist(i).ID = sdata(0)

                If GNData.CheckIfTemplate(sdata(0), "Stutzen") Then
                    'create key
                    WSM.CheckoutPart(sdata(0), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                    General.WaitForFile(General.currentjob.Workspace, sdata(0), "par", 100)
                    stutzenkey = CircProps.CreateStutzenkey(tempfigure, header, circuit.CoreTubeOverhang, Math.Abs(abvs), 0, False)
                    stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                    If skeys.IndexOf(stutzenkey) = -1 Then
                        header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, tempfigure, circuit.Pressure, abvs)
                        newStutzen.Add(header.StutzenDatalist(i).ID)
                        newAngles.Add(sdata(1))
                        skeys.Add(stutzenkey)
                    Else
                        header.StutzenDatalist(i).ID = newStutzen(skeys.IndexOf(stutzenkey))
                        sdata(1) = newAngles(skeys.IndexOf(stutzenkey))
                    End If
                End If
                header.StutzenDatalist(i).Angle = sdata(1)
                header.StutzenDatalist(i).SpecialTag = specialtag
                header.StutzenDatalist(i).Figure = tempfigure
            End If
        Next

    End Sub

    Shared Sub GACVSpecialCXE(ByRef header As HeaderData, circuit As CircuitData, consys As ConSysData)
        Dim sdata(), stutzenkey As String

        For i As Integer = 0 To header.Ylist.Count - 1
            Dim tempfigure As Integer = 0
            Dim abvs As Double
            Dim specialtag As String = ""

            'change specialtag for holeposition or add .offset
            If header.Ylist(i) = header.Ylist.Max Then
                tempfigure = 4
                specialtag = "s1star"
                abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 50 ^ 2), 3)
            ElseIf header.Ylist(i) = header.Ylist.Max - 25 Then
                tempfigure = 4
                specialtag = "s4starE1"
                abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 50 ^ 2), 3)
            ElseIf header.Ylist(i) = header.Ylist.Max - 50 Then
                tempfigure = 4
                specialtag = "s4starE1"
                abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 50 ^ 2), 3)
            ElseIf header.Ylist(i) = header.Ylist.Max - 100 Then
                tempfigure = 4
                specialtag = "s4starE2"
                abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 25 ^ 2), 3)
            End If

            If tempfigure <> 0 Then
                sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                header.StutzenDatalist(i).ID = sdata(0)

                If GNData.CheckIfTemplate(sdata(0), "Stutzen") Then
                    'create key
                    WSM.CheckoutPart(sdata(0), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                    General.WaitForFile(General.currentjob.Workspace, sdata(0), "par", 100)
                    stutzenkey = CircProps.CreateStutzenkey(tempfigure, header, circuit.CoreTubeOverhang, Math.Abs(abvs), 0, False)
                    stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                    If skeys.IndexOf(stutzenkey) = -1 Then
                        header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, tempfigure, circuit.Pressure, abvs)
                        newStutzen.Add(header.StutzenDatalist(i).ID)
                        newAngles.Add(sdata(1))
                        skeys.Add(stutzenkey)
                    Else
                        header.StutzenDatalist(i).ID = newStutzen(skeys.IndexOf(stutzenkey))
                        sdata(1) = newAngles(skeys.IndexOf(stutzenkey))
                    End If
                End If
                header.StutzenDatalist(i).Angle = sdata(1)
                header.StutzenDatalist(i).SpecialTag = specialtag
                header.StutzenDatalist(i).Figure = tempfigure
            End If
        Next

    End Sub

    Shared Sub GACVSpecialRXF(ByRef header As HeaderData, circuit As CircuitData, consys As ConSysData)
        Dim sdata(), stutzenkey As String

        For i As Integer = 0 To header.Ylist.Count - 1
            Dim tempfigure As Integer = 0
            Dim abvs As Double
            Dim specialtag As String = ""

            'change specialtag for holeposition or add .offset
            If header.Ylist(i) = header.Ylist.Max Then
                tempfigure = 4
                specialtag = "s1star"
                abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + header.Displacever ^ 2), 3)
            ElseIf header.Ylist(i) = header.Ylist.Max - 25 Then
                tempfigure = 45
                abvs = header.StutzenDatalist(i).ABV
                specialtag = "s4starE2"
            ElseIf header.Ylist(i) = header.Ylist.Max - 50 Then
                tempfigure = 4
                specialtag = "s4starF1"
                abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 25 ^ 2), 3)
            End If

            If tempfigure <> 0 Then
                sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 25, New Dictionary(Of Double, Double))
                header.StutzenDatalist(i).ID = sdata(0)

                If GNData.CheckIfTemplate(sdata(0), "Stutzen") Then
                    'create key
                    WSM.CheckoutPart(sdata(0), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                    General.WaitForFile(General.currentjob.Workspace, sdata(0), "par", 100)
                    stutzenkey = CircProps.CreateStutzenkey(tempfigure, header, circuit.CoreTubeOverhang, Math.Abs(abvs), 0, False)
                    stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                    If skeys.IndexOf(stutzenkey) = -1 Then
                        header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, tempfigure, circuit.Pressure, abvs)
                        newStutzen.Add(header.StutzenDatalist(i).ID)
                        newAngles.Add(sdata(1))
                        skeys.Add(stutzenkey)
                    Else
                        header.StutzenDatalist(i).ID = newStutzen(skeys.IndexOf(stutzenkey))
                        sdata(1) = newAngles(skeys.IndexOf(stutzenkey))
                    End If
                End If
                header.StutzenDatalist(i).Angle = sdata(1)
                header.StutzenDatalist(i).SpecialTag = specialtag
                header.StutzenDatalist(i).Figure = tempfigure
            End If
        Next

    End Sub

    Shared Sub GACVSpecialCXBottom(header As HeaderData, circuit As CircuitData, consys As ConSysData)
        Dim sdata(), stutzenkey As String

        For i As Integer = 0 To header.Ylist.Count - 1
            Dim abvs As Double

            If header.Ylist(i) = header.Ylist.Min And header.StutzenDatalist(i).ABV = 0 Then
                abvs = 25.6
                sdata = GNData.GetStutzenID(abvs, header, circuit, consys, 4, 0, New Dictionary(Of Double, Double))
                header.StutzenDatalist(i).ID = sdata(0)

                If sdata(0) = Library.TemplateParts.STUTZEN4 OrElse sdata(0) = Library.TemplateParts.STUTZEN5 OrElse sdata(0) = Library.TemplateParts.STUTZEN45IN OrElse sdata(0) = Library.TemplateParts.STUTZEN45OUT Then
                    'create key
                    WSM.CheckoutPart(sdata(0), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                    General.WaitForFile(General.currentjob.Workspace, sdata(0), "par", 100)
                    stutzenkey = CircProps.CreateStutzenkey(4, header, circuit.CoreTubeOverhang, Math.Abs(abvs), 0, False)
                    stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                    If skeys.IndexOf(stutzenkey) = -1 Then
                        header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, 4, circuit.Pressure, abvs)
                        newStutzen.Add(header.StutzenDatalist(i).ID)
                        newAngles.Add(sdata(1))
                        skeys.Add(stutzenkey)
                    Else
                        header.StutzenDatalist(i).ID = newStutzen(skeys.IndexOf(stutzenkey))
                        sdata(1) = newAngles(skeys.IndexOf(stutzenkey))
                    End If
                End If
                header.StutzenDatalist(i).Angle = sdata(1)
                header.StutzenDatalist(i).SpecialTag = "s4starN"
                header.StutzenDatalist(i).Figure = 4
                Exit For
            End If
        Next

    End Sub

    Shared Sub GACVSpecialCP269(ByRef header As HeaderData, circuit As CircuitData, consys As ConSysData)
        Dim uniqueoutlets As List(Of Double) = General.GetUniqueValues(header.Xlist)
        Dim sdata(), stutzenkey As String

        If uniqueoutlets.Count = 2 Then
            For i As Integer = 0 To header.StutzenDatalist.Count - 1
                If Math.Abs(header.StutzenDatalist(i).ABV) > 40 Then
                    sdata = GNData.GetStutzenID(header.StutzenDatalist(i).ABV, header, circuit, consys, 45, 25, New Dictionary(Of Double, Double))
                    header.StutzenDatalist(i).ID = sdata(0)
                    If sdata(0) = Library.TemplateParts.STUTZEN4 OrElse sdata(0) = Library.TemplateParts.STUTZEN5 OrElse sdata(0) = Library.TemplateParts.STUTZEN45IN OrElse sdata(0) = Library.TemplateParts.STUTZEN45OUT Then
                        'create key
                        WSM.CheckoutPart(sdata(0), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                        General.WaitForFile(General.currentjob.Workspace, sdata(0), "par", 100)
                        stutzenkey = CircProps.CreateStutzenkey(45, header, circuit.CoreTubeOverhang, Math.Abs(header.StutzenDatalist(i).ABV), 0, False)
                        stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                        If skeys.IndexOf(stutzenkey) = -1 Then
                            header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, 45, circuit.Pressure, header.StutzenDatalist(i).ABV)
                            newStutzen.Add(header.StutzenDatalist(i).ID)
                            newAngles.Add(sdata(1))
                            skeys.Add(stutzenkey)
                        Else
                            header.StutzenDatalist(i).ID = newStutzen(skeys.IndexOf(stutzenkey))
                            sdata(1) = newAngles(skeys.IndexOf(stutzenkey))
                        End If
                    End If
                    header.StutzenDatalist(i).Angle = sdata(1)
                    header.StutzenDatalist(i).Figure = 45
                    header.StutzenDatalist(i).SpecialTag = "sOutT45r1"
                End If
            Next
        End If

    End Sub

    Shared Sub GFDVSpecialOutlet(ByRef header As HeaderData, circuit As CircuitData, consys As ConSysData)
        Dim sdata(), stutzenkey As String

        For i As Integer = 0 To header.Ylist.Count - 1
            Dim tempfigure As Integer
            Dim abvs, offset As Double
            Dim specialtag As String = ""

            If header.Ylist(i) = header.Ylist.Max Then
                'F4-1P is different as it has 2 special ones
                If circuit.FinType = "G" Then
                    abvs = -50
                    offset = 25
                Else
                    abvs = -32
                    offset = 7
                End If
                If header.StutzenDatalist(i).ABV = 0 Then
                    tempfigure = 4
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    specialtag = "sOutT4n"
                Else
                    tempfigure = 45
                    sdata = GNData.GetStutzenID(header.StutzenDatalist(i).ABV, header, circuit, consys, tempfigure, abvs, New Dictionary(Of Double, Double))
                    specialtag = "sOutT45"
                End If
                header.StutzenDatalist(i).ID = sdata(0)
                If sdata(0) = Library.TemplateParts.STUTZEN4 OrElse sdata(0) = Library.TemplateParts.STUTZEN5 OrElse sdata(0) = Library.TemplateParts.STUTZEN45IN OrElse sdata(0) = Library.TemplateParts.STUTZEN45OUT Then
                    'create key
                    WSM.CheckoutPart(sdata(0), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                    General.WaitForFile(General.currentjob.Workspace, sdata(0), "par", 100)
                    stutzenkey = CircProps.CreateStutzenkey(tempfigure, header, circuit.CoreTubeOverhang, Math.Abs(abvs), 0, False)
                    stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                    header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, tempfigure, circuit.Pressure, abvs)
                End If
                header.StutzenDatalist(i).Angle = Math.Abs(CInt(sdata(1)))
                header.StutzenDatalist(i).Figure = tempfigure
            ElseIf circuit.FinType = "G" AndAlso header.Ylist(i) = header.Ylist.Max - 50 Then
                If header.StutzenDatalist(i).ABV = 0 Then
                    tempfigure = 4
                    sdata = GNData.GetStutzenID(-50, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    specialtag = "sOutT4n"
                Else
                    'Use???
                    tempfigure = 45
                    sdata = GNData.GetStutzenID(header.StutzenDatalist(i).ABV, header, circuit, consys, tempfigure, 50, New Dictionary(Of Double, Double))
                    specialtag = "sOutT45"
                End If
                offset = 75
                header.StutzenDatalist(i).ID = sdata(0)
                If GNData.CheckIfTemplate(sdata(0), "Stutzen") Then
                    'create key
                    WSM.CheckoutPart(sdata(0), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                    General.WaitForFile(General.currentjob.Workspace, sdata(0), "par", 100)
                    stutzenkey = CircProps.CreateStutzenkey(tempfigure, header, circuit.CoreTubeOverhang, Math.Abs(abvs), 0, False)
                    stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                    header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, tempfigure, circuit.Pressure, abvs)
                    If skeys.IndexOf(stutzenkey) = -1 Then
                        header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, tempfigure, circuit.Pressure, abvs)
                        newStutzen.Add(header.StutzenDatalist(i).ID)
                        newAngles.Add(sdata(1))
                        skeys.Add(stutzenkey)
                    Else
                        header.StutzenDatalist(i).ID = newStutzen(skeys.IndexOf(stutzenkey))
                        sdata(1) = newAngles(skeys.IndexOf(stutzenkey))
                    End If
                End If
                header.StutzenDatalist(i).Angle = sdata(1)
                header.StutzenDatalist(i).Figure = tempfigure
            End If
            header.StutzenDatalist(i).HoleOffset = offset
            header.StutzenDatalist(i).SpecialTag = specialtag
        Next

    End Sub

    Shared Sub GGDVSpecialInlet(ByRef header As HeaderData, circuit As CircuitData, consys As ConSysData)
        Dim sdata(), stutzenkey As String

        For i As Integer = 0 To header.Ylist.Count - 1
            Dim tempfigure As Integer = 0
            Dim abvs, offset As Double
            Dim specialtag As String = ""

            If header.Ylist(i) = header.Ylist.Min Then
                tempfigure = 4
                If Math.Abs(header.StutzenDatalist(i).ABV) = 25 OrElse (circuit.NoPasses = 4 And circuit.NoDistributions = 72) Then
                    abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 50 ^ 2), 2)
                Else
                    abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 25 ^ 2), 2)
                End If
                If header.StutzenDatalist(i).ABV = 0 Then
                    specialtag = "sInT4n"
                Else
                    specialtag = "sInT4r6"
                End If
                If circuit.NoPasses = 4 And circuit.NoDistributions = 72 Then
                    'E6
                    offset = 25
                ElseIf circuit.NoPasses = 3 And circuit.NoDistributions = 64 And header.Ylist(i) = 1637.5 Then
                    offset = 0
                    abvs = 50
                    header.Displacever = 50
                    header.Origin(2) += 25
                Else
                    offset = 0
                End If
                sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
            ElseIf header.Ylist(i) = header.Ylist.Min + 25 Then
                If header.StutzenDatalist(i).ABV = 0 Then
                    tempfigure = 4
                    abvs = 50
                    sdata = GNData.GetStutzenID(50, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    specialtag = "sInT4n"
                    offset = 25
                ElseIf Math.Abs(header.StutzenDatalist(i).ABV) = 50 Then
                    tempfigure = 45
                    abvs = header.StutzenDatalist(i).ABV
                    sdata = GNData.GetStutzenID(header.StutzenDatalist(i).ABV, header, circuit, consys, tempfigure, 25, New Dictionary(Of Double, Double))
                    specialtag = "sInT45"
                    offset = 0
                ElseIf circuit.NoPasses = 3 And circuit.NoDistributions = 64 And header.Ylist(i) = 862.5 Then
                    tempfigure = 4
                    abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 25 ^ 2), 2)
                    sdata = GNData.GetStutzenID(50, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    specialtag = "sInT4n"
                    offset = 25
                End If
            ElseIf header.Ylist(i) = header.Ylist.Min + 50 Then
                If circuit.NoPasses = 4 And circuit.NoDistributions = 72 Then
                    tempfigure = 4
                    abvs = 25
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    specialtag = "sInT4n"
                    offset = 50
                ElseIf circuit.NoPasses = 3 And circuit.NoDistributions = 64 And (header.Ylist(i) = 1687.5 Or header.Ylist(i) = 887.5) Then
                    tempfigure = 4
                    abvs = 25
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    If header.Ylist(i) = 887.5 Then
                        offset = 50
                    Else
                        offset = 25
                    End If
                    specialtag = "sInT4n"
                End If
            ElseIf header.Ylist(i) = header.Ylist.Min + 75 And header.StutzenDatalist(i).ABV = 0 Then
                tempfigure = 4
                abvs = 25
                sdata = GNData.GetStutzenID(25, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                specialtag = "sInT4n"
                offset = 50
            End If

            If tempfigure > 0 Then
                header.StutzenDatalist(i).ID = sdata(0)
                If GNData.CheckIfTemplate(sdata(0), "Stutzen") Then
                    'create key
                    WSM.CheckoutPart(sdata(0), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                    General.WaitForFile(General.currentjob.Workspace, sdata(0), "par", 100)
                    stutzenkey = CircProps.CreateStutzenkey(tempfigure, header, circuit.CoreTubeOverhang, Math.Abs(abvs), 0, False)
                    stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                    If skeys.IndexOf(stutzenkey) = -1 Then
                        header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, tempfigure, circuit.Pressure, abvs)
                        newStutzen.Add(header.StutzenDatalist(i).ID)
                        newAngles.Add(sdata(1))
                        skeys.Add(stutzenkey)
                    Else
                        header.StutzenDatalist(i).ID = newStutzen(skeys.IndexOf(stutzenkey))
                        sdata(1) = newAngles(skeys.IndexOf(stutzenkey))
                    End If
                End If
                header.StutzenDatalist(i).Angle = sdata(1)
                header.StutzenDatalist(i).Figure = tempfigure
                header.StutzenDatalist(i).HoleOffset = offset
                header.StutzenDatalist(i).SpecialTag = specialtag
            End If
        Next

    End Sub

    Shared Sub GGDVSpecialOutlet(ByRef header As HeaderData, circuit As CircuitData, consys As ConSysData)
        Dim sdata(), stutzenkey As String

        'instead of exact coordinates for the header splitting, use gap value, optional with 50mm default. 
        '→ if gap < 50mm, change the displacement and use a new type 4

        For i As Integer = 0 To header.Ylist.Count - 1
            Dim tempfigure As Integer = 0
            Dim abvs, offset As Double

            If header.Ylist(i) = header.Ylist.Max Then
                tempfigure = 4
                If Math.Abs(header.StutzenDatalist(i).ABV) = 25 OrElse (circuit.NoPasses = 4 And circuit.NoDistributions = 72) Then
                    abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 50 ^ 2), 2)
                Else
                    abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 25 ^ 2), 2)
                End If
                If header.StutzenDatalist(i).ABV = 0 Then
                    header.StutzenDatalist(i).SpecialTag = "sOutT4n"
                Else
                    header.StutzenDatalist(i).SpecialTag = "sOutT4r1"
                End If
                If circuit.NoPasses = 4 And circuit.NoDistributions = 72 Then
                    offset = 25
                ElseIf circuit.NoPasses = 3 And circuit.NoDistributions = 64 And header.Ylist(i) = 762.5 Then
                    offset = 0
                    abvs = 50
                    header.Displacever = -50
                Else
                    offset = 0
                End If
                sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
            ElseIf header.Ylist(i) = header.Ylist.Max - 25 Then
                If header.StutzenDatalist(i).ABV = 0 Then
                    tempfigure = 4
                    abvs = 50
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    header.StutzenDatalist(i).SpecialTag = "sOutT4n"
                    offset = 25
                ElseIf Math.Abs(header.StutzenDatalist(i).ABV) = 50 Then
                    tempfigure = 45
                    abvs = header.StutzenDatalist(i).ABV
                    sdata = GNData.GetStutzenID(header.StutzenDatalist(i).ABV, header, circuit, consys, tempfigure, 25, New Dictionary(Of Double, Double))
                    header.StutzenDatalist(i).SpecialTag = "sOutT45"
                    offset = 0
                ElseIf circuit.NoPasses = 3 And circuit.NoDistributions = 64 And header.Ylist(i) = 1537.5 Then
                    tempfigure = 4
                    abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 25 ^ 2), 2)
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    header.StutzenDatalist(i).SpecialTag = "sOutT4r2"
                    offset = 25
                End If
            ElseIf header.Ylist(i) = header.Ylist.Max - 50 Then
                If circuit.NoPasses = 4 And circuit.NoDistributions = 72 Then
                    tempfigure = 4
                    abvs = 25
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    header.StutzenDatalist(i).SpecialTag = "sOutT4n"
                    offset = 50
                ElseIf circuit.NoPasses = 3 And circuit.NoDistributions = 64 And (header.Ylist(i) = 712.5 Or 1512.5) Then
                    tempfigure = 4
                    abvs = 25
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    header.StutzenDatalist(i).SpecialTag = "sOutT4n"
                    If header.Ylist(i) = 712.5 Then
                        offset = 25
                    Else
                        offset = 50
                    End If
                End If
            ElseIf header.Ylist(i) = header.Ylist.Max - 75 And header.StutzenDatalist(i).ABV = 0 Then
                tempfigure = 4
                abvs = 25
                sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                header.StutzenDatalist(i).SpecialTag = "sOutT4n"
                offset = 50
            End If
            If tempfigure > 0 Then
                header.StutzenDatalist(i).ID = sdata(0)
                If GNData.CheckIfTemplate(sdata(0), "Stutzen") Then
                    'create key
                    WSM.CheckoutPart(sdata(0), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                    General.WaitForFile(General.currentjob.Workspace, sdata(0), "par", 100)
                    stutzenkey = CircProps.CreateStutzenkey(tempfigure, header, circuit.CoreTubeOverhang, Math.Abs(abvs), 0, False)
                    stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                    If skeys.IndexOf(stutzenkey) = -1 Then
                        header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, tempfigure, circuit.Pressure, abvs)
                        newStutzen.Add(header.StutzenDatalist(i).ID)
                        newAngles.Add(sdata(1))
                        skeys.Add(stutzenkey)
                    Else
                        header.StutzenDatalist(i).ID = newStutzen(skeys.IndexOf(stutzenkey))
                        sdata(1) = newAngles(skeys.IndexOf(stutzenkey))
                    End If
                End If
                header.StutzenDatalist(i).Angle = sdata(1)
                header.StutzenDatalist(i).Figure = tempfigure
                header.StutzenDatalist(i).HoleOffset = offset
            End If
        Next

    End Sub

    Shared Sub GCDVSpecialInlet(ByRef header As HeaderData, circuit As CircuitData, consys As ConSysData)
        Dim sdata(), stutzenkey As String

        For i As Integer = 0 To header.Ylist.Count - 1
            Dim tempfigure As Integer = 0
            Dim abvs, offset As Double
            Dim specialtag As String = ""

            If header.Ylist(i) = header.Ylist.Min Then
                If circuit.NoPasses = 1 Then
                    tempfigure = 45
                    abvs = header.StutzenDatalist(i).ABV
                    sdata = GNData.GetStutzenID(header.StutzenDatalist(i).ABV, header, circuit, consys, tempfigure, 35, New Dictionary(Of Double, Double))
                    specialtag = "sInT45"
                    offset = 0
                ElseIf circuit.NoPasses = 4 Then
                    'can only be E6
                    tempfigure = 4
                    abvs = 40
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    specialtag = "sInT4n"
                    offset = 15
                ElseIf Math.Abs(header.Displacehor) = 0 Then
                    If header.StutzenDatalist(i).ABV = 0 Then
                        tempfigure = 4
                        abvs = header.Displacever
                        sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                        specialtag = "sInT4n"
                    Else
                        tempfigure = 45
                        abvs = header.StutzenDatalist(i).ABV
                        sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 35, New Dictionary(Of Double, Double))
                        specialtag = "sInT45"
                    End If
                    offset = 0
                ElseIf circuit.NoPasses = 2 And circuit.NoDistributions = 96 Then
                    'describes E/4
                    tempfigure = 4
                    abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 35 ^ 2), 2)
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    specialtag = "sInT4r1"
                    offset = 0
                    'different specialtag necessary
                Else
                    Debug.Print("check, no match found")
                End If
            ElseIf header.Ylist(i) = header.Ylist.Min + 25 Then
                If Math.Abs(header.Displacehor) = 0 Then
                    If header.StutzenDatalist(i).ABV = 0 Then
                        tempfigure = 4
                        abvs = 10
                        sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                        specialtag = "sInT4n"
                    ElseIf circuit.NoPasses <= 2 Then
                        tempfigure = 45
                        abvs = header.StutzenDatalist(i).ABV
                        sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 10, New Dictionary(Of Double, Double))
                        specialtag = "sInT45"
                    End If
                    offset = 0
                Else
                    tempfigure = 4
                    If header.Tube.Diameter > 54 Then
                        abvs = Math.Round(Math.Sqrt(38.7 ^ 2 + header.StutzenDatalist(i).ABV ^ 2), 2)
                        specialtag = "sInT4r3"
                        offset = 28.7
                    Else
                        abvs = Math.Round(Math.Sqrt(60.7 ^ 2 + header.StutzenDatalist(i).ABV ^ 2), 2)
                        specialtag = "sInT4r2"
                        offset = 50.7
                    End If
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                End If
            ElseIf circuit.NoPasses = 4 AndAlso header.Ylist(i) = header.Ylist.Min + 50 Then
                tempfigure = 4
                abvs = 25
                sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                specialtag = "sInT4n"
                offset = 50
            End If

            If tempfigure > 0 Then
                header.StutzenDatalist(i).ID = sdata(0)
                If GNData.CheckIfTemplate(sdata(0), "Stutzen") Then
                    'create key
                    WSM.CheckoutPart(sdata(0), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                    General.WaitForFile(General.currentjob.Workspace, sdata(0), "par", 100)
                    stutzenkey = CircProps.CreateStutzenkey(tempfigure, header, circuit.CoreTubeOverhang, Math.Abs(abvs), 0, False)
                    stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                    If skeys.IndexOf(stutzenkey) = -1 Then
                        header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, tempfigure, circuit.Pressure, abvs)
                        newStutzen.Add(header.StutzenDatalist(i).ID)
                        newAngles.Add(sdata(1))
                        skeys.Add(stutzenkey)
                    Else
                        header.StutzenDatalist(i).ID = newStutzen(skeys.IndexOf(stutzenkey))
                        sdata(1) = newAngles(skeys.IndexOf(stutzenkey))
                    End If
                End If
                header.StutzenDatalist(i).Angle = sdata(1)
                header.StutzenDatalist(i).SpecialTag = specialtag
                header.StutzenDatalist(i).Figure = tempfigure
                header.StutzenDatalist(i).HoleOffset = offset
            End If
        Next

    End Sub

    Shared Sub GCDVNH3SpecialInlet(ByRef header As HeaderData, circuit As CircuitData, consys As ConSysData)
        Dim sdata(), stutzenkey As String

        For i As Integer = 0 To header.Ylist.Count - 1
            Dim tempfigure As Integer = 0
            Dim abvs, offset As Double
            Dim specialtag As String = ""

            If header.Ylist(i) = header.Ylist.Min Then
                If header.Displacehor <> 0 Then
                    'E4
                    tempfigure = 4
                    abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 29.5 ^ 2), 2)
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    specialtag = "sInT4r4"
                Else
                    'E6
                    tempfigure = 45
                    abvs = header.StutzenDatalist(i).ABV
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 37.5, New Dictionary(Of Double, Double))
                    specialtag = "sInT45"
                End If
                offset = 0
            ElseIf header.Ylist(i) = header.Ylist.Min + 25 Then
                If header.Displacehor <> 0 Then
                    'E4
                    tempfigure = 4
                    abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 25 ^ 2), 2)
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    offset = 20.5
                    specialtag = "sInT4r5"
                Else
                    If header.StutzenDatalist(i).ABV = 0 Then
                        tempfigure = 4
                        abvs = 25
                        sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                        specialtag = "sInT4n"
                    Else
                        tempfigure = 45
                        abvs = header.StutzenDatalist(i).ABV
                        sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 25, New Dictionary(Of Double, Double))
                        specialtag = "sInT45"
                    End If
                    offset = 12.5
                End If
            ElseIf header.Ylist(i) = header.Ylist.Min + 50 Then
                If header.Displacehor <> 0 Then
                    'E4
                    tempfigure = 5
                    abvs = header.StutzenDatalist(i).ABV
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    offset = 0
                Else
                    'E6
                    tempfigure = 45
                    abvs = header.StutzenDatalist(i).ABV
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 37.5, New Dictionary(Of Double, Double))
                    specialtag = "sInT45"
                    offset = 50
                End If
            End If

            If tempfigure > 0 Then
                header.StutzenDatalist(i).ID = sdata(0)
                If GNData.CheckIfTemplate(sdata(0), "Stutzen") Then
                    'create key
                    WSM.CheckoutPart(sdata(0), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                    General.WaitForFile(General.currentjob.Workspace, sdata(0), "par", 100)
                    stutzenkey = CircProps.CreateStutzenkey(tempfigure, header, circuit.CoreTubeOverhang, Math.Abs(abvs), 0, False)
                    stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                    If skeys.IndexOf(stutzenkey) = -1 Then
                        header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, tempfigure, circuit.Pressure, abvs)
                        newStutzen.Add(header.StutzenDatalist(i).ID)
                        newAngles.Add(sdata(1))
                        skeys.Add(stutzenkey)
                    Else
                        header.StutzenDatalist(i).ID = newStutzen(skeys.IndexOf(stutzenkey))
                        sdata(1) = newAngles(skeys.IndexOf(stutzenkey))
                    End If
                End If
                header.StutzenDatalist(i).Angle = sdata(1)
                header.StutzenDatalist(i).SpecialTag = specialtag
                header.StutzenDatalist(i).Figure = tempfigure
                header.StutzenDatalist(i).HoleOffset = offset
            End If
        Next

    End Sub

    Shared Sub GCDVNH3SpecialOutlet(ByRef header As HeaderData, circuit As CircuitData, consys As ConSysData)
        Dim sdata(), stutzenkey As String

        For i As Integer = 0 To header.Ylist.Count - 1
            Dim tempfigure As Integer = 0
            Dim abvs, offset As Double

            If header.Ylist(i) = header.Ylist.Max Then
                tempfigure = 45
                abvs = 25
                sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 37.5, New Dictionary(Of Double, Double))
                offset = 12.5
            End If

            If tempfigure > 0 Then
                header.StutzenDatalist(i).ID = sdata(0)
                If GNData.CheckIfTemplate(sdata(0), "Stutzen") Then
                    'create key
                    WSM.CheckoutPart(sdata(0), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                    General.WaitForFile(General.currentjob.Workspace, sdata(0), "par", 100)
                    stutzenkey = CircProps.CreateStutzenkey(tempfigure, header, circuit.CoreTubeOverhang, Math.Abs(abvs), 0, False)
                    stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                    If skeys.IndexOf(stutzenkey) = -1 Then
                        header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, tempfigure, circuit.Pressure, abvs)
                        newStutzen.Add(header.StutzenDatalist(i).ID)
                        newAngles.Add(sdata(1))
                        skeys.Add(stutzenkey)
                    Else
                        header.StutzenDatalist(i).ID = newStutzen(skeys.IndexOf(stutzenkey))
                        sdata(1) = newAngles(skeys.IndexOf(stutzenkey))
                    End If
                End If
                header.StutzenDatalist(i).Angle = sdata(1)
                header.StutzenDatalist(i).SpecialTag = "sOutT45"
                header.StutzenDatalist(i).Figure = tempfigure
                header.StutzenDatalist(i).HoleOffset = offset
            End If
        Next
    End Sub

    Shared Sub GxDVMCInletRR4(ByRef header As HeaderData, circuit As CircuitData, consys As ConSysData)
        Dim sdata(), stutzenkey As String
        Dim contype As String

        If circuit.Pressure < 17 Then
            If circuit.NoPasses = 3 Then
                Dim sindex As Integer = header.Ylist.IndexOf(header.Ylist.Min)
                If header.StutzenDatalist(sindex).ABV = 0 Then
                    contype = "Type1"
                    If header.Ylist.IndexOf(header.Ylist.Min + 25) > -1 Then
                        contype = "Type3"
                    ElseIf header.Ylist.IndexOf(header.Ylist.Min + 75) > -1 Then
                        contype = "Type4"
                    End If
                Else
                    contype = "Type2"
                End If
            ElseIf circuit.NoPasses = 2 Then
                Dim minindex As Integer = header.Ylist.IndexOf(header.Ylist.Min)
                Dim maxindex As Integer = header.Ylist.IndexOf(header.Ylist.Min + 25)
                If Math.Abs(header.StutzenDatalist(minindex).ABV) > Math.Abs(header.StutzenDatalist(maxindex).ABV) Then
                    contype = "Type1"
                    Select Case header.Tube.Diameter
                        Case 88.9
                            header.Displacever = 47.5
                        Case 76.1
                            header.Displacever = 50
                        Case Else
                            header.Displacever = 39
                    End Select
                Else
                    contype = "Type2"
                    Select Case header.Tube.Diameter
                        Case 88.9
                            header.Displacever = 45.4
                        Case 76.1
                            header.Displacever = 54
                        Case Else
                            header.Displacever = 39
                    End Select
                End If
            Else
                contype = ""
            End If

            For i As Integer = 0 To header.Ylist.Count - 1
                Dim tempfigure As Integer
                Dim abvs, offset As Double
                Dim specialtag As String = ""
                Dim tempsdata As New Dictionary(Of Double, Double)
                tempfigure = 0
                offset = 0
                If header.Ylist(i) = header.Ylist.Min Then
                    If circuit.NoPasses = 1 Then
                        If Math.Abs(header.StutzenDatalist(i).ABV) = 25 Then
                            tempfigure = 45
                            abvs = header.StutzenDatalist(i).ABV
                            specialtag = "sOutT45"
                        Else
                            tempfigure = 45
                            abvs = header.StutzenDatalist(i).ABV
                            specialtag = "sOutT45"
                        End If
                    ElseIf circuit.NoPasses = 2 Then
                        tempfigure = 4
                        specialtag = "sOutT4n"
                        abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + header.Displacever ^ 2), 3)
                        If contype = "Type1" And header.Tube.Diameter = 76.1 Then
                            offset = 8.9
                        End If
                    ElseIf circuit.NoPasses = 3 Then
                        If contype = "Type2" Then
                            header.Displacever = 35
                            tempfigure = 4
                            abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 35 ^ 2), 3)
                        Else
                            tempfigure = 4
                            abvs = 40
                            specialtag = "sOutT4n"
                            header.Displacever = 40
                            If header.Tube.Diameter < 104 Then
                                If contype = "Type1" Then
                                    header.Displacever = 25
                                    abvs = 25
                                Else
                                    header.Displacever = 0
                                    tempfigure = 0
                                    offset = 0
                                End If
                            Else
                                If contype = "Type3" Then
                                    offset = 15
                                    header.Displacever = 25
                                ElseIf contype = "Type4" Then
                                    abvs = 25
                                    header.Displacever = 25
                                End If
                            End If
                        End If
                    End If
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 35, New Dictionary(Of Double, Double))
                ElseIf header.Ylist(i) = header.Ylist.Min + 25 Then
                    If circuit.NoPasses = 1 Then
                        If header.StutzenDatalist(i).ABV = 0 Then
                            tempfigure = 4
                            abvs = 25
                            specialtag = "sOutT4n"
                            offset = 15
                        Else
                            tempfigure = 45
                            abvs = header.StutzenDatalist(i).ABV
                            specialtag = "sOutT45"
                        End If
                    ElseIf circuit.NoPasses = 2 Then
                        tempfigure = 4
                        specialtag = "sOutT4n"
                        If header.Tube.Diameter > 88.9 Then
                            abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + header.Displacever ^ 2), 3)
                            offset = 25
                        Else
                            If contype = "Type1" Then
                                If header.Tube.Diameter = 88.9 Then
                                    abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 45.4 ^ 2), 3)
                                    offset = 22.8
                                Else
                                    abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 54 ^ 2), 3)
                                    offset = 29
                                End If
                            Else
                                If header.Tube.Diameter = 88.9 Then
                                    abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 47.5 ^ 2), 3)
                                    offset = 17.2
                                Else
                                    abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 58.9 ^ 2), 3)
                                    offset = 19.9
                                End If
                            End If
                        End If
                        If contype = "Type1" And header.Tube.Diameter = 76.1 Then
                            offset = 8.9
                        End If
                    ElseIf circuit.NoPasses = 3 Then
                        If contype = "Type2" Then
                            tempfigure = 4
                            If header.Tube.Diameter = 104 Then
                                abvs = 40
                                specialtag = "sOutT4n"
                                offset = 30
                            Else
                                abvs = 25
                                specialtag = "sOutT4n"
                                offset = 21.5
                            End If
                        End If
                    End If
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 10, New Dictionary(Of Double, Double))
                ElseIf header.Ylist(i) = header.Ylist.Min + 50 Then
                    If circuit.NoPasses = 1 Then
                        If Math.Abs(header.StutzenDatalist(i).ABV) = 25 Then
                            tempfigure = 45
                            abvs = header.StutzenDatalist(i).ABV
                            specialtag = "sOutT45"
                            offset = 25
                        End If
                    ElseIf circuit.NoPasses = 2 Then
                        'only for header <= 104 and type2
                        tempfigure = 4
                        specialtag = "sOutT4n"
                        offset = 50
                        Select Case header.Tube.Diameter
                            Case 104
                                abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 39 ^ 2), 3)
                            Case 88.9
                                abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 45.4 ^ 2), 3)
                            Case Else
                                abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 54 ^ 2), 3)
                        End Select
                    ElseIf circuit.NoPasses = 3 And contype <> "Type2" And contype <> "Type4" Then
                        If header.Tube.Diameter = 104 Then
                            tempfigure = 4
                            abvs = 25
                            specialtag = "sOutT4n"
                            If contype = "Type1" Then
                                offset = 35
                            Else
                                offset = 50
                            End If
                        End If
                    End If
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 10, New Dictionary(Of Double, Double))
                End If
                If tempfigure > 0 Then
                    header.StutzenDatalist(i).ID = sdata(0)
                    If GNData.CheckIfTemplate(sdata(0), "Stutzen") Then
                        'create key
                        WSM.CheckoutPart(sdata(0), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                        General.WaitForFile(General.currentjob.Workspace, sdata(0), "par", 100)
                        stutzenkey = CircProps.CreateStutzenkey(tempfigure, header, circuit.CoreTubeOverhang, Math.Abs(abvs), 0, False)
                        stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                        If skeys.IndexOf(stutzenkey) = -1 Then
                            header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, tempfigure, circuit.Pressure, abvs)
                            newStutzen.Add(header.StutzenDatalist(i).ID)
                            newAngles.Add(sdata(1))
                            skeys.Add(stutzenkey)
                        Else
                            header.StutzenDatalist(i).ID = newStutzen(skeys.IndexOf(stutzenkey))
                            sdata(1) = newAngles(skeys.IndexOf(stutzenkey))
                        End If
                    End If
                    header.StutzenDatalist(i).Angle = sdata(1)
                    header.StutzenDatalist(i).SpecialTag = specialtag
                    header.StutzenDatalist(i).Figure = tempfigure
                    header.StutzenDatalist(i).HoleOffset = offset
                End If
            Next

        End If

    End Sub

    Shared Sub GxDVMCOutletRR4(ByRef header As HeaderData, circuit As CircuitData, consys As ConSysData)
        Dim sdata(), contype, stutzenkey As String

        If circuit.Pressure < 17 Then
            If circuit.NoPasses = 3 Then
                Dim sindex As Integer = header.Ylist.IndexOf(header.Ylist.Max)
                If header.StutzenDatalist(sindex).ABV = 0 Then
                    contype = "Type1"
                Else
                    contype = "Type2"
                End If
            ElseIf circuit.NoPasses = 2 Then
                Dim sindex As Integer = header.Ylist.IndexOf(header.Ylist.Max)
                If header.StutzenDatalist(sindex).ABV = 0 Then
                    contype = "Type1"
                    header.Displacever = -32
                Else
                    contype = "Type2"
                    header.Displacever = -31.2
                End If
            Else
                contype = ""
            End If
            For i As Integer = 0 To header.Ylist.Count - 1
                Dim tempfigure As Integer
                Dim abvs, offset As Double
                Dim specialtag As String = ""
                Dim tempsdata As New Dictionary(Of Double, Double)
                tempfigure = 0
                offset = 0
                If header.Ylist(i) = header.Ylist.Max Then
                    If circuit.NoPasses = 1 Then
                        If header.StutzenDatalist(i).ABV = 0 Then
                            tempfigure = 4
                            abvs = 45
                            specialtag = "sOutT4n"
                            offset = 10
                        Else
                            tempfigure = 45
                            abvs = header.StutzenDatalist(i).ABV
                            specialtag = "sOutT45"
                            offset = 11
                        End If
                    ElseIf circuit.NoPasses = 2 Then
                        tempfigure = 4
                        specialtag = "sOutT4n"
                        If contype = "Type1" And header.Tube.Diameter <= 88.9 Then
                            offset = 7
                            header.Displacever = -25
                        End If
                        abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + header.Displacever ^ 2), 3)
                    ElseIf circuit.NoPasses = 3 Then
                        tempfigure = 4
                        specialtag = "sOutT4n"
                        If contype = "Type1" Then
                            If header.Tube.Diameter >= 76.1 Then
                                abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 40 ^ 2), 3)
                            Else
                                abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 24.7 ^ 2), 3)
                            End If
                        Else
                            If header.Tube.Diameter >= 76.1 Then
                                header.Displacever = -40
                                abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 35 ^ 2), 3)
                            Else
                                header.Displacever = -25
                                abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 25 ^ 2), 3)
                            End If
                        End If
                    End If
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 46, New Dictionary(Of Double, Double))
                ElseIf header.Ylist(i) = header.Ylist.Max - 25 Then
                    If circuit.NoPasses = 1 Then
                        tempfigure = 45
                        specialtag = "sOutT45"
                        abvs = header.StutzenDatalist(i).ABV
                        'header position is centered → 2 different special stutzen needed, one left and one right
                        If (abvs < 0 And circuit.ConnectionSide = "left") Or (abvs > 0 And circuit.ConnectionSide = "right") Then
                            sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 10, New Dictionary(Of Double, Double))
                        Else
                            sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 10, New Dictionary(Of Double, Double), anglemp:=-1)
                        End If
                    ElseIf circuit.NoPasses = 2 Then
                        tempfigure = 4
                        specialtag = "sOutT4n"
                        If contype = "Type1" Then
                            If header.Tube.Diameter > 88.9 Then
                                abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 31.2 ^ 2), 3)
                                offset = 24.2
                            Else
                                tempfigure = 0
                            End If
                        Else
                            abvs = 32
                            offset = 25.8
                        End If
                        sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    ElseIf circuit.NoPasses = 3 Then
                        tempfigure = 4
                        specialtag = "sOutT4n"
                        If contype = "Type1" Then
                            If header.Tube.Diameter > 64 Then
                                abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 35 ^ 2), 3)
                                offset = 20
                            Else
                                abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 25 ^ 2), 3)
                                offset = 25.3
                            End If
                        ElseIf header.Tube.Diameter > 88.9 Then
                            abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 60 ^ 2), 3)
                            offset = 70
                        End If
                        sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    End If
                ElseIf header.Ylist(i) = header.Ylist.Max - 50 Then
                    If circuit.NoPasses = 1 Then
                        If header.StutzenDatalist(i).ABV = 0 Then
                            tempfigure = 4
                            abvs = 15.8
                            specialtag = "sOutT4n"
                            offset = 30.8
                        Else
                            tempfigure = 45
                            specialtag = "sOutT45"
                            abvs = header.StutzenDatalist(i).ABV
                            offset = 30
                        End If
                    ElseIf circuit.NoPasses = 2 Then
                        If contype = "Type1" And header.Tube.Diameter > 88.9 Then
                            tempfigure = 4
                            specialtag = "sOutT4n"
                            abvs = 32
                            offset = 50
                        End If
                    End If
                    sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 15, New Dictionary(Of Double, Double))
                End If
                If tempfigure > 0 Then
                    header.StutzenDatalist(i).ID = sdata(0)
                    If GNData.CheckIfTemplate(sdata(0), "Stutzen") Then
                        'create key
                        WSM.CheckoutPart(sdata(0), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                        General.WaitForFile(General.currentjob.Workspace, sdata(0), "par", 100)
                        stutzenkey = CircProps.CreateStutzenkey(tempfigure, header, circuit.CoreTubeOverhang, Math.Abs(abvs), 0, False)
                        stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                        If skeys.IndexOf(stutzenkey) = -1 Then
                            header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, tempfigure, circuit.Pressure, abvs)
                            newStutzen.Add(header.StutzenDatalist(i).ID)
                            newAngles.Add(sdata(1))
                            skeys.Add(stutzenkey)
                        Else
                            header.StutzenDatalist(i).ID = newStutzen(skeys.IndexOf(stutzenkey))
                            sdata(1) = newAngles(skeys.IndexOf(stutzenkey))
                        End If
                    End If
                    header.StutzenDatalist(i).Angle = sdata(1)
                    header.StutzenDatalist(i).SpecialTag = specialtag
                    header.StutzenDatalist(i).Figure = tempfigure
                    header.StutzenDatalist(i).HoleOffset = offset
                End If
            Next
        End If

    End Sub

    Shared Sub GxDVMCInletRR6(ByRef header As HeaderData, circuit As CircuitData, consys As ConSysData)
        Dim sdata(), stutzenkey As String

        If circuit.Pressure < 17 Then
            For i As Integer = 0 To header.Ylist.Count - 1
                Dim tempfigure, minangle As Integer
                Dim abvs, offset As Double
                Dim specialtag As String = ""
                Dim tempsdata As New Dictionary(Of Double, Double)
                tempfigure = 0
                offset = 0
                If header.Ylist(i) = header.Ylist.Min Then
                    If circuit.NoPasses = 2 Then
                        If header.Tube.Diameter <= 88.9 Then
                            tempfigure = 45
                            abvs = header.StutzenDatalist(i).ABV
                            specialtag = "sInT45"
                            sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 25, New Dictionary(Of Double, Double))
                        Else
                            'Ø104 & Ø133
                            tempfigure = 4
                            abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 39 ^ 2), 2)
                            specialtag = "sInT4r3"
                            offset = -1
                            sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, tempsdata)
                        End If
                    ElseIf circuit.NoPasses = 3 Then
                        tempfigure = 4
                        specialtag = "sInT4r3"
                        If header.Tube.Diameter <= 76.1 Then
                            abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 40 ^ 2), 2)
                            offset = 21
                        Else
                            abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 45 ^ 2), 2)
                            offset = 16
                        End If
                        sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    ElseIf circuit.NoPasses = 4 Then
                        If header.Tube.Diameter >= 104 Then
                            tempfigure = 4
                            specialtag = "sInT4r3"
                            abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 39 ^ 2), 2)
                            offset = 14
                            sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                        End If
                    End If
                ElseIf header.Ylist(i) = header.Ylist.Min + 25 Then
                    If circuit.NoPasses = 2 Then
                        If header.Tube.Diameter <= 88.9 Then
                            If Math.Abs(header.StutzenDatalist(i).ABV) = 50 Then
                                tempfigure = 5
                                'min angle of 48°
                                tempsdata.Add(50, 45)
                                abvs = header.StutzenDatalist(i).ABV
                                sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, tempsdata)
                                minangle = 48
                            End If
                        ElseIf header.Tube.Diameter = 104 Then
                            If Math.Abs(header.StutzenDatalist(i).ABV) < 25 Then
                                tempfigure = 4
                                abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 39 ^ 2), 2)
                                specialtag = "sInT4r3"
                                offset = 24
                                sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, tempsdata)
                            Else
                                tempfigure = 45
                                abvs = header.StutzenDatalist(i).ABV
                                specialtag = "sInT45"
                                sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 15, New Dictionary(Of Double, Double))
                            End If
                        Else
                            'Ø133
                            If Math.Abs(header.StutzenDatalist(i).ABV) < 25 Then
                                tempfigure = 4
                                abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 63.1 ^ 2), 2)
                                specialtag = "sInT4r7"
                                offset = 48.1
                                sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, tempsdata)
                            Else
                                tempfigure = 45
                                abvs = header.StutzenDatalist(i).ABV
                                specialtag = "sInT45"
                                sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 15, New Dictionary(Of Double, Double))
                            End If
                        End If
                    ElseIf circuit.NoPasses = 3 Then
                        If header.Tube.Diameter >= 88.9 Then
                            tempfigure = 45
                            abvs = header.StutzenDatalist(i).ABV
                            specialtag = "sInT45"
                            sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 10, New Dictionary(Of Double, Double))
                        End If
                    End If
                ElseIf header.Ylist(i) = header.Ylist.Min + 50 Then
                    If circuit.NoPasses = 3 And header.Tube.Diameter >= 88.9 Then
                        tempfigure = 4
                        specialtag = "sInT4r3"
                        abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 15.8 ^ 2), 2)
                        offset = 30.8
                        sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    End If
                End If

                If tempfigure > 0 Then
                    header.StutzenDatalist(i).ID = sdata(0)
                    If GNData.CheckIfTemplate(sdata(0), "Stutzen") Then
                        'create key
                        WSM.CheckoutPart(sdata(0), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                        General.WaitForFile(General.currentjob.Workspace, sdata(0), "par", 100)
                        stutzenkey = CircProps.CreateStutzenkey(tempfigure, header, circuit.CoreTubeOverhang, Math.Abs(abvs), minangle, False)
                        stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                        If skeys.IndexOf(stutzenkey) = -1 Then
                            header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, tempfigure, circuit.Pressure, abvs)
                            newStutzen.Add(header.StutzenDatalist(i).ID)
                            newAngles.Add(sdata(1))
                            skeys.Add(stutzenkey)
                        Else
                            header.StutzenDatalist(i).ID = newStutzen(skeys.IndexOf(stutzenkey))
                            sdata(1) = newAngles(skeys.IndexOf(stutzenkey))
                        End If
                    End If
                    header.StutzenDatalist(i).Angle = sdata(1)
                    header.StutzenDatalist(i).SpecialTag = specialtag
                    header.StutzenDatalist(i).Figure = tempfigure
                    header.StutzenDatalist(i).HoleOffset = offset
                End If
            Next
        End If

    End Sub

    Shared Sub GxDVMCOutletRR6(ByRef header As HeaderData, circuit As CircuitData, consys As ConSysData)
        Dim sdata(), stutzenkey As String

        If circuit.Pressure < 17 Then
            For i As Integer = 0 To header.Ylist.Count - 1
                Dim tempfigure As Integer
                Dim abvs, offset As Double
                Dim specialtag As String = ""
                Dim tempsdata As New Dictionary(Of Double, Double)
                tempfigure = 0
                offset = 0
                If header.Ylist(i) = header.Ylist.Max Then
                    If circuit.NoPasses = 2 Then
                        tempfigure = 4
                        abvs = 45
                        specialtag = "sOutT4n"
                        offset = 10
                        sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    ElseIf circuit.NoPasses = 3 Then
                        tempfigure = 4
                        specialtag = "sOutT4"
                        If header.Tube.Diameter <= 76.1 Then
                            abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 40 ^ 2), 2)
                            offset = 21
                        Else
                            abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 45 ^ 2), 2)
                            offset = 16
                        End If
                        sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    ElseIf circuit.NoPasses = 4 Then
                        If header.Tube.Diameter >= 104 And header.StutzenDatalist(i).ABV <> 0 Then
                            tempfigure = 45
                            specialtag = "sOutT4"
                            abvs = header.StutzenDatalist(i).ABV
                            sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 10, New Dictionary(Of Double, Double))
                        End If
                    End If
                ElseIf header.Ylist(i) = header.Ylist.Max - 25 Then
                    If circuit.NoPasses = 2 Then
                        tempfigure = 45
                        specialtag = "sOutT45"
                        abvs = header.StutzenDatalist(i).ABV

                        'header position is centered → 2 different special stutzen needed, one left and one right
                        If (abvs < 0 And circuit.ConnectionSide = "left") Or (abvs > 0 And circuit.ConnectionSide = "right") Then
                            sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 10, New Dictionary(Of Double, Double))
                        Else
                            sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 10, New Dictionary(Of Double, Double), anglemp:=-1)
                        End If
                    ElseIf circuit.NoPasses = 3 Then
                        If header.Tube.Diameter >= 88.9 Then
                            tempfigure = 45
                            abvs = header.StutzenDatalist(i).ABV
                            specialtag = "sInT45"
                            sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 10, New Dictionary(Of Double, Double))
                        End If
                    End If
                ElseIf header.Ylist(i) = header.Ylist.Max - 50 Then
                    If circuit.NoPasses < 4 Then
                        tempfigure = 4
                        abvs = Math.Round(Math.Sqrt(header.StutzenDatalist(i).ABV ^ 2 + 15.8 ^ 2), 2)
                        specialtag = "sOutT4n"
                        offset = 30.8
                        sdata = GNData.GetStutzenID(abvs, header, circuit, consys, tempfigure, 0, New Dictionary(Of Double, Double))
                    End If
                End If
                If tempfigure > 0 Then
                    header.StutzenDatalist(i).ID = sdata(0)
                    If GNData.CheckIfTemplate(sdata(0), "Stutzen") Then
                        'create key
                        WSM.CheckoutPart(sdata(0), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                        General.WaitForFile(General.currentjob.Workspace, sdata(0), "par", 100)
                        stutzenkey = CircProps.CreateStutzenkey(tempfigure, header, circuit.CoreTubeOverhang, Math.Abs(abvs), 0, False)
                        stutzenkey = circuit.CoreTube.Diameter.ToString + "\" + stutzenkey
                        If skeys.IndexOf(stutzenkey) = -1 Then
                            header.StutzenDatalist(i).ID = SEPart.CreateNewStutzen(General.currentjob.Workspace, stutzenkey, sdata(0), circuit.CoreTube.Material, circuit.FinType, tempfigure, circuit.Pressure, abvs)
                            newStutzen.Add(header.StutzenDatalist(i).ID)
                            newAngles.Add(sdata(1))
                            skeys.Add(stutzenkey)
                        Else
                            header.StutzenDatalist(i).ID = newStutzen(skeys.IndexOf(stutzenkey))
                            sdata(1) = newAngles(skeys.IndexOf(stutzenkey))
                        End If
                    End If
                    header.StutzenDatalist(i).Angle = sdata(1)
                    header.StutzenDatalist(i).SpecialTag = specialtag
                    header.StutzenDatalist(i).Figure = tempfigure
                    header.StutzenDatalist(i).HoleOffset = offset
                End If
            Next
        Else


        End If

    End Sub

    Shared Sub GetBrineStutzen(ByRef consys As ConSysData, coil As CoilData, circuit As CircuitData)
        Dim abvlist, uniqueabsabvlist As List(Of Double)
        Dim invalues, outvalues As New Dictionary(Of Double, String)       'abv - ID
        Dim abvangle As New Dictionary(Of Double, Double)       'abv - angle 
        Dim stutzenkeys, templateids, IDList As New List(Of String)
        Dim anglelist As New List(Of Double)
        Dim figurelist As New List(Of Integer)

        Try
            If circuit.NoDistributions > 1 Then
                With consys.InletHeaders
                    .First.Origin = GetBrineHeaderOrigin(.First, consys.HeaderAlignment, .First.Dim_a, circuit.ConnectionSide)

                    abvlist = GetABVList(.First, consys.HeaderAlignment)
                    uniqueabsabvlist = SortABV(abvlist)
                    uniqueabsabvlist.Sort()

                    For i As Integer = 0 To uniqueabsabvlist.Count - 1
                        Dim sID As String
                        Dim angle As Double
                        Dim figure As Integer = GNData.DefineFigure(consys.HeaderAlignment, circuit, .First, uniqueabsabvlist, uniqueabsabvlist(i), False, coil.NoRows)
                        Dim sentry() As String = GNData.GetStutzenID(uniqueabsabvlist(i), .First, circuit, consys, figure, 0, abvangle)
                        IDList.Add(sentry(0))
                        anglelist.Add(sentry(1))
                        sID = sentry(0)
                        angle = sentry(1)
                        abvangle.Add(uniqueabsabvlist(i), angle)
                        invalues.Add(uniqueabsabvlist(i), sID)
                        figurelist.Add(figure)
                    Next

                    'for every template, create a model
                    CheckForTemplate(IDList, anglelist, stutzenkeys, circuit, .First, uniqueabsabvlist)

                    For i As Integer = 0 To .First.Xlist.Count - 1
                        For j As Integer = 0 To uniqueabsabvlist.Count - 1
                            If abvlist(i) = uniqueabsabvlist(j) Then
                                .First.StutzenDatalist.Add(New StutzenData With {.ID = IDList(j), .Angle = anglelist(j), .ABV = abvlist(i), .SpecialTag = "", .Figure = figurelist(j)})
                            End If
                        Next
                    Next
                End With

                IDList.Clear()
                anglelist.Clear()
                abvangle.Clear()
                figurelist.Clear()

                With consys.OutletHeaders
                    .First.Origin = GetBrineHeaderOrigin(.First, consys.HeaderAlignment, .First.Dim_a, circuit.ConnectionSide)

                    abvlist = GetABVList(.First, consys.HeaderAlignment)
                    uniqueabsabvlist = SortABV(abvlist)
                    uniqueabsabvlist.Sort()

                    For i As Integer = 0 To uniqueabsabvlist.Count - 1
                        Dim sID As String
                        Dim angle As Double
                        Dim figure As Integer = GNData.DefineFigure(consys.HeaderAlignment, circuit, .First, uniqueabsabvlist, uniqueabsabvlist(i), False, coil.NoRows)
                        Dim sentry() As String = GNData.GetStutzenID(uniqueabsabvlist(i), .First, circuit, consys, figure, 0, abvangle)
                        IDList.Add(sentry(0))
                        anglelist.Add(sentry(1))
                        sID = sentry(0)
                        angle = sentry(1)
                        abvangle.Add(uniqueabsabvlist(i), angle)
                        outvalues.Add(uniqueabsabvlist(i), sID)
                        figurelist.Add(figure)
                    Next

                    'for every template, create a model
                    CheckForTemplate(IDList, anglelist, stutzenkeys, circuit, .First, uniqueabsabvlist)

                    For i As Integer = 0 To .First.Xlist.Count - 1
                        For j As Integer = 0 To uniqueabsabvlist.Count - 1
                            If abvlist(i) = uniqueabsabvlist(j) Then
                                .First.StutzenDatalist.Add(New StutzenData With {.ID = IDList(j), .Angle = anglelist(j), .ABV = abvlist(i), .SpecialTag = "", .Figure = figurelist(j)})
                            End If
                        Next
                    Next
                End With
            Else
                Dim spezification, partID As String
                Dim l1calc, l2calc, mcap, xcoord As Double
                Dim stutzendata(), l1list, l2list, alphalist As List(Of String)
                Dim wallthickness As Double

                l2calc = consys.InletHeaders.First.Dim_a - circuit.CoreTubeOverhang + 5
                If consys.HeaderMaterial = "C" Then
                    spezification = "SP01-A"
                    mcap = 0
                    wallthickness = Database.GetTubeThickness("Stub", consys.InletHeaders.First.Tube.Diameter, "C", circuit.Pressure)
                Else
                    spezification = "SP03-2"
                    mcap = 67
                    wallthickness = Database.GetTubeThickness("Stub", consys.InletHeaders.First.Tube.Diameter, "V", circuit.Pressure)
                End If

                consys.InletHeaders.First.Tube.WallThickness = wallthickness
                consys.OutletHeaders.First.Tube.WallThickness = wallthickness

                stutzendata = Database.GetStutzenData("5", consys.InletHeaders.First.Tube.Diameter, spezification, consys.InletHeaders.First.Tube.WallThickness, circuit.Pressure)
                IDList = stutzendata(0)
                l1list = stutzendata(1)
                l2list = stutzendata(2)
                alphalist = stutzendata(3)
                xcoord = consys.InletHeaders.First.Xlist.First

                If circuit.ConnectionSide = "right" Then
                    xcoord = coil.FinnedDepth - xcoord
                End If
                'l1calc is the theoretical length needed to reach to specific horizontal position for the inlet
                'inlet
                l1calc = Math.Round(90 + General.currentunit.TubeSheet.Dim_d + xcoord - (26 + mcap))
                partID = ""
                For i As Integer = 0 To IDList.Count - 1
                    If alphalist(i) = "90" And l2list(i) = l2calc.ToString And Math.Abs(CDbl(l1list(i)) - l1calc) < 3 Then
                        partID = IDList(i)
                        Exit For
                    End If
                Next

                If partID <> "" Then
                    consys.InletHeaders.First.StutzenDatalist.Add(New StutzenData With {.XPos = consys.InletHeaders.First.Xlist.First, .YPos = circuit.CoreTubeOverhang - 5, .ZPos = consys.InletHeaders.First.Ylist.First, .ID = partID})
                End If

                xcoord = consys.OutletHeaders.First.Xlist.First

                If circuit.ConnectionSide = "right" Then
                    xcoord = coil.FinnedDepth - xcoord
                End If
                'outlet
                l1calc = Math.Round(90 + General.currentunit.TubeSheet.Dim_d + xcoord - (26 + mcap))
                partID = ""
                For i As Integer = 0 To IDList.Count - 1
                    If alphalist(i) = "90" And l2list(i) = l2calc.ToString And Math.Abs(CDbl(l1list(i)) - l1calc) < 3 Then
                        partID = IDList(i)
                        Exit For
                    End If
                Next

                If partID <> "" Then
                    consys.OutletHeaders.First.StutzenDatalist.Add(New StutzenData With {.XPos = consys.OutletHeaders.First.Xlist.First, .YPos = circuit.CoreTubeOverhang - 5, .ZPos = consys.OutletHeaders.First.Ylist.First, .ID = partID})
                End If

            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Function GetBrineHeaderOrigin(header As HeaderData, alignment As String, a As Double, conside As String) As Double()
        Dim x, y, z, origin() As Double

        Try
            If alignment = "vertical" Then
                If header.Tube.HeaderType = "inlet" Then
                    If conside = "right" Then
                        x = header.Xlist.Max
                    Else
                        x = header.Xlist.Min
                    End If
                    y = Math.Round(header.Tube.Diameter / 2 + a + 1, 2)
                    z = header.Ylist.Min - 18
                Else
                    If conside = "right" Then
                        x = header.Xlist.Max
                    Else
                        x = header.Xlist.Min
                    End If
                    y = Math.Round(header.Tube.Diameter / 2 + a + 1, 2)
                    z = header.Ylist.Min - 18
                End If
            Else
                If header.Tube.HeaderType = "inlet" Then
                    z = header.Ylist.Min
                Else
                    z = header.Ylist.Max
                End If
                x = header.Xlist.Min - header.Overhangbottom
                y = Math.Round(header.Tube.Diameter / 2 + a + General.currentunit.TubeSheet.Thickness, 2)
            End If
            origin = {x, -y, z}
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return origin
    End Function

    Shared Function GetVentAngle(diameter As Double, coilsize As String, headermat As String, fintype As String, finnedheight As Double, conside As String, headertype As String) As Double
        Dim angle As Double

        If fintype = "N" Then
            If headertype = "inlet" Then
                angle = -90
            Else
                angle = 0
            End If
        Else
            If headertype = "inlet" Then
                'can only be CU
                angle = 90
            Else
                If headermat = "C" Then
                    If diameter < 30 Then
                        angle = 25
                    Else
                        Select Case diameter
                            Case 35
                                angle = 15
                            Case 64
                                If coilsize = "F4" Then
                                    angle = -15
                                End If
                            Case 76.1
                                If coilsize = "F4" Then
                                    angle = -35
                                Else
                                    angle = -5
                                End If
                            Case 88.9
                                If coilsize = "F4" Then
                                    angle = -35
                                Else
                                    angle = -5
                                End If
                            Case 104
                                angle = -10
                        End Select
                    End If
                Else
                    If coilsize = "F4" And diameter > 76 Then
                        angle = -90
                    End If
                End If
            End If
        End If

        If conside = "left" Then
            angle = -angle
        End If

        Return angle
    End Function

    Shared Sub SVPosition(ByRef consys As ConSysData, circuit As CircuitData)

        If consys.OutletHeaders.First.Tube.IsBrine Then
            If consys.HeaderAlignment = "horizontal" Then
                If circuit.NoDistributions = 1 Then
                    consys.InletHeaders.First.Tube.SVPosition = {"header", "axial"}
                    If consys.HeaderMaterial = "V" Then
                        consys.InletHeaders.First.Tube.SVPosition(0) = "cap"
                    End If
                End If
                consys.OutletHeaders.First.Tube.SVPosition = {"header", "axial"}
                If consys.HeaderMaterial = "V" Then
                    consys.OutletHeaders.First.Tube.SVPosition(0) = "cap"
                End If
            Else
                consys.InletNipples.First.SVPosition = {"nipple", "perp"}
            End If
        Else
            If General.currentunit.ApplicationType = "Evaporator" Then
                If consys.HeaderAlignment = "horizontal" Then
                    If consys.OutletHeaders.First.Nippletubes = 0 Then
                        If circuit.CoreTube.Materialcodeletter = "C" Then
                            consys.InletHeaders.First.Tube.SVPosition = {"header", "axial"}
                        Else
                            consys.InletHeaders.First.Tube.SVPosition = {"cap", "axial"}
                        End If
                        consys.OutletHeaders.First.Tube.SVPosition = consys.InletHeaders.First.Tube.SVPosition
                    Else
                        consys.InletNipples.First.SVPosition = {"nipple", "axial"}
                        consys.OutletNipples.First.SVPosition = {"nipple", "axial"}
                    End If
                Else
                    If consys.VType = "X" Then
                        'outlet nipple
                        For Each n In consys.OutletNipples
                            n.SVPosition = {"nipple", "perp"}
                            If General.currentunit.ModelRangeSuffix = "AX" Then
                                n.SVPosition = {"cap", "axial"}
                            End If
                        Next
                    Else
                        If consys.HasFTCon Then
                            consys.OutletNipples.First.SVPosition = {"flange", "axial"}
                            If consys.ConType = 1 Then
                                consys.OutletNipples.First.SVPosition = {"", ""}
                            End If
                        ElseIf consys.VType = "P" And circuit.Pressure > 16 Then
                            If consys.InletNipples.First.Materialcodeletter = "C" Then
                                consys.InletNipples.First.SVPosition = {"nipple", "axial"}
                            Else
                                consys.InletNipples.First.SVPosition = {"cap", "axial"}
                                consys.OutletNipples.First.SVPosition = {"cap", "axial"}
                            End If
                        Else
                            consys.OutletNipples.First.SVPosition = {"nipple", "perp"}
                        End If
                    End If
                End If
            Else
                If General.currentunit.UnitDescription = "VShape" Then
                    If consys.HasFTCon Then
                        consys.OutletNipples.First.SVPosition = {"flange", "axial"}
                        consys.InletNipples.First.SVPosition = {"flange", "axial"}
                        If consys.ConType = 1 Then
                            consys.OutletNipples.First.SVPosition = {"", ""}
                            consys.InletNipples.First.SVPosition = {"", ""}
                        End If
                    ElseIf circuit.Pressure < 17 Then
                        consys.OutletNipples.First.SVPosition = {"nipple", "axial"}
                        consys.InletNipples.First.SVPosition = {"nipple", "axial"}
                    Else
                        If consys.HeaderMaterial = "C" Then
                            consys.InletHeaders.First.Tube.SVPosition = {"header", "perp"}
                            consys.InletHeaders.Last.Tube.SVPosition = {"header", "perp"}
                        Else
                            consys.InletNipples.First.SVPosition = {"cap", "axial"}
                        End If
                        consys.OutletNipples.First.SVPosition = {"cap", "axial"}
                    End If
                Else
                    If circuit.Pressure > 16 And General.currentunit.ModelRangeSuffix.Substring(0, 1) <> "A" And circuit.Pressure < 79 Then
                        consys.InletHeaders.First.Tube.SVPosition = {"header", "perp"}
                    Else
                        consys.InletNipples.First.SVPosition = {"nipple", "axial"}
                    End If
                    consys.OutletNipples.First.SVPosition = {"nipple", "axial"}
                    If consys.ConType = 1 Then
                        consys.InletNipples.First.SVPosition = {"", ""}
                        consys.OutletNipples.First.SVPosition = {"", ""}
                    End If
                End If
            End If
        End If

        WSM.CheckoutPart(GNData.GetSVID(circuit.Pressure), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)

    End Sub

    Shared Sub GetAddPlates(asmdoc As SolidEdgeAssembly.AssemblyDocument, consys As ConSysData, coretubeoverhang As Double)
        Dim platelist As New List(Of PlateData)
        Dim stutzenlist, sortedstutzenlist As New List(Of StutzenData)
        Dim occlist As New List(Of SolidEdgeAssembly.Occurrence)
        Dim platediameter As Double
        Dim n, totalplatecount, platecount, tubeno As Integer
        Dim plateID As String

        Try
            n = 1
            totalplatecount = 0

            Do
                plateID = PCFData.GetValue("OrificePlate" + n.ToString, "PDMID")
                platecount = PCFData.GetValue("OrificePlate" + n.ToString, "Quantity", "double")
                platediameter = PCFData.GetValue("OrificePlate" + n.ToString, "OrificeDiameter", "double")
                Dim newplate As New PlateData With {.ID = plateID, .Quantity = platecount, .InnerDiameter = platediameter}
                platelist.Add(newplate)
                WSM.CheckoutPart(plateID, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                n += 1
                totalplatecount += platecount
            Loop Until totalplatecount = consys.InletHeaders.First.Xlist.Count Or plateID = ""

            Dim sortedlist = consys.InletHeaders.First.StutzenDatalist.OrderBy(Function(c) c.ZPos)

            For i As Integer = 0 To sortedlist.Count - 1
                tubeno = SEAsm.GetTubeNo(asmdoc.Occurrences, {sortedlist(i).XPos, sortedlist(i).YPos, sortedlist(i).ZPos}, sortedlist(i).ID)
                If tubeno > 0 Then
                    occlist.Add(asmdoc.Occurrences.Item(tubeno))
                    sortedstutzenlist.Add(sortedlist(i))
                End If
            Next

            SEAsm.AddPlates(asmdoc, occlist, platelist, sortedstutzenlist, General.currentjob.Workspace, consys.HeaderMaterial, consys.Occlist, coretubeoverhang)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function GCDVDisplacever(circuit As CircuitData, coil As CoilData, headerdiameter As Double) As Double
        Dim displacever As Double

        If circuit.Pressure = 46 Or coil.ConSyss.First.HeaderMaterial = "C" Then
            If headerdiameter = 64 AndAlso coil.NoRows = 4 AndAlso circuit.NoPasses = 3 Then
                displacever = 25
            ElseIf headerdiameter = 64 AndAlso coil.NoRows = 6 AndAlso circuit.NoPasses = 4 Then
                displacever = 25
            ElseIf circuit.NoPasses <= 3 Then
                displacever = 35
            End If
        Else
            If coil.NoRows = 4 Then
                If circuit.NoPasses = 2 Then
                    displacever = 29.5
                End If
            Else
                If circuit.NoPasses <= 3 Then
                    displacever = 37.5
                End If
            End If
        End If

        Return displacever
    End Function

    Shared Function GFDVDisplacever(circuit As CircuitData, coil As CoilData, headerdiameter As Double, headertype As String) As Double
        Dim displacever As Double

        If headertype = "inlet" Then
            If circuit.Pressure < 17 Then
                If coil.NoRows = 6 Then
                    If circuit.NoPasses = 2 Then
                        Select Case headerdiameter
                            Case 104
                                displacever = 39
                            Case 133
                                displacever = 39
                            Case Else
                                displacever = 25
                        End Select
                    ElseIf circuit.NoPasses = 3 Then
                        If headerdiameter <= 76.1 Then
                            displacever = 25
                        Else
                            displacever = 35
                        End If
                    ElseIf circuit.NoPasses = 4 And headerdiameter >= 104 Then
                        displacever = 25
                    Else
                        displacever = 0
                    End If
                Else
                    If circuit.NoPasses = 1 Then
                        displacever = 35
                    ElseIf circuit.NoPasses = 2 Then
                        If circuit.ConnectionSide = "right" Then
                            Select Case headerdiameter
                                Case 88.9
                                    displacever = 47.5
                                Case 76.1
                                    displacever = 50
                                Case Else
                                    displacever = 39
                            End Select
                        Else
                            Select Case headerdiameter
                                Case 88.9
                                    displacever = 45.4
                                Case 76.1
                                    displacever = 54
                                Case Else
                                    displacever = 39
                            End Select
                        End If
                    ElseIf circuit.NoPasses = 3 Then
                        displacever = 40
                    Else
                        displacever = 0
                    End If
                End If
            End If
        Else
            If circuit.Pressure < 17 Then
                If coil.NoRows = 6 Then
                    If circuit.NoPasses = 2 Then
                        displacever = -35
                    ElseIf circuit.NoPasses = 3 Then
                        If headerdiameter <= 76.1 Then
                            displacever = -25
                        Else
                            displacever = 35
                        End If
                    ElseIf circuit.nopasses = 4 And headerdiameter >= 104 Then
                        displacever = -10
                    Else
                        displacever = 0
                    End If
                Else
                    If circuit.NoPasses = 1 Then
                        displacever = -35
                    ElseIf circuit.NoPasses = 2 Then
                        If circuit.ConnectionSide = "right" Then
                            If headerdiameter <= 88.9 Then
                                displacever = -25
                            Else
                                displacever = -32
                            End If
                        Else
                            displacever = -31.2
                        End If
                    ElseIf circuit.NoPasses = 3 Then
                        displacever = -40
                    Else
                        displacever = 0
                    End If
                End If
            End If
        End If

        Return displacever
    End Function

    Shared Function ThreadData(erpcode As String) As Double()
        Dim offset, reduction As Double

        Select Case erpcode
            Case "235"
                offset = 12.5
                reduction = 15.5
            Case "233"
                offset = 15
                reduction = 14
            Case "236"
                offset = 14
                reduction = 17
            Case "238"
                offset = 26
                reduction = 18.4
            Case "239"
                offset = 26
                reduction = 20
            Case "240"
                offset = 30
                reduction = 28
            Case "242"
                offset = 35
                reduction = 35
            Case Else
                offset = 0
                reduction = 0
        End Select

        Return {offset, reduction}
    End Function

    Shared Sub AssignCoords(stutzencoords() As List(Of Double), ByRef consys As ConSysData, finneddepth As Double)

        Try
            For i As Integer = 0 To stutzencoords(0).Count - 1
                If stutzencoords(0)(i) > 0 Then
                    If stutzencoords(0)(i) < finneddepth Then
                        'put it in first header
                        consys.InletHeaders.First.Xlist.Add(stutzencoords(0)(i))
                        consys.InletHeaders.First.Ylist.Add(stutzencoords(1)(i))
                    Else
                        consys.InletHeaders.Last.Xlist.Add(stutzencoords(0)(i))
                        consys.InletHeaders.Last.Ylist.Add(stutzencoords(1)(i))
                    End If
                End If
            Next
            For i As Integer = 0 To stutzencoords(2).Count - 1
                If stutzencoords(2)(i) > 0 Then
                    If stutzencoords(2)(i) < finneddepth Then
                        'put it in first header
                        consys.OutletHeaders.First.Xlist.Add(stutzencoords(2)(i))
                        consys.OutletHeaders.First.Ylist.Add(stutzencoords(3)(i))
                    Else
                        consys.OutletHeaders.Last.Xlist.Add(stutzencoords(2)(i))
                        consys.OutletHeaders.Last.Ylist.Add(stutzencoords(3)(i))
                    End If
                End If
            Next
        Catch ex As Exception

        End Try

    End Sub

End Class
