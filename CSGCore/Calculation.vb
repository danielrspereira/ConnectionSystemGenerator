Public Class Calculation
    Shared Function GetPartOrigin(x As Double, z As Double, side As String, ctoverhang As Double, offsetx As Double, offsetz As Double) As Double()
        Dim origin(), y As Double
        Dim finnedlength As Double
        Dim x1, z1 As Double


        If side.Contains("front") Then
            y = -1 * ctoverhang
        Else
            finnedlength = General.currentunit.Coillist.First.FinnedLength
            y = General.currentunit.TubeSheet.Thickness * 2 + ctoverhang + finnedlength
        End If

        x1 = Math.Round(x + offsetx, 3)
        z1 = Math.Round(z + offsetz, 3)

        origin = {x1, y, z1}
        Return origin
    End Function

    Shared Function GetCircsize(ByRef circuit As CircuitData, alignment As String, circframe As Double(), objsheet As SolidEdgeDraft.Sheet) As Boolean
        Dim switchcoords As Boolean = False
        Dim circsize As Double

        If alignment = "horizontal" Then
            circuit.CircuitSize = {circframe(0) * circuit.Quantity, 0}
        Else
            circuit.CircuitSize = {0, circframe(1) * circuit.Quantity}
        End If
        If SEDraft.GetCoilPosition(objsheet) <> alignment Then
            switchcoords = True
            If alignment = "horizontal" Then
                circsize = circframe(0)
                circuit.CircuitSize = {circframe(1) * circuit.Quantity, 0}
            Else
                circsize = circframe(1)
                circuit.CircuitSize = {0, circframe(0) * circuit.Quantity}
            End If
        End If
        Return switchcoords
    End Function

    Shared Function GetAngleRad(x1 As Double, y1 As Double, x2 As Double, y2 As Double, side As String) As Double
        Dim anglerad, anglegrad As Double
        Dim anglecase As Integer
        Dim posangles As List(Of Double)

        'get all possible angles
        posangles = GetPossibleAngles(x1, y1, x2, y2)

        'get the actual angle
        anglecase = GetAngleCase(x1, y1, x2, y2, side)

        'select the correct angle
        anglegrad = Math.Round(posangles(anglecase - 1), 4)

        anglerad = Math.Round(anglegrad / 180 * Math.PI, 5)

        Return anglerad
    End Function

    Shared Function GetMeasures(side As String, type As String) As Boolean()
        Dim measures() As Boolean

        measures = {False, False, False}

        If side.Contains("front") Then
            measures(2) = False
        Else
            measures(2) = True
        End If

        If type = "stutzen" Then
            measures(0) = True
            measures(2) = False
        End If

        Return measures
    End Function

    Shared Function GetPossibleAngles(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As List(Of Double)
        Dim anglelist As New List(Of Double)
        Dim dx, dy, angletemp, dmax, dmin As Double

        dx = x2 - x1
        dy = y2 - y1
        dmax = Math.Max(Math.Abs(dy), Math.Abs(dx))
        dmin = Math.Min(Math.Abs(dy), Math.Abs(dx))
        angletemp = Math.Atan(Math.Abs(dmin / dmax))
        angletemp = Math.Round(angletemp * 180 / Math.PI, 5)

        anglelist.Add(0 + angletemp)
        anglelist.Add(90 - angletemp)
        anglelist.Add(90 + angletemp)
        anglelist.Add(180 - angletemp)
        anglelist.Add(180 + angletemp)
        anglelist.Add(270 - angletemp)
        anglelist.Add(270 + angletemp)
        anglelist.Add(360 - angletemp)
        anglelist.Add(0)
        anglelist.Add(90)
        anglelist.Add(180)
        anglelist.Add(270)

        Return anglelist
    End Function

    Shared Function GetAngleCase(x1 As Double, y1 As Double, x2 As Double, y2 As Double, side As String) As Integer
        Dim anglecase As Integer = 0
        Dim dx, dy, absdx, absdy As Double

        dy = y2 - y1

        If side.Contains("front") Then
            dx = x2 - x1
            If side.Contains("hp") Then
                dy = y1 - y2
            End If
        Else
            dx = x1 - x2
        End If

        absdx = Math.Abs(dx)
        absdy = Math.Abs(dy)

        If absdx >= absdy And dx > 0 And dy > 0 Then
            anglecase = 1
        ElseIf absdx < absdy And dx > 0 And dy > 0 Then
            anglecase = 2
        ElseIf absdx < absdy And dx < 0 And dy > 0 Then
            anglecase = 3
        ElseIf absdx >= absdy And dx < 0 And dy > 0 Then
            anglecase = 4
        ElseIf absdx >= absdy And dx < 0 And dy < 0 Then
            anglecase = 5
        ElseIf absdx < absdy And dx < 0 And dy < 0 Then
            anglecase = 6
        ElseIf absdx < absdy And dx > 0 And dy < 0 Then
            anglecase = 7
        ElseIf absdx >= absdy And dx > 0 And dy < 0 Then
            anglecase = 8
        End If
        If dy = 0 Then
            If dx > 0 Then
                anglecase = 9
            Else
                anglecase = 11
            End If
        ElseIf dx = 0 Then
            If dy > 0 Then
                anglecase = 10
            Else
                anglecase = 12
            End If
        End If

        Return anglecase
    End Function

    Shared Function CreateNewOverhang(headertype As String, headermaterial As String, nodistributions As Integer) As Double
        Dim d, overhang As Double

        'does not contain the distance from coil start to first stutzen

        d = General.currentunit.TubeSheet.Dim_d

        'requires SV for correct length; now outlet no SV if VA
        overhang = d + 60
        If headertype = "outlet" And headermaterial <> "C" And nodistributions > 1 Then
            overhang += 30
        End If

        Return overhang
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

    Shared Function GetNewNippleLength(diameter As Double, flangetype As String, a As Double) As Double
        Dim length As Double

        If Not flangetype.Contains("Loose") Then
            length = 150
        Else
            length = Math.Round(300 - a - diameter / 2 - 3)
        End If

        Return length
    End Function

    Shared Function MinimumAngle(abvlist As List(Of Double), abv As Double, figure As Integer) As Double
        Dim abslist As New List(Of Double)
        Dim uniquelist As List(Of Double)
        Dim rank As Integer
        Dim minangle As Double

        For i As Integer = 0 To abvlist.Count - 1
            abslist.Add(Math.Abs(abvlist(i)))
        Next
        uniquelist = General.GetUniqueValues(abslist)
        uniquelist.Sort()

        rank = uniquelist.IndexOf(Math.Abs(abv))
        If abslist.IndexOf(0) > -1 Then
            rank -= 1
        End If

        If rank < 0 Then
            minangle = 0
        Else
            Dim angles() As Double = GNData.AllowedAngles("C", figure)

            minangle = angles(rank)
        End If

        Return minangle
    End Function

    Shared Function CompareAFig8(l1 As Double, a As Double, headerwt As Double, ctoverhang As Double, holeposition As String) As Double
        Dim delta, offset As Double

        If holeposition = "tangential" Then
            offset = 5
        Else
            offset = 0
        End If

        delta = Math.Round(l1 + ctoverhang - 5 - (a + headerwt + offset), 3)

        Return delta
    End Function

    Shared Function CompareABVFig5(l1sp As Double, l2sp As Double, abv As Double, angle As Double, a As Double, headerdia As Double, headerwt As Double, ctoverhang As Double, figure As Integer) As Double
        Dim xH, yH, yS, l1, l2, l2dash, ycheck, returna As Double
        Dim override As Boolean = False
        Dim acalcvalues() As Double = {a, a}
        Dim abvcalvalues() As Double = {abv, abv}
        Dim no As Integer

        For i As Integer = 1 To 2
            'run 2 loops, fig 5 stutzen can be switched around
            If i = 1 Then
                l1 = l1sp
                l2 = l2sp
            Else
                l1 = l2sp
                l2 = l1sp
            End If

            'calculating the distance between header hole and actual position of the stutzen 
            xH = Math.Round(headerdia / 2 + a - ctoverhang + 5 - headerdia / 2 * Math.Cos(angle * Math.PI / 180), 3)
            yH = Math.Round(abv - headerdia / 2 * Math.Sin(angle * Math.PI / 180), 3)

            'l2dash = distance between bending center of stutzen and contact point with header
            If angle = 90 Then
                ycheck = Math.Abs(xH - l1)
                l2dash = yH
            Else
                l2dash = Math.Round((xH - l1) / Math.Cos(angle * Math.PI / 180), 3)
                yS = Math.Round(l2dash * Math.Sin(angle * Math.PI / 180), 3)
                ycheck = Math.Abs(yH - yS)
            End If

            If l2 > (l2dash + headerwt + 2) And l2 < (l2dash + headerdia / 2) And ycheck < 3 And l2dash > (headerwt + 25 * Math.Sin(angle / 2 * Math.PI / 180)) Then  ' Or figure = 45)
                'consider length validation of l2 (should not be "too long")
                If (l2 - headerdia / 2 + headerwt) < l2dash Then
                    acalcvalues(i - 1) = Math.Abs(ycheck)
                Else
                    Debug.Print("l2 too long - l2 = " + l2.ToString)
                End If
            End If
        Next

        If acalcvalues.Min < 3 Then
            If acalcvalues(0) = acalcvalues.Min Then
                no = 0
            Else
                no = 1
            End If
            override = True
        End If

        If override Then
            returna = acalcvalues(no)
        Else
            returna = a
        End If

        Return returna
    End Function

    Shared Function CompareAFig4(l1 As Double, a As Double, headerdia As Double, headerwt As Double, ctoverhang As Double, abv As Double, offset As Double) As Double
        Dim delta As Double

        delta = l1 + ctoverhang - 5 - (a + headerwt)

        'stutzen must not be too short
        If delta < 1 Or delta > headerdia / 2 - headerwt Then
            'offset must match
            delta = l1
        Else
            If Math.Abs(Math.Abs(abv) - offset) > 2 Then
                delta = l1
            End If
        End If

        Return delta
    End Function

    Shared Function Fig5Parameters(a As Double, abv As Double, diameter As Double, wallthickness As Double, angles As Double(), ctoverhang As Double) As Double()
        Dim angle, l1, l2, parameters() As Double
        Dim loopexit As Boolean = False
        Dim i As Integer = 0

        parameters = {0, 0, 0}

        Do
            angle = angles(i)
            l2 = Math.Round((abv - diameter / 2 * Math.Sin(angle * Math.PI / 180)) / Math.Sin(angle * Math.PI / 180) + wallthickness + 3, 2)
            l1 = Math.Round(a + diameter / 2 - abv / Math.Tan(angle * Math.PI / 180) - ctoverhang + 5, 2)

            If l1 > 15 And l2 > 15 Then
                loopexit = True
                parameters = {l1, l2, angle}
            End If
            i += 1
            If i = angles.Count Then
                loopexit = True
            End If
        Loop Until loopexit

        Return parameters
    End Function

    Shared Function Fig4Parameters(l1 As Double, abv As Double, Angles() As Double) As Double()
        Dim angle, lm, l2, params() As Double
        Dim i As Integer = 0
        Dim loopexit As Boolean = False

        params = {0, 0}
        Do
            angle = Math.Round(Angles(i) * Math.PI / 180, 4)
            lm = Math.Round(Math.Abs(abv) / Math.Sin(angle), 2)

            If lm > 25 Then
                l2 = Math.Round(l1 - lm * Math.Cos(angle), 2) / 2
                If l1 > 50 And l2 > 15 Then
                    loopexit = True
                    params = {l2, Angles(i)}
                End If
            End If

            i = i + 1
            If i = Angles.Count Then
                loopexit = True
            End If
        Loop Until loopexit = True

        Return params
    End Function

    Shared Function GetTubeLength(figure As Integer, partdoc As SolidEdgePart.PartDocument, diameter As Double) As Double
        Dim L1, L2, ABV, angle, radius, length, a1, a2, a3, L3s, dh, a4, a5 As Double

        Try
            radius = SEPart.GetSetVariableValue("Radius", partdoc.Variables, "get")
            L1 = SEPart.GetSetVariableValue("L1", partdoc.Variables, "get")
            Select Case figure
                Case 1
                    ABV = SEPart.GetSetVariableValue("ABV", partdoc.Variables, "get")
                    a1 = L1 - diameter / 2 - radius
                    a2 = Math.PI * radius / 2
                    a3 = ABV - 2 * radius
                    length = Math.Round(2 * a1 + 2 * a2 + a3, 0)
                Case 5
                    L2 = SEPart.GetSetVariableValue("L2", partdoc.Variables, "get")
                    angle = SEPart.GetSetVariableValue("WKL", partdoc.Variables, "get")
                    a1 = L1 - radius * Math.Tan(Math.PI * angle / 360)
                    a2 = angle * Math.PI * radius / 180
                    a3 = L2 - radius * Math.Tan(Math.PI * angle / 360)
                    length = Math.Round(a1 + a2 + a3, 0)
                Case 8
                    length = L1
                Case 9
                    L2 = SEPart.GetSetVariableValue("L2", partdoc.Variables, "get")
                    L3s = SEPart.GetSetVariableValue("L3Stern", partdoc.Variables, "get")
                    angle = SEPart.GetSetVariableValue("WKL3", partdoc.Variables, "get")
                    a2 = angle * Math.PI * radius / 180
                    dh = Math.Sqrt((radius / Math.Cos(Math.PI * angle / 360) - radius) ^ 2 - (radius * (1 - Math.Cos(Math.PI * angle / 360))) ^ 2) * Math.Tan(Math.PI * angle / 360) + radius * (1 - Math.Cos(Math.PI * angle / 360))
                    a3 = (L2 - dh) / Math.Sin(angle * Math.PI / 180) - radius
                    a4 = Math.PI * radius / 2
                    a5 = ABV - 2 * radius
                    length = Math.Round(2 * L3s + 2 * a3 + 2 * a4 + a5, 0)
            End Select

        Catch ex As Exception

        End Try

        Return length / 1000
    End Function

    Shared Function DefaultHeaderLength(ByRef header As HeaderData, alignment As String) As Double
        If alignment = "horizontal" Then
            header.Tube.Length = Math.Round(header.Xlist.Max - header.Xlist.Min + header.Overhangbottom + header.Overhangtop - header.Displacehor)
        Else
            header.Tube.Length = Math.Round(header.Ylist.Max - header.Ylist.Min + header.Overhangtop + header.Overhangbottom + header.Displacever)
        End If
        Return header.Tube.Length
    End Function

    Shared Sub SplitHeader(ByRef consys As ConSysData, co2split As Boolean, rdsplit As Boolean, circuit As CircuitData, coil As CoilData)
        Dim headerquantity As Integer = 2

        Try
            If co2split Then
                If Math.Max(consys.InletHeaders.First.Tube.Diameter, consys.OutletHeaders.First.Tube.Diameter) > 60 And (consys.InletHeaders.First.Tube.Length / 2 > 800 Or consys.OutletHeaders.First.Tube.Length / 2 > 800) Then
                    headerquantity = 3
                End If
            End If

            For i As Integer = 2 To headerquantity
                'fill headerdata list
                If co2split Or rdsplit Then
                    consys.InletHeaders.Add(New HeaderData)
                    consys.InletHeaders(i - 1).Tube.TubeFile.Fullfilename = General.currentjob.Workspace + "\InletHeader" + i.ToString + "_" + circuit.CircuitNumber.ToString + "_" + coil.Number.ToString + ".par"
                    consys.InletHeaders(i - 1).Tube.FileName = consys.InletHeaders(i - 1).Tube.TubeFile.Fullfilename
                    consys.InletHeaders(i - 1).Tube.TubeFile.Shortname = General.GetShortName(consys.InletHeaders(i - 1).Tube.FileName)

                    With consys.InletHeaders
                        .Last.Dim_a = .First.Dim_a
                        .Last.Displacehor = .First.Displacehor
                        .Last.Displacever = .First.Displacever
                        .Last.Nippletubes = 1
                        .Last.Origin = {0, 0, 0}
                        .Last.Overhangbottom = .First.Overhangbottom
                        .Last.Overhangtop = .First.Overhangtop
                        .Last.StutzenDatalist = .First.StutzenDatalist
                        .Last.Xlist = .First.Xlist
                        .Last.Ylist = .First.Ylist
                        .Last.Tube.Diameter = .First.Tube.Diameter
                        .Last.Tube.HeaderType = .First.Tube.HeaderType
                        .Last.Tube.Materialcodeletter = .First.Tube.Materialcodeletter
                        .Last.Tube.RawMaterial = .First.Tube.RawMaterial
                        .Last.Tube.WallThickness = .First.Tube.WallThickness
                    End With

                    My.Computer.FileSystem.CopyFile(consys.InletHeaders.First.Tube.TubeFile.Fullfilename, consys.InletHeaders.Last.Tube.TubeFile.Fullfilename)

                    consys.InletNipples.Add(New TubeData)
                    consys.InletNipples(i - 1).TubeFile.Fullfilename = General.currentjob.Workspace + "\InletNipple" + i.ToString + "_" + circuit.CircuitNumber.ToString + "_" + coil.Number.ToString + ".par"
                    consys.InletNipples(i - 1).FileName = consys.InletNipples(i - 1).TubeFile.Fullfilename
                    consys.InletNipples(i - 1).TubeFile.Shortname = General.GetShortName(consys.InletNipples(i - 1).FileName)

                    With consys.InletNipples
                        .Last.BottomCapNeeded = .First.BottomCapNeeded
                        .Last.Diameter = .First.Diameter
                        .Last.Length = .First.Length
                        .Last.Material = .First.Material
                        .Last.Materialcodeletter = .First.Materialcodeletter
                        .Last.Quantity = 1
                        .Last.RawMaterial = .First.RawMaterial
                        .Last.SVPosition = .First.SVPosition
                        .Last.TopCapNeeded = .First.TopCapNeeded
                        .Last.WallThickness = .First.WallThickness
                        .Last.Angle = .First.Angle
                    End With

                    My.Computer.FileSystem.CopyFile(consys.InletNipples.First.TubeFile.Fullfilename, consys.InletNipples.Last.TubeFile.Fullfilename)
                End If

                If Not rdsplit Then
                    consys.OutletHeaders.Add(New HeaderData)
                    consys.OutletHeaders(i - 1).Tube.TubeFile.Fullfilename = General.currentjob.Workspace + "\OutletHeader" + i.ToString + "_" + circuit.CircuitNumber.ToString + "_" + coil.Number.ToString + ".par"
                    consys.OutletHeaders(i - 1).Tube.FileName = consys.OutletHeaders(i - 1).Tube.TubeFile.Fullfilename
                    consys.OutletHeaders(i - 1).Tube.TubeFile.Shortname = General.GetShortName(consys.OutletHeaders(i - 1).Tube.FileName)

                    With consys.OutletHeaders
                        .Last.Dim_a = .First.Dim_a
                        .Last.Displacehor = .First.Displacehor
                        .Last.Displacever = .First.Displacever
                        .Last.Nippletubes = 1
                        .Last.Origin = {0, 0, 0}
                        .Last.Overhangbottom = .First.Overhangbottom
                        .Last.Overhangtop = .First.Overhangtop
                        .Last.StutzenDatalist = .First.StutzenDatalist
                        .Last.Xlist = .First.Xlist
                        .Last.Ylist = .First.Ylist
                        .Last.Tube.Diameter = .First.Tube.Diameter
                        .Last.Tube.HeaderType = .First.Tube.HeaderType
                        .Last.Tube.Materialcodeletter = .First.Tube.Materialcodeletter
                        .Last.Tube.RawMaterial = .First.Tube.RawMaterial
                        .Last.Tube.WallThickness = .First.Tube.WallThickness
                    End With

                    My.Computer.FileSystem.CopyFile(consys.OutletHeaders.First.Tube.TubeFile.Fullfilename, consys.OutletHeaders.Last.Tube.TubeFile.Fullfilename)

                    consys.OutletNipples.Add(New TubeData)
                    consys.OutletNipples(i - 1).TubeFile.Fullfilename = General.currentjob.Workspace + "\OutletNipple" + i.ToString + "_" + circuit.CircuitNumber.ToString + "_" + coil.Number.ToString + ".par"
                    consys.OutletNipples(i - 1).FileName = consys.OutletNipples(i - 1).TubeFile.Fullfilename
                    consys.OutletNipples(i - 1).TubeFile.Shortname = General.GetShortName(consys.OutletNipples(i - 1).FileName)

                    With consys.OutletNipples
                        .Last.BottomCapNeeded = .First.BottomCapNeeded
                        .Last.Diameter = .First.Diameter
                        .Last.Length = .First.Length
                        .Last.Material = .First.Material
                        .Last.Materialcodeletter = .First.Materialcodeletter
                        .Last.Quantity = 1
                        .Last.RawMaterial = .First.RawMaterial
                        .Last.SVPosition = .First.SVPosition
                        .Last.TopCapNeeded = .First.TopCapNeeded
                        .Last.WallThickness = .First.WallThickness
                        .Last.Angle = .First.Angle
                    End With

                    My.Computer.FileSystem.CopyFile(consys.OutletNipples.First.TubeFile.Fullfilename, consys.OutletNipples.Last.TubeFile.Fullfilename)
                End If
            Next

            Dim nipplequantity As Integer
            If co2split Or rdsplit Then
                If General.currentunit.UnitDescription <> "VShape" Or rdsplit Then
                    ReworkHeader(consys.InletHeaders, consys.HeaderAlignment, headerquantity, rdsplit, circuit, consys)
                Else
                    ReworkCO2HeaderVShape(consys.InletHeaders, headerquantity, circuit, consys)
                End If
                nipplequantity = consys.InletNipples.First.Quantity
                If nipplequantity = 1 Then
                    consys.InletNipples.First.Quantity = 1
                    consys.InletHeaders.First.Nippletubes = 1
                    consys.InletNipples.Last.Quantity = 1
                    consys.InletHeaders.Last.Nippletubes = 1
                Else
                    consys.InletNipples.First.Quantity = Math.Ceiling(nipplequantity / headerquantity)
                    consys.InletHeaders.First.Nippletubes = Math.Ceiling(nipplequantity / headerquantity)
                    consys.InletNipples.Last.Quantity = Math.Floor(nipplequantity / headerquantity)
                    consys.InletHeaders.Last.Nippletubes = Math.Floor(nipplequantity / headerquantity)
                End If
            End If

            If Not rdsplit Then
                If General.currentunit.UnitDescription <> "VShape" Or rdsplit Then
                    ReworkHeader(consys.OutletHeaders, consys.HeaderAlignment, headerquantity, False, circuit, consys)
                Else
                    ReworkCO2HeaderVShape(consys.OutletHeaders, headerquantity, circuit, consys)
                End If
                nipplequantity = consys.OutletNipples.First.Quantity
                If nipplequantity = 1 Then
                    consys.InletNipples.First.Quantity = 1
                    consys.InletHeaders.First.Nippletubes = 1
                    consys.InletNipples.Last.Quantity = 1
                    consys.InletHeaders.Last.Nippletubes = 1
                Else
                    consys.OutletNipples.First.Quantity = Math.Ceiling(nipplequantity / headerquantity)
                    consys.OutletHeaders.First.Nippletubes = Math.Ceiling(nipplequantity / headerquantity)
                    consys.OutletNipples.Last.Quantity = Math.Floor(nipplequantity / headerquantity)
                    consys.OutletHeaders.Last.Nippletubes = Math.Floor(nipplequantity / headerquantity)
                End If
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub ReworkCO2HeaderVShape(ByRef headerlist As List(Of HeaderData), headerquantity As Integer, circuit As CircuitData, consys As ConSysData)
        Dim valuelist1, gaplist As New List(Of Double)
        Dim deltabottom, deltatop As Double
        Dim newlengthlist As New List(Of Double)

        valuelist1.AddRange(headerlist.First.Ylist.ToArray)
        If headerlist.First.Tube.HeaderType = "inlet" Then
            deltabottom = headerlist.First.Displacever
        Else
            deltatop = Math.Abs(headerlist.First.Displacever)
        End If

        valuelist1.Sort()
        For i As Integer = 0 To valuelist1.Count - 2
            gaplist.Add(Math.Round(valuelist1(i + 1) - valuelist1(i)))
        Next

        For i As Integer = 1 To headerquantity
            Dim startpoint As Double
            Select Case i
                Case 1
                    startpoint = headerlist.First.Origin(2) - deltabottom
                Case 2
                    startpoint = headerlist.First.Origin(2) + headerlist.First.Tube.Length + deltatop
                Case Else
                    startpoint = headerlist(i - 2).Origin(2) + headerlist(i - 2).Tube.Length + deltatop
            End Select

            Dim maxlength As Double
            'get maximum length - consider length/n ? → maxlength = f(diameter, length, i, neededheadercount)
            maxlength = GNData.GetMaxLength(headerlist.First.Tube.Diameter, False)

            Dim firstcoord As Double
            'find first stutzen for the header
            firstcoord = GetFirstStutzen(startpoint, valuelist1)

            Dim headerframe(1) As Double
            'set the frame
            headerframe(0) = startpoint
            headerframe(1) = headerframe(0) + maxlength + deltabottom

            Dim relevantcoords As List(Of Double)
            'find all stutzen within the frame
            relevantcoords = GetAllStutzen(headerframe, valuelist1, headerquantity, "vertical", i)

            Dim newstartpoint As Double = GetStartCoord(headerlist(i - 1), relevantcoords.Min, "vertical")
            headerlist(i - 1).Origin(2) = newstartpoint + deltabottom
            For j As Integer = 0 To 2
                If j <> 2 Then
                    headerlist(i - 1).Origin(j) = headerlist.First.Origin(j)
                End If
            Next

            Dim newlength As Double = Math.Round(headerlist(i - 1).Overhangbottom + headerlist(i - 1).Overhangtop + relevantcoords.Max - relevantcoords.Min - deltabottom - deltatop, 3)
            headerlist(i - 1).Tube.Length = newlength

            'shorten xlist, ylist, idlist, anglelist, abvlist
            ReAssignListItems(headerlist(i - 1), relevantcoords, "vertical", circuit, consys)
        Next

    End Sub

    Shared Sub ReworkHeader(ByRef headerlist As List(Of HeaderData), alignment As String, headerquantity As Integer, rdsplit As Boolean, circuit As CircuitData, consys As ConSysData)
        Dim valuelist1 As New List(Of Double)
        Dim coordno As Integer
        Dim displacement As Double
        Dim newlengthlist As New List(Of Double)

        If alignment = "horizontal" Then
            valuelist1.AddRange(headerlist.First.Xlist.ToArray)
            coordno = 0
            displacement = headerlist.First.Displacehor
        Else
            valuelist1.AddRange(headerlist.First.Ylist.ToArray)
            coordno = 2
            displacement = headerlist.First.Displacever
        End If

        If General.currentunit.ModelRangeName = "GACV" Then
            displacement = 0
        End If

        valuelist1.Sort()

        For i As Integer = 1 To headerquantity
            Dim startpoint As Double
            Select Case i
                Case 1
                    startpoint = headerlist.First.Origin(coordno)
                Case 2
                    startpoint = headerlist.First.Origin(coordno) + headerlist.First.Tube.Length - displacement
                Case Else
                    startpoint = headerlist(i - 2).Origin(coordno) + headerlist(i - 2).Tube.Length - displacement
            End Select

            Dim maxlength As Double
            'get maximum length - consider length/n ? → maxlength = f(diameter, length, i, neededheadercount)
            maxlength = GNData.GetMaxLength(headerlist.First.Tube.Diameter, rdsplit)

            Dim firstcoord As Double
            'find first stutzen for the header
            firstcoord = GetFirstStutzen(startpoint, valuelist1)

            Dim headerframe() As Double
            'set the frame
            headerframe = GetHeaderframe(firstcoord, maxlength)

            Dim relevantcoords As List(Of Double)
            'find all stutzen within the frame
            relevantcoords = GetAllStutzen(headerframe, valuelist1, headerquantity, alignment, i)

            Dim newstartpoint As Double = GetStartCoord(headerlist(i - 1), relevantcoords.Min, alignment)
            headerlist(i - 1).Origin(coordno) = newstartpoint
            For j As Integer = 0 To 2
                If j <> coordno Then
                    headerlist(i - 1).Origin(j) = headerlist.First.Origin(j)
                End If
            Next

            Dim newlength As Double = Math.Round(headerlist(i - 1).Overhangbottom + headerlist(i - 1).Overhangtop + relevantcoords.Max - relevantcoords.Min, 3)
            headerlist(i - 1).Tube.Length = newlength

            'shorten xlist, ylist, idlist, anglelist, abvlist
            ReAssignListItems(headerlist(i - 1), relevantcoords, alignment, circuit, consys)
        Next
    End Sub

    Shared Function GetStartCoord(header As HeaderData, minvalue As Double, alignment As String) As Double
        Dim startpos As Double

        If alignment = "horizontal" Then
            startpos = Math.Round(minvalue - header.Overhangbottom, 2)
        Else
            startpos = Math.Round(minvalue - header.Overhangbottom, 2)
        End If

        Return startpos
    End Function

    Shared Function GetFirstStutzen(startpos As Double, valuelist As List(Of Double)) As Double
        Dim positionlist As New List(Of Double)
        Dim delta, firstposition As Double

        For i As Integer = 0 To valuelist.Count - 1
            delta = valuelist(i) - startpos
            If delta > 0 Then
                '> 0 → relevant for check - min value = start
                positionlist.Add(valuelist(i))
            End If
        Next

        firstposition = positionlist.Min

        Return firstposition
    End Function

    Shared Function GetHeaderframe(startpos As Double, maxlength As Double) As Double()
        Dim minvalue, maxvalue, headerframe() As Double

        minvalue = startpos
        maxvalue = startpos + maxlength

        headerframe = {minvalue, maxvalue}

        Return headerframe
    End Function

    Shared Function GetAllStutzen(headerframe() As Double, valuelist As List(Of Double), neededheadercount As Integer, alignment As String, number As Integer) As List(Of Double)
        Dim stutzenlist As New List(Of Double)
        Dim maxcount, modvalue As Integer

        modvalue = valuelist.Count Mod neededheadercount
        If alignment = "vertical" Then
            If modvalue = number Then
                maxcount = Math.Floor(valuelist.Count / neededheadercount)
            Else
                maxcount = Math.Ceiling(valuelist.Count / neededheadercount)
            End If
        Else
            If modvalue = number Then
                maxcount = Math.Ceiling(valuelist.Count / neededheadercount)
            Else
                maxcount = Math.Floor(valuelist.Count / neededheadercount)
            End If
        End If

        For i As Integer = 0 To valuelist.Count - 1
            If valuelist(i) >= headerframe(0) And valuelist(i) <= headerframe(1) And stutzenlist.Count < maxcount Then
                stutzenlist.Add(valuelist(i))
            End If
        Next

        Return stutzenlist
    End Function

    Shared Sub ReAssignListItems(ByRef header As HeaderData, relevantcoords As List(Of Double), alignment As String, circuit As CircuitData, consys As ConSysData)
        Dim donelist As New List(Of Integer)
        Dim newxlist, newylist, newabvlist, newanglelist, totalxlist, totalylist As New List(Of Double)
        Dim newidlist, staglist, ssidlist As New List(Of String)
        Dim newsdatalist As New List(Of StutzenData)

        For i As Integer = 0 To header.StutzenDatalist.Count - 1
            staglist.Add(header.StutzenDatalist(i).SpecialTag)
            totalxlist.Add(header.StutzenDatalist(i).XPos)
            totalylist.Add(header.StutzenDatalist(i).ZPos)
            ssidlist.Add(header.StutzenDatalist(i).ID)
            For j As Integer = 0 To relevantcoords.Count - 1
                Dim dojob As Boolean = False
                If alignment = "horizontal" Then
                    If header.StutzenDatalist(i).XPos = relevantcoords(j) AndAlso donelist.IndexOf(j) = -1 Then
                        dojob = True
                    End If
                Else
                    If header.StutzenDatalist(i).ZPos = relevantcoords(j) AndAlso donelist.IndexOf(j) = -1 Then
                        dojob = True
                    End If
                End If

                If dojob Then
                    newxlist.Add(header.StutzenDatalist(i).XPos)
                    newylist.Add(header.StutzenDatalist(i).ZPos)
                    newsdatalist.Add(header.StutzenDatalist(i))
                    donelist.Add(j)
                    Exit For
                End If
            Next
        Next

        With header
            .Xlist = newxlist
            .Ylist = newylist
            .StutzenDatalist = newsdatalist
        End With

        If General.currentunit.ModelRangeName = "GACV" Then
            If consys.SpecialCX Then
                If circuit.FinType = "E" Then
                    CSData.GACVSpecialCXE(header, circuit, consys)
                Else
                    CSData.GACVSpecialOutletN(header, circuit, consys)
                End If
            ElseIf consys.SpecialRX Then
                CSData.GACVSpecialRXF(header, circuit, consys)
            Else
                'can only be dx or co2 split → just outlet
                CSData.GACVSpecialOutlet(header, circuit, consys)
            End If
        Else
            'only GGDV needs special items in every header
            If circuit.Pressure > 100 And General.currentunit.UnitDescription = "VShape" Then
                If header.Displacever <> 0 Then
                    If header.Tube.HeaderType = "inlet" Then
                        CSData.GGDVSpecialInlet(header, circuit, consys)
                    Else
                        CSData.GGDVSpecialOutlet(header, circuit, consys)
                    End If
                End If
            Else
                If relevantcoords.Min < 100 Then
                    'transfer special
                    Dim specialentries As New Dictionary(Of Double, String)
                    Dim specialxloc, specialyloc As New List(Of Double)
                    Dim sidlist As New List(Of String)
                    For i As Integer = 0 To staglist.Count - 1
                        If staglist(i) <> "" Then
                            Dim coord As Double
                            If alignment = "horizontal" Then
                                coord = totalxlist(i)
                            Else
                                coord = totalylist(i)
                            End If
                            specialxloc.Add(totalxlist(i))
                            specialyloc.Add(totalylist(i))
                            sidlist.Add(ssidlist(i))
                            specialentries.Add(coord, staglist(i))
                        End If
                    Next
                    Dim delta As Double
                    If alignment = "horizontal" Then
                        delta = totalxlist.Max - newxlist.Max
                        For i As Integer = 0 To newxlist.Count - 1
                            Dim j As Integer = 0
                            For Each entry In specialentries
                                If entry.Key = newxlist(i) + delta Then
                                    If specialyloc(j) = newylist(i) Then
                                        newsdatalist(i).SpecialTag = entry.Value
                                        newsdatalist(i).ID = sidlist(i)
                                    End If
                                End If
                                j += 1
                            Next
                        Next
                    Else
                        delta = totalylist.Max - newylist.Max
                        For i As Integer = 0 To newylist.Count - 1
                            Dim j As Integer = 0
                            For Each entry In specialentries
                                If entry.Key = newylist(i) + delta Then
                                    If specialxloc(j) = newxlist(i) Then
                                        newsdatalist(i).SpecialTag = entry.Value
                                        newsdatalist(i).ID = sidlist(i)
                                    End If
                                End If
                                j += 1
                            Next
                        Next
                    End If

                End If
                With header
                    .Xlist = newxlist
                    .Ylist = newylist
                    .StutzenDatalist = newsdatalist
                End With
            End If
        End If

    End Sub

    Shared Sub NipplePosition(ByRef header As HeaderData, circuit As CircuitData, consys As ConSysData, no As Integer, finnedheight As Double)
        Dim position As Double

        If header.Tube.Length <= 99 Then
            position = header.Tube.Length / 2
        Else
            If header.Tube.IsBrine And consys.HeaderAlignment = "vertical" Then
                If header.Tube.HeaderType = "outlet" Then
                    position = header.Tube.Length - 93
                Else
                    position = 43
                    If header.Tube.Diameter > 76 Or (header.Tube.Diameter > 40 And header.Tube.Materialcodeletter = "V") Then
                        position = 93
                    End If
                End If
            Else
                If General.currentunit.ApplicationType = "Condenser" Then
                    If circuit.Pressure <= 16 Then
                        If General.currentunit.ModelRangeName.Substring(2, 2) = "HV" Or General.currentunit.ModelRangeName.Substring(2, 2) = "VV" Or General.currentunit.ModelRangeName.Substring(2, 2) = "DV" Then
                            position = DefaultWSPos(header, consys.HeaderAlignment, no - 1, consys.OutletNipples.First.Quantity, finnedheight)
                        Else
                            position = 80 + header.Tube.Diameter / 2 + (no - 1) * header.Tube.Length / consys.OutletNipples.Count
                        End If
                    Else
                        If General.currentunit.UnitDescription = "VShape" Then
                            If header.Tube.HeaderType = "inlet" Then
                                position = header.Tube.Length / (2 * consys.InletNipples.First.Quantity) + (no - 1) * header.Tube.Length / consys.InletNipples.First.Quantity
                            Else
                                position = Math.Round(Math.Max(header.Tube.WallThickness * 3, 10) + header.Tube.Diameter / 2 + (no - 1) * header.Tube.Length / consys.OutletNipples.First.Quantity)
                            End If
                        Else
                            If header.Tube.HeaderType = "inlet" Then
                                position = header.Tube.Length / (2 * consys.InletNipples.First.Quantity) + (no - 1) * header.Tube.Length / consys.InletNipples.First.Quantity
                            Else
                                position = header.Overhangbottom - circuit.CoreTube.Diameter / 2 + consys.OutletNipples.First.Diameter / 2 + (no - 1) * header.Tube.Length / consys.OutletNipples.First.Quantity
                            End If
                        End If
                    End If
                    If (header.Tube.HeaderType = "outlet" And consys.HeaderAlignment = "horizontal") Or (header.Tube.HeaderType = "inlet" And consys.HeaderAlignment = "vertical") Then
                        position = header.Tube.Length - position
                    End If
                Else
                    If consys.VType = "X" Then
                        Dim diameter As Double = consys.OutletNipples.First.Diameter
                        If consys.OutletNipples.First.Materialcodeletter = "D" Then
                            diameter = GNData.GetNippleDiameterK65(consys.OutletNipples.First.Diameter)
                        End If
                        If circuit.Pressure = 32 And header.Tube.Diameter <= 35 And header.Tube.Materialcodeletter = "C" Then
                            If circuit.FinType = "N" Or circuit.FinType = "M" Then
                                position = 4 + diameter / 2
                            Else
                                position = 12 + diameter / 2
                            End If
                        ElseIf circuit.Pressure < 80 And (circuit.CoreTube.Materialcodeletter = "C" Or ((circuit.FinType = "N" Or circuit.FinType = "M") And circuit.Pressure = 32)) Then
                            position = header.Overhangbottom - circuit.CoreTube.Diameter / 2 + diameter / 2

                        Else
                            position = DefaultDXPos(consys.OutletHeaders.First, circuit.Pressure, circuit.FinType, circuit.CoreTube.Materialcodeletter)
                        End If
                    Else
                        position = DefaultCSPos(header, consys.InletNipples.First.Diameter, circuit.Pressure, circuit.FinType, General.GetUniqueValues(header.Xlist).Count)
                        If General.currentunit.ModelRangeName = "GACV" And circuit.Pressure < 17 And no = 2 And consys.OutletNipples.First.Quantity = 2 Then
                            If header.Tube.HeaderType = "inlet" Then
                                position = header.Nipplepositions.First - header.Tube.Length / 2 - 150
                            Else
                                position = header.Nipplepositions.First + header.Tube.Length / 2 + 150
                            End If
                        End If
                    End If

                End If
            End If
        End If

        header.Nipplepositions.Add(position)
    End Sub

    Shared Function DefaultWSPos(header As HeaderData, alignment As String, no As Integer, quantity As Integer, finnedheight As Double) As Double
        Dim delta, offset As Double
        Dim position As Double

        Select Case quantity
            Case 1
                delta = 0
            Case 2
                If General.currentunit.UnitDescription = "VShape" Then
                    delta = 1200
                Else
                    delta = 1100
                End If
            Case 3
                delta = 800
            Case 4
                delta = 530
        End Select

        If General.currentunit.UnitDescription = "VShape" Then
            offset = 155
        Else
            offset = 150
        End If

        'If General.currentunit.UnitDescription <> "VShape" Then
        '    If alignment = "horizontal" Then
        '        If header.Tube.HeaderType = "inlet" Then
        '            position = offset - header.Origin(0) + no * delta
        '        Else
        '            position = header.Origin(0) + header.Tube.Length - finnedheight + offset + no * delta
        '        End If
        '    Else
        '        If header.Tube.HeaderType = "inlet" Then
        '            position = header.Origin(2) + header.Tube.Length - finnedheight + offset + no * delta
        '        Else
        '            position = offset - header.Origin(2) + no * delta
        '        End If
        '    End If
        'Else
        'End If
        position = offset + no * delta

        Return position
    End Function

    Shared Function DefaultCSPos(header As HeaderData, nipplediameter As Double, pressure As Integer, fintype As String, rows As Integer) As Double
        Dim position, m, overhang, k As Double
        Dim material As String = header.Tube.Materialcodeletter
        Dim incr As Integer

        If pressure <= 16 Then
            incr = GetIncrementCount(fintype, header.Tube.HeaderType, header.Tube.Diameter, rows)

            If header.Tube.HeaderType = "inlet" Then
                overhang = header.Overhangtop
                position = header.Tube.Length - 25 * incr - overhang
                'if cutout would overlap the tubeend, add one increment
                If position < nipplediameter + overhang Then
                    position += 25
                End If
            Else
                overhang = header.Overhangbottom
                position = 25 * incr + overhang
            End If
        Else
            If header.Tube.HeaderType = "inlet" Then
                'inlet AP / CP
                If material = "C" Then
                    If header.Tube.Diameter > 45 Then
                        m = 12.5
                    Else
                        m = 9
                    End If
                Else
                    If pressure = 32 Then
                        m = 10
                    Else
                        m = 11
                    End If
                End If
                position = m + nipplediameter / 2
            Else
                'outlet AP / CP
                If header.Tube.Materialcodeletter = "C" Then
                    If header.Tube.Diameter <= 42 Then
                        k = 43
                    Else
                        k = 47.5
                    End If
                ElseIf pressure = 54 Then
                    Select Case header.Tube.Diameter
                        Case 60.3
                            k = 73
                        Case 76.1
                            k = 75
                        Case 88.9
                            k = 127
                        Case Else
                            k = 43
                    End Select
                Else
                    If header.Tube.Diameter = 60.3 Then
                        k = 93
                    Else
                        k = 75
                    End If
                End If
                position = header.Tube.Length - k
            End If
        End If

        Return position
    End Function

    Shared Function DefaultDXPos(header As HeaderData, pressure As Integer, fintype As String, ctmaterial As String) As Double
        Dim k As Double

        Select Case pressure
            Case 32
                Select Case header.Tube.Diameter
                    Case 21.3
                        k = 21
                    Case 26.9
                        k = 21
                    Case 33.7
                        k = 23
                    Case 42.4
                        k = 27
                    Case 48.3
                        k = 31
                    Case 60.3
                        k = 34
                    Case 76.1
                        k = 42
                    Case 88.9
                        k = 50
                End Select
            Case 54
                Select Case header.Tube.Diameter
                    Case 21.3
                        k = 21
                    Case 26.9
                        k = 21
                    Case 33.7
                        k = 23
                    Case 42.4
                        k = 27
                    Case 48.3
                        k = 31
                    Case 60.3
                        If fintype = "E" Then
                            k = 39
                        Else
                            k = 34
                        End If
                    Case 76.1
                        k = 42
                End Select
            Case Else
                Select Case header.Tube.Diameter
                    Case 21.3
                        k = 21
                    Case 26.9
                        k = 21
                    Case 33.7
                        If ctmaterial = "C" Then
                            k = 27
                        Else
                            k = 23
                        End If
                    Case 42.4
                        k = 29
                    Case 48.3
                        k = 36
                    Case 60.3
                        k = 39
                End Select
        End Select

        Return k
    End Function

    Shared Function DefaultVShapePos(header As HeaderData, pressure As Integer) As Double
        Dim position As Double = 0

        If pressure > 100 Then
            Select Case header.Tube.Diameter
                Case 26.9
                    position = 20
                Case 33.7
                    position = 25
                Case 42.4
                    position = 30
                Case 48.3
                    position = 35
                Case 60.3
                    position = 40
            End Select
        Else
            If header.Tube.Materialcodeletter = "C" Then
                Select Case header.Tube.Diameter
                    Case 28
                        position = 24
                    Case 35
                        position = 28
                    Case 42
                        position = 31
                    Case 54
                        position = 38
                    Case 64
                        position = 45
                    Case 76.1
                        position = 51
                    Case 88.9
                        position = 59
                End Select
            Else
                Select Case header.Tube.Diameter
                    Case 26.9
                        position = 20
                    Case 33.7
                        position = 25
                    Case 42.4
                        position = 30
                    Case 48.3
                        position = 35
                    Case 60.3
                        position = 40
                    Case 76.1
                        position = 50
                    Case 88.9
                        position = 60
                End Select
            End If
        End If

        Return position
    End Function

    Shared Function GetIncrementCount(fintype As String, headertype As String, diameter As Double, rows As Integer) As Integer
        Dim finalinc, drem As Integer
        Dim gap, offset, flangediameter, mingap As Double

        flangediameter = GNData.GetFlangeDiameter(diameter)

        If headertype = "inlet" Then
            'calculation based on case with flange → only check with ceiling necessary
            If fintype = "F" Then
                offset = 16.5
            Else
                offset = 29
            End If
            mingap = 7
        Else
            offset = 10
            mingap = 1
        End If

        For i As Integer = 1 To 10
            gap = offset + i * 25 - flangediameter / 2
            If gap >= mingap Then
                finalinc = i
                Exit For
            End If
        Next

        Math.DivRem(finalinc, 2, drem)
        If fintype = "F" Then
            If headertype = "inlet" Then
                If rows = 2 And drem = 0 Then
                    'finalinc is even → +1
                    finalinc += 1
                ElseIf rows = 3 And drem > 0 Then
                    'finalinc is odd → +1
                    finalinc += 1
                End If
            Else
                If drem > 0 Then
                    'must be even
                    finalinc += 1
                End If
            End If
        Else
            'for N, inc must be odd
            If drem = 0 Then
                finalinc += 1
            End If
        End If

        Return finalinc
    End Function

    Shared Function VentWSPos(ByRef header As HeaderData, ctdiameter As Double) As Double
        Dim holediameter As Double = GNData.GetVentDiameter(header.Tube.Materialcodeletter, GNData.GetDryCoolerVentsize(header.Tube.Materialcodeletter, header.Tube.Diameter))
        Dim position As Double

        If header.Tube.HeaderType = "inlet" Then
            position = header.Tube.Length - (header.Overhangtop - ctdiameter / 2 + holediameter / 2)
        Else
            position = header.Overhangbottom + holediameter / 2 - ctdiameter / 2
        End If

        Return position
    End Function

    Shared Function VentGACVPos(ByRef header As HeaderData, circuit As CircuitData, RR As Integer) As Double
        Dim holediameter As Double = GNData.GetVentDiameter(header.Tube.Materialcodeletter, GNData.GetGACVVentsize(header, circuit, RR))
        Dim position As Double

        If header.Tube.HeaderType = "inlet" Then
            Dim offset As Double
            If circuit.FinType = "N" Or circuit.FinType = "M" Then
                If header.Tube.Materialcodeletter = "C" Then
                    If header.Tube.Diameter < 88 Then
                        offset = 16
                    Else
                        offset = 25
                    End If
                Else
                    If header.Tube.Diameter < 42 Then
                        offset = 18
                    ElseIf header.Tube.Diameter < 76 Then
                        offset = 22
                    Else
                        offset = 32
                    End If
                End If
            Else
                offset = 11
            End If
            position = header.Tube.Length - offset
        Else
            If circuit.FinType = "N" Or circuit.FinType = "M" Then
                If header.Tube.Materialcodeletter = "C" Then
                    If header.Tube.Diameter < 88 Then
                        position = 16
                    Else
                        position = 25
                    End If
                Else
                    If header.Tube.Diameter < 42 Then
                        position = 18
                    ElseIf header.Tube.Diameter < 76 Then
                        position = 22
                    Else
                        position = 32
                    End If
                End If
            Else
                If header.Tube.Materialcodeletter <> "C" Then
                    position = holediameter / 2 + 12
                Else
                    Select Case GNData.GetGACVVentsize(header, circuit, RR)
                        Case "G3/8"
                            position = 10
                        Case "G1/2"
                            position = 12
                        Case Else
                            position = 18
                    End Select
                End If
            End If
        End If

        Return position
    End Function

    Shared Function CoilGap(coillist As List(Of CoilData)) As Double
        Dim gap As Double = 0

        Try
            If General.currentunit.UnitDescription <> "VShape" Then
                gap = coillist.First.FinnedDepth
            Else
                If coillist.First.Circuits.First.ConnectionSide = "left" Then
                    gap = 1000
                Else
                    gap = -1000
                End If
            End If
        Catch ex As Exception

        End Try
        Return gap
    End Function

    Shared Function ControlNippleLengthGACV(nippletube As TubeData, contype As Integer, stutzencoords As List(Of Double), connectionside As String, displacement As Double, finneddepth As Double)
        Dim controllength As Double = nippletube.Length
        Dim x, mp, values() As Double

        Try
            If (connectionside = "left" And Not nippletube.IsBrine) Or (connectionside = "right" And nippletube.IsBrine) Then
                x = finneddepth - stutzencoords.Max
                mp = -1
            Else
                x = stutzencoords.Min
                mp = 1
            End If

            'if mirrored, displacement turns into positive value, but the nipple doesnt become longer!
            controllength = 90 + General.currentunit.TubeSheet.Dim_d + x + displacement * mp

            If nippletube.Materialcodeletter = "D" Then
                controllength += GNData.GetAdapterOffset(GNData.GetAdapterID(nippletube.Diameter), "") - GNData.GetAdapterOffset(GNData.GetAdapterID(nippletube.Diameter), "adapter")
            End If

            If contype > 1 And nippletube.HeaderType = "outlet" Then
                controllength += 70
            ElseIf contype = 1 Then
                values = GNData.ThreadData(PCFData.GetValue("ThreadFlangeConnectionFC1", "ERPCode"))
                controllength -= values(1)
            End If
            If contype = 3 Then
                values = GNData.WeldedData(PCFData.GetValue("ThreadFlangeConnectionFC1", "ERPCode"))
                controllength = controllength - values(0) + values(1)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return controllength
    End Function

    Shared Function ControlNippleLengthVShape(currentlength As Double, totallength As Double, contype As Integer, header As HeaderData) As Double
        Dim controllength As Double = currentlength

        Try
            controllength = totallength - header.Dim_a - header.Tube.Diameter / 2
        Catch ex As Exception

        End Try

        Return controllength
    End Function

    Shared Function GetStutzen5Alignment(l1sp As Double, l2sp As Double, abv As Double, angle As Double, a As Double, headerdia As Double, headerwt As Double, ctoverhang As Double) As String
        Dim xH, yH, yS, l1, l2, l2dash, ycheck As Double
        Dim acalcvalues() As Double = {a, a}
        Dim abvcalvalues() As Double = {abv, abv}
        Dim alignment As String

        For i As Integer = 1 To 2
            'run 2 loops, fig 5 stutzen can be switched around
            If i = 1 Then
                l1 = l1sp
                l2 = l2sp
            Else
                l1 = l2sp
                l2 = l1sp
            End If

            'calculating the distance between header hole and actual position of the stutzen 
            xH = Math.Round(headerdia / 2 + a - ctoverhang + 5 - headerdia / 2 * Math.Cos(angle * Math.PI / 180), 3)
            yH = Math.Round(abv - headerdia / 2 * Math.Sin(angle * Math.PI / 180), 3)

            'l2dash = distance between bending center of stutzen and contact point with header
            If angle = 90 Then
                ycheck = Math.Abs(xH - l1)
                l2dash = yH
            Else
                l2dash = Math.Round((xH - l1) / Math.Cos(angle * Math.PI / 180), 3)
                yS = Math.Round(l2dash * Math.Sin(angle * Math.PI / 180), 3)
                ycheck = Math.Abs(yH - yS)
            End If

            If l2 > (l2dash + headerwt) And l2 < (l2dash + headerdia / 2) And ycheck < 3 And l2dash > (headerwt + 25 * Math.Sin(angle * Math.PI / 180)) Then
                acalcvalues(i - 1) = Math.Abs(ycheck)
            End If
        Next

        If acalcvalues(0) = acalcvalues.Min Then
            alignment = "normal"
        Else
            alignment = "reverse"
        End If

        Return alignment
    End Function

    Shared Function MaxDVcount(objDV As SolidEdgeDraft.DrawingView, finneddepth As Double) As Integer
        Dim xmin, ymin, xmax, ymax As Double
        Dim n As Integer

        objDV.Range(xmin, ymin, xmax, ymax)

        If General.currentunit.ModelRangeName = "GACV" Then
            n = Math.Floor((0.59 - 0.03) / (finneddepth / 1000 + 0.127) / objDV.ScaleFactor)
        Else
            n = Math.Floor((0.41 - 0.07) / (finneddepth / 1000 + 0.127) / objDV.ScaleFactor)
        End If
        Return n
    End Function

    Shared Function RescaleFactor(scalefactor As Double, ymax As Double, scalemode As String, limit As Double) As Double
        Dim scalefactorlist As New List(Of Double) From {0.1, 0.15, 0.2, 0.25, 0.3, 0.333, 0.4, 0.5, 1}
        Dim newsf As Double

        If scalemode = "down" Then
            If Math.Round(2 * ymax, 6) > limit Then
                If scalefactor <= scalefactorlist.Min Then
                    newsf = 0.9 * scalefactor
                Else
                    newsf = scalefactorlist(scalefactorlist.IndexOf(scalefactor) - 1)
                End If
            Else
                newsf = scalefactor
            End If
        Else
            If scalefactor = scalefactorlist.Max Then
                newsf = 1.1 * scalefactor
            Else
                newsf = scalefactorlist(scalefactorlist.IndexOf(scalefactor) + 1)
            End If
        End If

        Return newsf
    End Function

    Shared Function GetRotationDir(coil As CoilData, conside As String) As Integer
        Dim rotation As Integer

        '0 = no rotation
        '1 = rotation ccw / right
        '-1 = rotation cw / left
        If coil.Alignment = "horizontal" Then
            rotation = 0
        Else
            If conside = "left" Then
                rotation = -1
            Else
                rotation = 1
            End If
            If General.currentunit.UnitDescription = "VShape" Then
                rotation *= -1
            End If
        End If

        Return rotation
    End Function

    Shared Sub GetFlangeLength(ByRef consys As ConSysData)
        Dim partdoc As SolidEdgePart.PartDocument
        Dim flanschrevpro As SolidEdgePart.RevolvedProtrusion = Nothing
        Dim objprofile As SolidEdgePart.Profile
        Dim objmodel As SolidEdgePart.Model
        Dim filename As String = ""
        Dim x1, y1, x2, y2, f2 As Double
        Dim xlist As New List(Of Double)

        Try
            If consys.FlangeID <> "" Then
                filename = General.GetFullFilename(General.currentjob.Workspace, consys.FlangeID, "par")
                If filename <> "" Then
                    partdoc = General.seapp.Documents.Open(filename)
                    General.seapp.DoIdle()

                    consys.FlangeDims.HF = SEPart.GetSetVariableValue("HF", partdoc.Variables, "get")
                    consys.FlangeDims.sB = SEPart.GetSetVariableValue("sB", partdoc.Variables, "get")
                    consys.FlangeDims.DF = SEPart.GetSetVariableValue("DF", partdoc.Variables, "get")
                    f2 = SEPart.GetSetVariableValue("f2", partdoc.Variables, "get")

                    objmodel = partdoc.Models.Item(1)

                    For Each revpro As SolidEdgePart.RevolvedProtrusion In objmodel.RevolvedProtrusions
                        If revpro.Name = "Flansch" Then
                            flanschrevpro = revpro
                        End If
                    Next
                    objprofile = flanschrevpro.Profile

                    For Each l As SolidEdgeFrameworkSupport.Line2d In objprofile.Lines2d
                        If l.Style.LinearName.Contains("Solid") Then
                            l.GetStartPoint(x1, y1)
                            l.GetEndPoint(x2, y2)
                            xlist.Add(Math.Round(x1, 6))
                            xlist.Add(Math.Round(x2, 6))
                        End If
                    Next

                    consys.FlangeDims.Length = f2 + Math.Abs(xlist.Max * 1000)
                End If
            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        Finally
            If filename <> "" Then
                General.seapp.Documents.CloseDocument(filename, False, DoIdle:=True)
            End If
        End Try

    End Sub

    Shared Function MirrorCoordsFromList(valuelist() As List(Of Double), finneddepth As Double, numbers As List(Of Integer)) As List(Of Double)()
        Dim newvaluelist(3) As List(Of Double)
        For j As Integer = 0 To valuelist.Count - 1
            If numbers.IndexOf(j) > 0 Then
                For i As Integer = 0 To valuelist(0).Count
                    newvaluelist(j).Add(Math.Round(finneddepth - valuelist(j)(i), 3))
                Next
            Else
                newvaluelist(j) = valuelist(j)
            End If
        Next
        Return newvaluelist
    End Function

    Shared Function ConvertToPointlist(lists As List(Of List(Of Double)())) As List(Of PointData)
        Dim plist As New List(Of PointData)

        'which front + back + in/out point appears only once
        For j As Integer = 0 To lists.Count - 1
            For i As Integer = 0 To lists(j)(0).Count - 1
                plist.Add(New PointData With {.X = lists(j)(0)(i), .Y = lists(j)(1)(i)})
                plist.Add(New PointData With {.X = lists(j)(2)(i), .Y = lists(j)(3)(i)})
            Next
        Next
        Return plist
    End Function

    Shared Function GetUniquePoint(checklist As List(Of PointData), plist As List(Of PointData)) As PointData
        Dim flist As List(Of PointData)
        Dim p0 As New PointData
        Try
            For Each p As PointData In checklist
                flist = plist.FindAll(Function(s) s.X.Equals(p.X) And s.Y.Equals(p.Y))
                If flist.Count = 1 Then
                    p0 = p
                    Exit For
                End If
            Next
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
        Return p0
    End Function

    Shared Sub AddGADCHeader(ByRef headerlist As List(Of HeaderData), pressure As Integer)
        headerlist.Add(New HeaderData)

        headerlist(1).Tube.TubeFile.Fullfilename = General.currentjob.Workspace + "\" + General.TextUpperCase(headerlist.First.Tube.HeaderType) + "Header2_1_1.par"
        headerlist(1).Tube.FileName = headerlist(1).Tube.TubeFile.Fullfilename
        headerlist(1).Tube.TubeFile.Shortname = General.GetShortName(headerlist(1).Tube.FileName)

        With headerlist
            .Last.Dim_a = .First.Dim_a
            .Last.Displacehor = .First.Displacehor
            .Last.Displacever = .First.Displacever
            .Last.Nippletubes = 1
            .Last.Origin = {0, 0, 0}
            .Last.Overhangbottom = .First.Overhangbottom
            .Last.Overhangtop = .First.Overhangtop
            .Last.Tube.Diameter = .First.Tube.Diameter
            .Last.Tube.HeaderType = .First.Tube.HeaderType
            .Last.Tube.Materialcodeletter = .First.Tube.Materialcodeletter
            .Last.Tube.RawMaterial = .First.Tube.RawMaterial
            .Last.Tube.WallThickness = .First.Tube.WallThickness
        End With

        SEPart.CreateTube(headerlist.Last.Tube, headerlist.Last.Tube.FileName, "header", pressure)

    End Sub
End Class
