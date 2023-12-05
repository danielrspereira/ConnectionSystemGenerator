Public Class GNData

    Shared Function GetTubeDiameter(fin As String) As Double
        Dim diameter As Double
        Select Case fin
            Case "D"
                diameter = 7
            Case "E"
                diameter = 9.52
            Case "F"
                diameter = 12
            Case "G"
                diameter = 15
            Case "H"
                diameter = 9.52
            Case "K"
                diameter = 22
            Case "M"
                diameter = 12
            Case "N"
                diameter = 15
            Case "S"
                diameter = 22
        End Select

        Return diameter
    End Function

    Shared Function GetMaterialcode(material As String, parttyp As String) As String
        Dim materialcode As String = ""

        If parttyp = "bow" Then
            If material.Contains("R") Then
                materialcode = "C"
            ElseIf material.Contains("X") Then
                materialcode = "C"
            ElseIf material.Contains("V4") Then
                materialcode = "W"
            Else
                materialcode = material.Substring(0, 1)
            End If
        ElseIf parttyp = "tube" Then
            If material.Contains("V2A") Then
                materialcode = "V"
            ElseIf material.Contains("V4A") Then
                materialcode = "W"
            ElseIf material.Contains("X") Then
                materialcode = "X"
            ElseIf material.Contains("R") Then
                materialcode = "R"
            ElseIf material.Contains("K65") Then
                materialcode = "K"
            Else
                materialcode = "C"
            End If
        ElseIf parttyp = "coretube" Then
            If material.Contains("V2A") Then
                materialcode = "V"
            ElseIf material.Contains("V4A") Then
                materialcode = "W"
            ElseIf material.Contains("X") Then
                materialcode = "X"
            ElseIf material.Contains("R") Then
                materialcode = "R"
            Else
                materialcode = "C"
            End If
        End If

        Return materialcode
    End Function

    Shared Function GetFinPitch(coilposition As String, fin As String) As Double()
        Dim pitchx, pitchy, finpitch() As Double

        pitchx = 25
        pitchy = 25

        Select Case coilposition
            Case "vertical"
                If fin = "D" Or fin = "H" Then
                    pitchx = 21.65
                    pitchy = 12.5
                ElseIf fin = "K" Then
                    pitchx = 51.96
                    pitchy = 30
                End If
            Case "horizontal"
                If fin = "D" Or fin = "H" Then
                    pitchx = 12.5
                    pitchy = 21.65
                ElseIf fin = "K" Then
                    pitchx = 30
                    pitchy = 51.96
                End If
        End Select

        If fin = "N" Or fin = "M" Then
            pitchx = 50
            pitchy = 50
        End If

        finpitch = {pitchx, pitchy}

        Return finpitch
    End Function

    Shared Function GetBowIDs(frontbowprops() As List(Of Double), backbowprops() As List(Of Double), material As String, fintype As String, pressure As Integer, orbitalwelding As Boolean) As List(Of String)()
        Dim l1list, wtlist, l1levels, firstlevel As New List(Of Double)
        Dim bowkey, spezi, materialcode, bow() As String
        Dim bowkeylist, uniquekeys, keylist, bowidlist, frontbowkeys, backbowkeys, frontids, backids, uniqueids, lnlevels, lengthlist As New List(Of String)
        Dim bowids() As List(Of String)
        Dim diameter, wallthickness, refl1 As Double
        Dim bowlevel As Integer = 2
        Dim bowindex, keyindex, maxlevel As Integer
        Dim flatten As Boolean = False
        Dim hasdiagonal As Boolean = False
        Dim bowlist As New List(Of BowData)
        Dim totalbowlist As New List(Of List(Of BowData))
        Dim currentbow As BowData

        Try
            If fintype = "E" Or fintype = "F" Or fintype = "N" Or fintype = "M" Then
                flatten = True
            End If

            If frontbowprops IsNot Nothing Then     'If passnumber = 2 → no bows on the front
                frontbowkeys = CircProps.CreateBowkeys(frontbowprops)
            End If

            If backbowprops IsNot Nothing Then
                backbowkeys = CircProps.CreateBowkeys(backbowprops)
            End If

            uniquekeys = General.GetUniqueStrings(frontbowkeys, backbowkeys)

            'Starting with level 1 bows
            uniquekeys.Sort()
            maxlevel = uniquekeys.Last.Substring(0, 1)

            'Gathering data
            diameter = GetTubeDiameter(fintype)
            spezi = GetSpecification(pressure, material.Substring(0, 2), "bow", fintype)
            materialcode = GetMaterialcode(material, "bow")

            If materialcode = "V" And fintype = "E" And material.Contains("V2") Then
                materialcode = "W"
            End If

            wallthickness = Database.GetTubeThickness("Bow", diameter.ToString, materialcode, pressure)

            'creating a list for each uniquekey
            For i As Integer = 0 To uniquekeys.Count - 1
                bow = uniquekeys(i).Split({"\"}, 0)
                bowids = Database.GetBowID(bow(2), diameter, wallthickness, spezi, bow(1), orbitalwelding)
                If bowids(0).Count > 0 Then
                    For j As Integer = 0 To bowids(0).Count - 1
                        Dim newbow As New BowData With {.ID = bowids(0)(j), .L1 = bowids(1)(j), .Length = bow(1), .Wallthickness = bowids(2)(j), .Level = bow(0), .Uniquekey = uniquekeys(i)}
                        If totalbowlist.Count > i Then
                            totalbowlist(i).Add(newbow)
                        Else
                            totalbowlist.Add(New List(Of BowData) From {newbow})
                        End If
                    Next
                Else
                    Dim newbow As New BowData With {.ID = "", .L1 = 0, .Length = bow(1), .Wallthickness = wallthickness, .Level = bow(0), .Uniquekey = uniquekeys(i)}
                    'Add template bow for adjustment
                    If bow(2) = "9" Then
                        newbow.ID = Library.TemplateParts.BOW9
                    ElseIf orbitalwelding Then
                        newbow.ID = Library.TemplateParts.ORIBITALBOW1
                    Else
                        newbow.ID = Library.TemplateParts.BOW1
                    End If
                    newbow.ID = "0000" + newbow.ID
                    If totalbowlist.Count > i Then
                        totalbowlist(i).Add(newbow)
                    Else
                        totalbowlist.Add(New List(Of BowData) From {newbow})
                    End If
                End If
            Next

            Dim simplekey = From allkeys In uniquekeys Where allkeys.Substring(0, 1) = "1"

            For Each level1keys In simplekey
                keyindex = uniquekeys.IndexOf(level1keys)
                'all bows for this key
                bowlist = totalbowlist(keyindex)
                'get the smallest bow for this key
                Dim sortedbows = From sortedlist In bowlist Order By sortedlist.L1

                For Each l1bow In sortedbows
                    If General.IntegerMod(l1bow.Length, 50) <> 0 Then
                        hasdiagonal = True
                    End If
                Next
                If sortedbows(0).L1 = 0 Then
                    firstlevel.Add(Math.Min(Math.Round(sortedbows(0).Length / 2 + 20, 1), 50))
                Else
                    firstlevel.Add(sortedbows(0).L1)
                End If
            Next

            'L1 of first level
            If flatten And hasdiagonal = False Then
                lnlevels.Add(firstlevel.Min)
            Else
                lnlevels.Add(firstlevel.Max)
            End If

            If maxlevel > 1 Then
                Do
                    l1levels.Clear()
                    refl1 = lnlevels.Max + 10 + diameter

                    Dim nextkeys = From allkeys In uniquekeys Where allkeys.Substring(0, 1) = bowlevel.ToString

                    For Each uniquekey In nextkeys
                        keyindex = uniquekeys.IndexOf(uniquekey)
                        'all bows for this key
                        bowlist = totalbowlist(keyindex)
                        'get the smallest bow this key
                        Dim sortedbows = From sortedlist In bowlist Order By sortedlist.L1

                        For i As Integer = 0 To sortedbows.Count - 1
                            If sortedbows(i).L1 >= refl1 And sortedbows(i).L1 <= refl1 + 5 Then
                                l1levels.Add(sortedbows(i).L1)
                            End If
                        Next
                    Next

                    If l1levels.Count > 0 Then
                        lnlevels.Add(l1levels.Min)
                    Else
                        lnlevels.Add(refl1)
                    End If
                    bowlevel += 1
                Loop Until lnlevels.Count = maxlevel
            End If

            'now check for each key if a bow has the L1 value
            For i As Integer = 0 To uniquekeys.Count - 1
                Dim uid As String
                bowkey = uniquekeys(i)
                bowlist = totalbowlist(i)
                Dim tempbows = From allbows In bowlist Order By allbows.L1

                If bowkey.Substring(0, 1) = "1" Then
                    currentbow = tempbows(0)
                    'check if vertical or horizontal bow with "too high" L1
                    If flatten And General.IntegerMod(currentbow.Length, 50) = 0 And bowkey.Substring(bowkey.LastIndexOf("\") + 1) <> "9" Then
                        If currentbow.L1 <= lnlevels(0) + 1 Then
                            uid = currentbow.ID
                        Else
                            uid = Library.TemplateParts.BOW1
                            uid = "0000" + uid
                        End If
                        uniqueids.Add(uid)
                    Else
                        uniqueids.Add(currentbow.ID)
                    End If
                Else
                    bowlevel = bowkey.Substring(0, 1) - 1
                    For j As Integer = 0 To tempbows.Count - 1
                        currentbow = tempbows(j)
                        If currentbow.L1 = lnlevels(bowlevel) Then
                            uniqueids.Add(currentbow.ID)
                            j = tempbows.Count - 1
                        End If
                    Next
                    If uniqueids.Count - 1 < i Then
                        If bowkey.Substring(bowkey.LastIndexOf("\") + 1) = "9" Then
                            uid = Library.TemplateParts.BOW9
                        ElseIf orbitalwelding Then
                            uid = Library.TemplateParts.ORIBITALBOW1
                        Else
                            uid = Library.TemplateParts.BOW1
                        End If
                        uniqueids.Add("0000" + uid)
                    End If
                End If
            Next

            If uniqueids.Count <> uniquekeys.Count Then

            Else
                'Assign the bow to front and back list
                If frontbowprops IsNot Nothing Then
                    For i As Integer = 0 To frontbowkeys.Count - 1
                        bowkey = frontbowkeys(i)
                        bowindex = uniquekeys.IndexOf(bowkey)
                        frontids.Add(uniqueids(bowindex))
                    Next
                End If

                For i As Integer = 0 To backbowkeys.Count - 1
                    bowkey = backbowkeys(i)
                    bowindex = uniquekeys.IndexOf(bowkey)
                    backids.Add(uniqueids(bowindex))
                Next

            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        'Returns a list of article numbers for each bow + the max L1 for each level 
        bowids = {frontids, backids, lnlevels}

        Return bowids
    End Function

    Shared Function GetSpecification(pressure As Integer, material As String, coretype As String, fintype As String) As String
        Dim spezi As String
        Dim shortmaterial As String

        If material.Length > 2 Then
            shortmaterial = material.Substring(0, 2)
        Else
            shortmaterial = material
        End If

        If shortmaterial = "CU" Or shortmaterial = "C" Then
            If fintype = "E" Then
                If pressure >= 80 Then
                    spezi = "SP01-A1"
                Else
                    spezi = "SP01-A"
                End If
            Else
                spezi = "SP01-A"
            End If
        ElseIf material = "V2" And coretype <> "CT" Then
            spezi = "SP03-1"
        Else
            spezi = "SP03-2"
        End If

        Return spezi
    End Function

    Shared Function GetStutzenID(abv As Double, header As HeaderData, circuit As CircuitData, consys As ConSysData, figure As Integer, displacever As Double,
                                 ByRef sdata As Dictionary(Of Double, Double), Optional gap As Double = 0, Optional anglemp As Integer = 1) As String()
        Dim stutzendata(), idlist As List(Of String)
        Dim specification As String = GetSpecification(circuit.Pressure, circuit.CoreTube.Material, "CT", circuit.FinType)
        Dim l1list, l2list, anglelist, offsetlist As List(Of Double)
        Dim deltalist As New List(Of Double)
        Dim delta, angle, minwallthickness As Double
        Dim index As Integer
        Dim holeposition As String = ""
        Dim stutzenID As String = ""
        Dim materialcode As String

        Try
            If figure <> 0 Then
                materialcode = GetMaterialcode(circuit.CoreTube.Material, "bow")

                If materialcode = "V" And circuit.CoreTube.Diameter = 9.52 And circuit.CoreTube.Material.Contains("V2") Then
                    materialcode = "W"
                End If

                minwallthickness = Database.GetTubeThickness("Stub", circuit.CoreTube.Diameter, materialcode, circuit.Pressure)

                'only checking the normal solution (figure 4, 5, 8) // S* stutzen in a different function to replace specific items at position x/y
                stutzendata = Database.GetStutzenData(figure.ToString, circuit.CoreTube.Diameter, specification, minwallthickness, circuit.Pressure)
                idlist = stutzendata(0)
                l1list = General.ConvertList(stutzendata(1))

                If figure <> 8 Then
                    l2list = General.ConvertList(stutzendata(2))
                    anglelist = General.ConvertList(stutzendata(3))
                    If figure <> 5 Then
                        offsetlist = General.ConvertList(stutzendata(4))
                    End If
                End If

                If idlist.Count = 0 Then
                    deltalist.Add(header.Dim_a)
                End If

                Select Case figure
                    Case 4
                        'compare only a and abv
                        For i As Integer = 0 To idlist.Count - 1
                            If Math.Abs(Math.Abs(abv) - offsetlist(i)) < 0.65 Then
                                delta = Calculation.CompareAFig4(l1list(i), header.Dim_a - gap, header.Tube.Diameter, header.Tube.WallThickness, circuit.CoreTubeOverhang, abv, offsetlist(i))
                                deltalist.Add(delta)
                            Else
                                deltalist.Add(header.Dim_a)
                            End If
                        Next
                        'delta = how deep it goes into the header
                        If deltalist.Min <= header.Tube.Diameter / 2 - header.Tube.WallThickness Then
                            index = deltalist.IndexOf(deltalist.Min)
                            stutzenID = idlist(index)
                            angle = 0
                        Else
                            stutzenID = Library.TemplateParts.STUTZEN4
                            angle = 0
                        End If
                    Case 5
                        'check angle 

                        Dim minangle As Double
                        If sdata.Count > 0 Then
                            minangle = CompareAngleFig5(sdata)
                        End If

                        For i As Integer = 0 To idlist.Count - 1
                            If circuit.CircuitType.ToLower.Contains("defrost") Then
                                If anglelist(i) = 90 Then
                                    delta = Calculation.CompareABVFig5(l1list(i), l2list(i), Math.Abs(abv), anglelist(i), header.Dim_a, header.Tube.Diameter, header.Tube.WallThickness, circuit.CoreTubeOverhang, 5)
                                Else
                                    delta = header.Dim_a
                                End If
                            Else
                                If anglelist(i) >= minangle Then
                                    delta = Calculation.CompareABVFig5(l1list(i), l2list(i), Math.Abs(abv), anglelist(i), header.Dim_a, header.Tube.Diameter, header.Tube.WallThickness, circuit.CoreTubeOverhang, 5)
                                Else
                                    delta = header.Dim_a
                                End If
                            End If
                            deltalist.Add(delta)
                        Next
                        If deltalist.Min < 5 Then
                            index = deltalist.IndexOf(deltalist.Min)
                            stutzenID = idlist(index)
                            angle = anglelist(index)
                        Else
                            stutzenID = Library.TemplateParts.STUTZEN5

                            'default angle based on abv rank / bigger than angle for smaller abv
                            'get current max angle (only figure 5 has an angle, 4 and 8 are 0)
                            'value is only starting value for exact calculation later

                            Dim angles() As Double = {0, 16, 32, 48, 60, 90}
                            If sdata.Count > 0 Then
                                For i As Integer = 0 To angles.Count - 1
                                    For Each entry In sdata
                                        If angles(i) > entry.Value And angle = 0 Then
                                            angle = angles(i)
                                            Exit For
                                        End If
                                    Next
                                    If angle > 0 Then
                                        Exit For
                                    End If
                                Next
                            Else
                                angle = 16
                            End If
                        End If
                    Case 8
                        If (circuit.CoreTube.Materialcodeletter = "V" Or circuit.CoreTube.Materialcodeletter = "W") And consys.HeaderAlignment = "horizontal" And header.Displacever > 0 And General.currentunit.ModelRangeName = "GACV" Then
                            holeposition = "tangential"
                        End If
                        'compare only a
                        For i As Integer = 0 To idlist.Count - 1
                            delta = Calculation.CompareAFig8(l1list(i), header.Dim_a, header.Tube.WallThickness, circuit.CoreTubeOverhang, holeposition)
                            If delta > 0 Then
                                deltalist.Add(delta)
                            Else
                                deltalist.Add(l1list(i))
                            End If
                        Next
                        index = deltalist.IndexOf(deltalist.Min)
                        If deltalist.Min < header.Tube.Diameter / 2 Then
                            stutzenID = idlist(index)
                        Else
                            stutzenID = Library.TemplateParts.STUTZEN8
                        End If
                        angle = 0
                    Case 45
                        For i As Integer = 0 To idlist.Count - 1
                            Dim correctangle As Boolean = False
                            If circuit.NoPasses = 3 And header.Tube.HeaderType = "outlet" Then
                                If circuit.ConnectionSide = "left" Then
                                    If anglelist(i) > 0 Then
                                        correctangle = True
                                    End If
                                Else
                                    If anglelist(i) < 0 Then
                                        correctangle = True
                                    End If
                                End If
                            Else
                                If circuit.ConnectionSide = "left" Then
                                    If anglelist(i) * anglemp < 0 Then
                                        correctangle = True
                                    End If
                                Else
                                    If anglelist(i) * anglemp > 0 Then
                                        correctangle = True
                                    End If
                                End If
                            End If
                            If Math.Abs(Math.Abs(displacever) - offsetlist(i)) < 0.5 And correctangle Then
                                delta = Calculation.CompareABVFig5(l1list(i), l2list(i), Math.Abs(abv), Math.Abs(anglelist(i)), header.Dim_a, header.Tube.Diameter, header.Tube.WallThickness, circuit.CoreTubeOverhang, 45)
                                deltalist.Add(delta)
                            Else
                                deltalist.Add(header.Dim_a)
                            End If
                        Next
                        If deltalist.Min < 10 Then
                            index = deltalist.IndexOf(deltalist.Min)
                            stutzenID = idlist(index)
                            angle = anglelist(index)
                        Else
                            If header.Tube.HeaderType = "inlet" Then
                                If sdata.Count = 1 Then
                                    angle = 45
                                    stutzenID = Library.TemplateParts.STUTZEN45OUT
                                Else
                                    angle = 90
                                    stutzenID = Library.TemplateParts.STUTZEN45IN
                                End If
                            Else
                                angle = 90
                                stutzenID = Library.TemplateParts.STUTZEN45OUT
                            End If
                        End If
                End Select
            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return {stutzenID, angle}
    End Function

    Shared Function CompareAngleFig5(sdata As Dictionary(Of Double, Double)) As Double
        Dim minangle As Double
        Dim angles() As Double = {0, 16, 32, 48, 60, 90}

        If sdata.Count > 0 Then
            For i As Integer = 0 To angles.Count - 1
                For Each entry In sdata
                    If angles(i) > entry.Value And minangle = 0 Then
                        minangle = angles(i)
                        Exit For
                    End If
                Next
                If minangle > 0 Then
                    Exit For
                End If
            Next
        Else
            minangle = 16
        End If
        Return minangle
    End Function

    Shared Function DefineFigure(alignment As String, circuit As CircuitData, header As HeaderData, abvlist As List(Of Double), abv As Double, overridefig As Boolean, norows As Integer) As Integer
        Dim fixedfigure As Integer = 0
        Dim firstcolumn As Boolean

        If alignment = "horizontal" And General.currentunit.ModelRangeName = "GACV" And (circuit.CoreTube.Materialcodeletter = "V" Or circuit.CoreTube.Materialcodeletter = "W") And circuit.CircuitType.Contains("Defrost") = False Then
            'for horizontal AP and certain diameters, type 4 stutzen needed
            If abvlist.Count > 1 Then
                fixedfigure = GetXPStutzentype(header.Tube.Diameter, abv, header.Tube.HeaderType)
            Else
                If header.Tube.HeaderType = "outlet" And header.Tube.Diameter >= 21.3 And circuit.NoDistributions > 1 Then
                    fixedfigure = 4
                End If
            End If
        ElseIf General.currentunit.ModelRangeName = "GCDV" And (circuit.CoreTube.Materialcodeletter = "V" Or circuit.CoreTube.Materialcodeletter = "W") Then
            If norows = 4 And abv <> 0 Then
                fixedfigure = 4
            ElseIf norows = 6 Then
                If Math.Abs(abv) = 25 Then
                    fixedfigure = 4
                ElseIf Math.Abs(abv) > 25 Then
                    fixedfigure = 5
                End If
            End If
        End If

        If fixedfigure = 0 Then
            If abv = 0 Then
                fixedfigure = 8
            ElseIf General.currentunit.ModelRangeSuffix = "AD" Or (General.currentunit.ApplicationType = "Condenser" And circuit.CoreTube.Materialcodeletter = "V") Then
                fixedfigure = CheckNH3Column(abvlist, abv)
            Else
                If General.currentunit.ModelRangeName = "GACV" Then
                    If circuit.FinType = "F" Or circuit.FinType = "E" Then
                        'check column
                        firstcolumn = CheckColumn(abvlist, abv, circuit.CoreTube.Materialcodeletter, circuit.ConnectionSide)
                    ElseIf header.Displacever <> 0 And alignment = "vertical" And General.currentunit.ModelRangeSuffix <> "AP" And (General.currentunit.ModelRangeSuffix <> "CP" Or header.Tube.HeaderType <> "outlet") Then
                        firstcolumn = CheckColumn(abvlist, abv, circuit.CoreTube.Materialcodeletter, circuit.ConnectionSide)
                    End If
                Else
                    firstcolumn = False
                End If
                If Not overridefig And firstcolumn And header.Displacehor <> 0 Then
                    'FP second row
                    fixedfigure = 4
                Else
                    fixedfigure = 5
                End If
            End If
        End If

        Return fixedfigure
    End Function

    Shared Function GetXPStutzentype(headerdiameter As Double, currentabv As Double, headertype As String) As Integer
        Dim apfigure As Integer = 0

        'latest decision: always figure 4 for middle row
        If headerdiameter > 26 Then
            If Math.Abs(currentabv) = 50 Or Math.Abs(currentabv) = 55.9 Then
                apfigure = 4
            End If
        End If
        'tangential for inlet lowest
        If headerdiameter = 60.3 And headertype = "inlet" And Math.Abs(currentabv) = 10 Then
            apfigure = 8
        End If

        Return apfigure
    End Function

    Shared Function CheckColumn(abvlist As List(Of Double), abv As Double, materialcode As String, conside As String) As Boolean
        Dim firstcolumn As Boolean = False
        Dim containsnull As Boolean = False

        For Each value In abvlist
            If value = 0 Then
                If abv < 0 Then
                    firstcolumn = True
                End If
                containsnull = True
            End If
        Next

        If containsnull = False Then
            If (materialcode = "V" Or materialcode = "W") And General.currentunit.ApplicationType = "Condenser" Then
                firstcolumn = True
            Else
                If abvlist.Max < 0 And abv = abvlist.Max Then
                    firstcolumn = True
                ElseIf conside = "left" And abvlist.Min > 0 And abv = abvlist.Min Then
                    firstcolumn = True
                End If
            End If
        End If

        Return firstcolumn
    End Function

    Shared Function CheckNH3Column(abvlist As List(Of Double), abv As Double) As Integer
        Dim absabvlist As New List(Of Double)
        Dim tempfigure As Integer

        If abv <> 0 Then
            For Each value As Double In abvlist
                absabvlist.Add(Math.Abs(value))
            Next

            If abvlist.Count = 3 Then
                If Math.Abs(abv) = absabvlist.Max Then
                    tempfigure = 5
                Else
                    tempfigure = 4
                End If
            Else
                tempfigure = 4
            End If
        Else
            tempfigure = 8
        End If

        Return tempfigure
    End Function

    Shared Function CheckTemplate(stutzenID As String) As Integer
        Dim figure As Integer
        Dim sID As String = stutzenID

        If stutzenID.Length > 10 Then
            sID = stutzenID.Substring(0, 10)
        End If

        Select Case sID
            Case Library.TemplateParts.STUTZEN4
                figure = 4
            Case Library.TemplateParts.STUTZEN5
                figure = 5
            Case Library.TemplateParts.STUTZEN8
                figure = 8
            Case Else
                figure = 0
        End Select

        Return figure
    End Function

    Shared Function AllowedAngles(material As String, figure As Integer) As Double()
        Dim angles() As Double

        If figure = 5 Then
            If material = "C" Then
                angles = {16, 32, 45, 48, 60, 90}
            Else
                angles = {16, 32, 45, 48, 60, 90}
            End If
        Else
            angles = {15, 25, 60, 90}
        End If

        Return angles
    End Function

    Shared Function GetCurrentFigure(abv As Double, rowcount As Integer) As Integer
        Dim figure As Integer = 8

        If abv <> 0 Then
            figure = 5
            'check if GCHV NH3
            If General.currentunit.ApplicationType = "Condenser" And General.currentunit.ModelRangeSuffix.Substring(0, 1) = "A" Then
                figure = 4
                'if it's 3rd row entry then figure 5
                If rowcount = 3 Then
                    figure = 5
                End If
            ElseIf General.currentunit.ModelRangeName = "GACV" Then
                If rowcount = 1 Then
                    figure = 4
                Else
                    figure = 45
                End If
            End If
        End If

        Return figure
    End Function

    Shared Function GetSingleStutzen(circuit As CircuitData, ByRef consys As ConSysData, headertype As String, coil As CoilData, a As Double, xcoord As Double) As String
        Dim figure As Integer
        Dim specification, partID As String
        Dim sdiameter, l1calc, m, n, l1, l2, angle, l3, l1angle, holderoffset, svoffset, column As Double
        Dim sdata(), idlist, l1list, l2list, anglelist, l3list As List(Of String)
        Dim corangle, corl1, corl2 As Boolean

        Try
            specification = GetSpecification(79, circuit.CoreTube.Material, "Stutzen", circuit.FinType)
            sdiameter = GetTubeDiameter(circuit.FinType)
            partID = ""

            'only DX outlet and FP-N outlet is type 5
            If consys.VType = "P" Or headertype = "inlet" Then
                figure = 3
            Else
                figure = 5
            End If
            If headertype = "outlet" And (circuit.FinType = "N" Or circuit.FinType = "M") And consys.VType = "P" Then
                figure = 5
            End If
            If consys.VType = "P" And specification = "SP01-A" Then
                If circuit.CoreTube.Diameter = 12 Then 'circuit.FinType = "F" Or circuit.FinType = "M" 
                    sdiameter = 15
                Else
                    'can only be N
                    sdiameter = 18
                End If
            End If
            sdata = Database.GetStutzenData(figure, sdiameter, specification, General.TextToDouble(circuit.CoreTube.WallThickness), circuit.Pressure)
            idlist = sdata(0)
            l1list = sdata(1)
            If figure <> 8 Then
                l2list = sdata(2)
                anglelist = sdata(3)
                If figure = 3 Then
                    l3list = sdata(4)
                End If
            End If

            If circuit.FinType = "N" Or circuit.FinType = "M" Then
                holderoffset = 70
                '70 because the SV will be put much deeper into the tube
                svoffset = 60
                If consys.VType = "X" Then
                    svoffset += 10
                End If
            Else
                holderoffset = 100
                svoffset = 60
            End If
            If consys.VType = "X" And headertype = "inlet" Then
                'l1calc is the theoretical length needed to reach to specific horizontal position for the inlet
                l1calc = Math.Ceiling((coil.FinnedDepth + holderoffset - circuit.PitchX / 2) / Math.Cos(9 * Math.PI / 180))
            Else
                'for VA an adapter is needed to create a connection Ø17.2mm → if E fin, 2 adapters needed
                If specification = "SP01-A" Then
                    m = 0
                    If consys.VType = "P" And circuit.FinType = "F" Then
                        'SV not at the end of the tube
                        m = -30
                    End If
                Else
                    'for XP m=f(coil)
                    If consys.VType = "P" Then
                        If circuit.FinType = "N" Or circuit.FinType = "M" Then
                            m = 62
                        Else
                            m = GetDimM(consys.VType, coil.FinnedHeight) - 6
                        End If
                    Else
                        'tube is longer that to the coretube
                        If consys.VType = "P" Then
                            m = -27.5
                        Else
                            If circuit.FinType = "N" Or circuit.FinType = "M" Then
                                m = 60
                            Else
                                m = 48
                            End If
                        End If
                    End If
                End If
                If circuit.FinType = "E" And specification.Contains("SP03") Then
                    n = 31
                    'initally it is 37, but has depth of 6mm
                Else
                    n = 0
                End If

                If headertype = "inlet" Then
                    l1calc = Math.Ceiling(General.currentunit.TubeSheet.Dim_d + circuit.PitchX / 2 + 60 - m)
                Else
                    'checking where the outlet is, either 1 or 3
                    If circuit.ConnectionSide = "left" Then
                        column = General.IntegerRem(coil.FinnedDepth - xcoord, circuit.PitchX / 2)
                    Else
                        column = General.IntegerRem(xcoord, circuit.PitchX / 2)
                    End If
                    l1calc = Math.Ceiling(General.currentunit.TubeSheet.Dim_d + column * circuit.PitchX / 2 + svoffset - m - n)
                End If
            End If

            For i As Integer = 0 To sdata(0).Count - 1
                corangle = False
                corl1 = False
                corl2 = False
                l1 = l1list(i)
                If figure = "8" Then
                    l2 = -(circuit.CoreTubeOverhang - 5 - a)
                    angle = 90
                Else
                    l2 = l2list(i)
                    angle = anglelist(i)
                End If

                'check l2
                If Math.Abs(l2 + circuit.CoreTubeOverhang - 5 - a) < 3 Then
                    corl2 = True
                End If

                'check angle
                If figure = "3" Then
                    If angle < 0 And circuit.ConnectionSide = "left" Then
                        corangle = True
                    ElseIf angle > 0 And circuit.ConnectionSide = "right" Then
                        corangle = True
                    End If
                Else
                    If angle = 90 Then
                        corangle = True
                    End If
                End If

                'check l1
                If consys.VType = "X" Then
                    If Math.Abs(l1calc - l1) < 8 Then
                        corl1 = True
                    End If
                ElseIf consys.VType = "P" And figure <> 5 Then
                    l3 = l3list(i)
                    If circuit.FinType = "N" Or circuit.FinType = "M" Then
                        l1angle = 90
                    Else
                        l1angle = 45
                    End If
                    If Math.Abs(l1 * Math.Cos(l1angle * Math.PI / 180) + l3 - l1calc) < 4.5 Then
                        corl1 = True
                    End If
                Else
                    If Math.Abs(l1 - l1calc) < 4.5 Then
                        corl1 = True
                    End If
                End If

                If corangle And corl1 And corl2 Then
                    'use this stutzen
                    partID = idlist(i)
                    If partID = "0000911162" Then
                        partID = "0000922337"
                    End If
                    i = sdata(0).Count - 1
                End If
            Next
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
        Return partID
    End Function

    Shared Function GetSingleStutzenGADC(circuit As CircuitData, ByRef consys As ConSysData, headertype As String) As String
        If headertype = "inlet" Then

        End If
    End Function

    Shared Function GetDimM(vtype As String, finnedheight As Double) As Double
        Dim m As Double

        If vtype = "X" Then
            If finnedheight = 400 Then
                m = General.currentunit.TubeSheet.Dim_d - 94
            Else
                If General.currentunit.TubeSheet.Dim_d Mod 50 = 0 Then
                    m = 56
                ElseIf General.currentunit.TubeSheet.Dim_d Mod 2 = 0 Then
                    m = 86
                Else
                    m = 91
                End If
            End If
        Else
            m = General.currentunit.TubeSheet.Dim_d + 6
        End If

        Return m
    End Function

    Shared Function GetIDM(m As Double, pressure As Integer) As String
        Dim madapterID As String
        Dim index As Integer
        Dim stutzendata(), idlist As List(Of String)
        Dim l1list As List(Of Double)

        'use sql search, type 8, l1, diameter 17.2
        stutzendata = Database.GetStutzenData("8", 17.2, "SP03-2", 2.3, pressure)
        idlist = stutzendata(0)
        l1list = General.ConvertList(stutzendata(1))

        index = l1list.IndexOf(m)

        If index > -1 Then
            madapterID = idlist(index).Substring(3)
        Else
            madapterID = "738276"
        End If

        Return madapterID
    End Function

    Shared Function GetCutoutOffset(headerdiameter As Double, nipplediameter As Double) As Double
        Dim offset As Double
        Dim depth As Double = headerdiameter / 2        'equals dim e on all drawings

        Select Case headerdiameter
            Case 18
                depth = 7.2
            Case 22
                depth = 7.9
            Case 28
                Select Case nipplediameter
                    Case 28
                        depth = 11.5
                    Case Else
                        depth = 6.5
                End Select
            Case 35
                Select Case nipplediameter
                    Case 22
                        depth = 7.3
                    Case 28
                        depth = 9
                    Case 35
                        depth = 14
                End Select
            Case 42
                Select Case nipplediameter
                    Case 22
                        depth = 5
                    Case 28
                        depth = 6.5
                    Case 35
                        depth = 10
                    Case 42
                        depth = 16
                End Select
            Case 54
                Select Case nipplediameter
                    Case 28
                        depth = 6.2
                    Case 35
                        depth = 8
                    Case 42
                        depth = 11
                    Case 54
                        depth = 22
                End Select
            Case 64
                Select Case nipplediameter
                    Case 42
                        depth = 9
                    Case 54
                        depth = 15
                    Case 64
                        depth = 24
                End Select
            Case 76.1
                Select Case nipplediameter
                    Case 54
                        depth = 11
                    Case 64
                        depth = 18
                    Case 76.1
                        depth = 32
                End Select
            Case 88.9
                Select Case nipplediameter
                    Case 64
                        depth = 14
                    Case 76.1
                        depth = 21
                    Case 88.9
                        depth = 38
                End Select
            Case 104
                Select Case nipplediameter
                    Case 88.9
                        depth = 25
                    Case Else
                        depth = 40
                End Select
        End Select

        offset = headerdiameter / 2 - depth

        Return offset
    End Function

    Shared Function CheckCaps(material As String, diameter As Double, wallthickness As Double, onebranch As Boolean, pressure As Integer, tubetype As String) As Boolean
        Dim capsneeded As Boolean = False

        If General.currentunit.PEDCategory > 1 Or material = "V" Or material = "W" Then
            capsneeded = True
        ElseIf material = "D" Then
            capsneeded = False
        Else
            If Not onebranch Then
                If tubetype = "header" Then
                    If pressure > 41 Or diameter > 60 Then
                        capsneeded = True
                    Else
                        If ((diameter = 42 And wallthickness = 1.8) Or (diameter = 54 And wallthickness = 2.4)) And pressure > 32 Then
                            capsneeded = True
                        ElseIf ((diameter = 42 And wallthickness = 1.6) Or (diameter = 22 And wallthickness = 1) Or (diameter = 54 And wallthickness = 2)) And pressure > 16 Then
                            capsneeded = True
                        End If
                    End If
                Else
                    If pressure > 54 Or diameter > 60 Then
                        capsneeded = True
                    ElseIf ((diameter = 42 And wallthickness = 1.6) Or (diameter = 22 And wallthickness = 1) Or (diameter = 54 And wallthickness = 2)) And pressure > 16 Then
                        capsneeded = True
                    ElseIf ((diameter = 42 And wallthickness = 1.8) Or (diameter = 35 And wallthickness = 1.5)) And pressure > 41 Then
                        capsneeded = True
                    ElseIf (diameter = 54 And wallthickness = 2 And pressure > 32) Or (diameter < 20 And pressure > 46) Then
                        capsneeded = True
                    End If
                End If
            End If
        End If

        Return capsneeded
    End Function

    Shared Function CheckDXSplit(vtype As String, nodistr As Integer) As Integer
        Dim headercount As Integer = 1

        If General.currentunit.ApplicationType = "Evaporator" And vtype = "X" And General.isProdavit Then
            If nodistr > 36 Then
                headercount = 2
            ElseIf nodistr > 1 Then
                headercount = PCFData.GetValue("ConnectionSystemOutlet,HeaderTube", "Quantity")
            End If
        End If
        Return headercount
    End Function

    Shared Function CheckCO2Split(diameter As Double, length As Double) As Boolean
        Dim splitneeded As Boolean = False

        If diameter > 60 And length > 800 Then
            splitneeded = True
        ElseIf length > 1200 Then
            splitneeded = True
        End If
        Return splitneeded
    End Function

    Shared Function GetMaxLength(diameter As Double, rdsplit As Boolean) As Double
        Dim length As Double = 1199

        Select Case diameter
            Case 21.3
                length = 4200
            Case 26.9
                length = 2700
            Case 33.7
                length = 1700
            Case 42.4
                length = 1650
            Case 48.3
                length = 1250
            Case 60.3
                length = 800
        End Select

        If rdsplit Then
            length = 2250
        ElseIf diameter > 60 Then
            length = 799
        End If
        Return length
    End Function

    Shared Function GetBrineVentSize(material As String, diameter As Double) As String
        Dim ventsize As String

        If material = "C" Then
            If diameter < 30 Then
                ventsize = "G3/8"
            Else
                ventsize = "G1/2"
            End If
        Else
            If diameter < 20 Then
                ventsize = "G1/8"
            ElseIf diameter <= 26.9 Then
                ventsize = "G3/8"
            ElseIf diameter = 33.7 Then
                ventsize = "G1/2"
            ElseIf diameter = 42.4 Then
                ventsize = "G3/4"
            Else
                ventsize = "G1"
            End If
        End If
        Return ventsize
    End Function

    Shared Function GetDryCoolerVentsize(material As String, diameter As Double) As String
        Dim ventsize As String = ""

        If material = "C" Then
            If diameter < 23 Then
                ventsize = "G1/8"
            ElseIf diameter < 35 Then
                ventsize = "G3/8"
            ElseIf diameter < 54 Then
                ventsize = "G1/2"
            ElseIf diameter < 76 Then
                ventsize = "G3/4"
            Else
                ventsize = "G1"
            End If
        Else
            If diameter < 22 Then
                ventsize = "G1/8"
            ElseIf diameter > 88 Then
                ventsize = "G2"
            ElseIf diameter > 76 Then
                ventsize = "G3/2"
            ElseIf diameter > 60 Then
                ventsize = "G5/4"
            ElseIf diameter > 48 Then
                ventsize = "G1"
            ElseIf diameter > 42 Then
                ventsize = "G3/4"
            ElseIf diameter > 33 Then
                ventsize = "G1/2"
            ElseIf diameter > 26 Then
                ventsize = "G3/8"
            ElseIf diameter > 22 Then
                ventsize = "G1/4"
            End If
        End If

        Return ventsize
    End Function

    Shared Function GetGACVVentsize(header As HeaderData, circuit As CircuitData, RR As Integer) As String
        Dim ventsize As String

        If header.Tube.IsBrine Then
            If header.Tube.Materialcodeletter = "C" Then
                If header.Tube.Diameter <= 28 Then
                    ventsize = "G3/8"
                Else
                    ventsize = "G1/2"
                End If
            Else
                If circuit.NoDistributions = 1 Then
                    ventsize = "G1/8"
                Else
                    If header.Tube.Diameter <= 26.9 Then
                        ventsize = "G3/8"
                    ElseIf header.Tube.Diameter = 33.7 Then
                        ventsize = "G1/2"
                    ElseIf header.Tube.Diameter = 42.4 Then
                        ventsize = "G3/4"
                    Else
                        ventsize = "G1"
                    End If
                End If
            End If
        Else
            If header.Tube.HeaderType = "inlet" Then
                If header.Tube.Materialcodeletter = "C" Then
                    If ((circuit.FinType = "N" Or circuit.FinType = "M") And header.Tube.Diameter <= 64) Or (circuit.FinType = "F" And header.Tube.Diameter <= 54) Then
                        ventsize = "G3/8"
                    Else
                        ventsize = "G1/2"
                    End If
                Else
                    If circuit.FinType = "F" Then
                        ventsize = "G1/8"
                    Else
                        If header.Tube.Diameter = 26.9 Then
                            ventsize = "G1/8"
                        ElseIf header.Tube.Diameter = 33.7 Then
                            ventsize = "G3/8"
                        ElseIf header.Tube.Diameter < 61 Then
                            ventsize = "G1/2"
                        Else
                            ventsize = "G1"
                        End If
                    End If
                End If
            Else
                If header.Tube.Materialcodeletter = "C" Then
                    If header.Tube.Diameter < 30 Then
                        ventsize = "G3/8"
                    ElseIf header.Tube.Diameter < 100 And header.Tube.Diameter > 70 And circuit.FinType = "F" And RR = 4 Then
                        ventsize = "G1/2"
                    ElseIf (header.Tube.Diameter = 64 Or header.Tube.Diameter = 76.1) And (circuit.FinType = "N" Or circuit.FinType = "M") Then
                        ventsize = "G3/4"
                    ElseIf header.Tube.Diameter = 104 Or (header.Tube.Diameter = 88.9 And (circuit.FinType = "N" Or circuit.FinType = "M")) Then
                        ventsize = "G1"
                    Else
                        ventsize = "G1/2"
                    End If
                Else
                    If circuit.FinType = "F" Then
                        If header.Tube.Diameter < 22 Then
                            ventsize = "G1/8"
                        ElseIf header.Tube.Diameter < 42 Then
                            ventsize = "G1/4"
                        ElseIf header.Tube.Diameter < 60 Then
                            ventsize = "G3/8"
                        ElseIf header.Tube.Diameter < 76 Then
                            ventsize = "G1/2"
                        Else
                            ventsize = "G3/4"
                        End If
                    Else
                        Select Case header.Tube.Diameter
                            Case 26.9
                                ventsize = "G1/8"
                            Case 33.7
                                ventsize = "G3/8"
                            Case Else
                                If header.Tube.Diameter > 76 Then
                                    ventsize = "G1"
                                Else
                                    ventsize = "G1/2"
                                End If
                        End Select
                    End If
                End If
            End If
        End If

        If circuit.IsOnebranchEvap Then
            If header.Tube.Materialcodeletter = "C" And header.Tube.HeaderType = "outlet" Then
                ventsize = "G3/8"
            Else
                ventsize = "G1/8"
            End If
        End If

        Return ventsize
    End Function

    Shared Function GACVInletVent(header As HeaderData, fin As String) As Boolean
        Dim onheader As Boolean = False

        If (fin = "F" And header.Tube.Diameter < 64 And header.Tube.Materialcodeletter = "C") Or (fin = "N" And header.Tube.Diameter < 70 And header.Tube.Materialcodeletter = "C") Or (fin = "N" And header.Tube.Materialcodeletter <> "C") Then
            onheader = True
        End If

        Return onheader
    End Function

    Shared Function GetVentDiameter(material As String, ventsize As String) As Double
        Dim diameter As Double

        If General.currentunit.ApplicationType = "Evaporator" Then
            Select Case material
                Case "C"
                    Select Case ventsize
                        Case "G1/8"
                            diameter = 12
                        Case "G3/8"
                            diameter = 12
                        Case "G1/2"
                            diameter = 16
                        Case "G3/4"
                            diameter = 34
                        Case "G1"
                            diameter = 28
                    End Select
                Case Else
                    Select Case ventsize
                        Case ""
                            diameter = 21.5
                        Case "G1/8"
                            diameter = 14.5
                        Case "G1/4"
                            diameter = 17.5
                        Case "G3/8"
                            diameter = 21.5
                        Case "G1/2"
                            diameter = 26.5
                        Case "G3/4"
                            diameter = 33
                        Case "G1"
                            diameter = 39.5
                        Case "G5/4"
                            diameter = 49.5
                        Case "G3/2"
                            diameter = 56.5
                        Case Else
                            diameter = 68.5
                    End Select
            End Select
        Else
            Select Case material
                Case "C"
                    Select Case ventsize
                        Case "G1/8"
                            diameter = 12
                        Case "G3/8"
                            diameter = 23
                        Case "G1/2"
                            diameter = 28
                        Case "G3/4"
                            diameter = 34
                        Case Else
                            diameter = 42
                    End Select
                Case Else
                    Select Case ventsize
                        Case ""
                            diameter = 21.5
                        Case "G1/8"
                            diameter = 14.5
                        Case "G1/4"
                            diameter = 17.5
                        Case "G3/8"
                            diameter = 21.5
                        Case "G1/2"
                            diameter = 26.5
                        Case "G3/4"
                            diameter = 33
                        Case "G1"
                            diameter = 39.5
                        Case "G5/4"
                            diameter = 49.5
                        Case "G3/2"
                            diameter = 56.5
                        Case Else
                            diameter = 68.5
                    End Select
            End Select
        End If

        Return diameter
    End Function

    Shared Function GetFlangeDiameter(diameter As Double) As Double
        Dim fdiameter As Double = 0
        Select Case diameter
            Case 21.3
                fdiameter = 95
            Case 22
                fdiameter = 105
            Case 26.9
                fdiameter = 105
            Case 28
                fdiameter = 115
            Case 33.7
                fdiameter = 115
            Case 35
                fdiameter = 140
            Case 42
                fdiameter = 150
            Case 42.4
                fdiameter = 140
            Case 48.3
                fdiameter = 170
            Case 54
                fdiameter = 165
            Case 60.3
                fdiameter = 170
            Case 64
                fdiameter = 185
            Case 76.1
                fdiameter = 185
            Case 88.9
                fdiameter = 200
            Case 104
                fdiameter = 220
            Case 114.3
                fdiameter = 235
        End Select
        Return fdiameter
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

    Shared Function WeldedData(erpcode As String) As Double()
        Dim reduction As Double = 0
        Dim offset As Double = 7
        Dim neutralerp As String = erpcode.Substring(5)

        Select Case neutralerp
            Case "15"
                reduction = 63
            Case "20"
                reduction = 70
            Case "25"
                reduction = 70
            Case "32"
                reduction = 72
            Case "40"
                reduction = 75
            Case "50"
                reduction = 85
            Case "65"
                reduction = 85
                offset = 9
            Case "65.1"
                reduction = 85
                offset = 9
            Case "80"
                reduction = 90
                offset = 11
            Case "100"
                reduction = 97
                offset = 11
            Case "020-1"
                reduction = 70
            Case "025-1"
                reduction = 70
            Case "032-1"
                reduction = 72
            Case "040-1"
                reduction = 75
            Case "050-1"
                reduction = 88
            Case "065-1"
                reduction = 92
                offset = 9
            Case "065-1.1"
                reduction = 92
                offset = 9
            Case "080-1"
                reduction = 98
                offset = 11
            Case "100-1"
                reduction = 110
                offset = 11
        End Select

        Return {reduction, offset}
    End Function

    Shared Function GetVentIDs(header As HeaderData, ventsize As String, nodistr As Integer, coilsize As String) As String()
        Dim ventids() As String

        If General.currentunit.ModelRangeName = "GACV" Then
            If header.Tube.Materialcodeletter = "C" Then
                If nodistr = 1 Then
                    If ventsize = "G1/8" Then
                        ventids = {"0002936"}
                    Else
                        ventids = {"0309896", "0319976", "0001831"}
                    End If
                Else
                    Select Case ventsize
                        Case "G3/8"
                            ventids = {"0309896", "0319976", "0001831"}
                        Case "G1/2" '3rd 0001595
                            If header.Tube.HeaderType = "outlet" And header.Tube.Diameter > 60 And coilsize = "F4" Then
                                ventids = {"0359386", "0320043", "1368686", "0898859"}
                            ElseIf header.Tube.HeaderType = "inlet" And header.Tube.Diameter > 60 Then
                                ventids = {"0359386", "0320043", "1368686", "0858142"}
                            Else
                                ventids = {"0359386", "0320043", "1368686"}
                            End If
                        Case "G1"
                            ventids = {"00359374", "0320045", "0001376"}
                    End Select
                End If
            Else
                If nodistr = 1 Then
                    ventids = {"0013787", "0043976", "0043879", "0928192"}
                Else
                    Select Case ventsize
                        Case "G1/8"
                            If header.Tube.HeaderType = "inlet" And coilsize.Substring(0, 1) <> "N" Then
                                'special additional pipe
                                ventids = {"0013787", "0043976", "0043879", "1280175"} '1280175 '0859004
                            Else
                                ventids = {"0013787", "0043976", "0043879"}
                            End If
                        Case "G1/4"
                            ventids = {"0003359", "0043995", "0043883"}
                        Case "G3/8"
                            ventids = {"0003347", "0043986", "0043884"}
                        Case "G1/2"
                            ventids = {"0003345", "0043987", "0043885"}
                        Case "G3/4"
                            ventids = {"0003369", "0043988", "0043886"}
                        Case "G1"
                            ventids = {"0003368", "0043993", "0043889"}
                        Case "G5/4"
                            ventids = {"0043864", "0043985", "0043901"}
                        Case "G3/2"
                            ventids = {"0013162", "0043996", "0043902"}
                        Case Else
                            ventids = {"0042819", "0043992", "0042813"}
                    End Select
                End If
            End If
        Else
            If header.Tube.Materialcodeletter = "C" Then
                Select Case ventsize
                    Case "G1/8"
                        ventids = {"0309896", "0319976", "0001831"}
                    Case "G3/8"
                        ventids = {"0309896", "0319976", "0001842"}
                    Case "G1/2"
                        ventids = {"0359386", "0320043", "0001862"}
                    Case "G3/4"
                        ventids = {"0359396", "0320044", "0001910"}
                    Case Else
                        ventids = {"0359374", "0320045", "0001873"}
                End Select
            Else
                Select Case ventsize
                    Case "G1/8"
                        ventids = {"0013787", "0043976", "0043879"}
                    Case "G1/4"
                        ventids = {"0003359", "0043995", "0043883"}
                    Case "G3/8"
                        ventids = {"0003347", "0043986", "0043884"}
                    Case "G1/2"
                        ventids = {"0003345", "0043987", "0043885"}
                    Case "G3/4"
                        ventids = {"0003369", "0043988", "0043886"}
                    Case "G1"
                        ventids = {"0003368", "0043993", "0043889"}
                    Case "G5/4"
                        ventids = {"0043864", "0043985", "0043901"}
                    Case "G3/2"
                        ventids = {"0013162", "0043996", "0043902"}
                    Case Else
                        ventids = {"0042819", "0042892", "0042813"}
                End Select
            End If
        End If

        Return ventids
    End Function

    Shared Function CondenserLength(consys As ConSysData) As Boolean
        Dim split As Boolean = False

        If consys.InletHeaders.First.Ylist.Max - consys.InletHeaders.First.Ylist.Min + consys.InletHeaders.First.Overhangbottom + consys.InletHeaders.First.Overhangtop > 2250 Then
            split = True
        End If
        Return split
    End Function

    Shared Function GetAdapterID(diameter As Double) As String
        Dim adapterid As String = ""

        Select Case diameter
            Case 22.23
                adapterid = "533584"
            Case 28.57
                adapterid = "533660"
            Case 34.92
                adapterid = "524074"
            Case 41.27
                adapterid = "523883"
            Case 53.97
                adapterid = "856153"
        End Select

        Return adapterid
    End Function

    Shared Function GetNippleDiameterK65(diameter As Double) As Double
        Dim holediameter As Double

        'override the cutout diameter for K65 >=34.92
        Select Case diameter
            Case 28.57
                holediameter = 27
            Case 34.92
                holediameter = 27.3
            Case 41.27
                holediameter = 34.3
            Case 53.97
                holediameter = 39.3
            Case Else
                holediameter = 20.5
        End Select

        Return holediameter
    End Function

    Shared Function GetAdapterOffset(adapterID As String, type As String) As Double
        Dim offset As Double = 0

        If type = "adapter" Then
            Select Case adapterID
                Case "533584"
                    offset = 30
                Case "533660"
                    offset = 40.5
                Case "524074"
                    offset = 40
                Case "523883"
                    offset = 45
                Case "533680"
                    offset = 55
            End Select
        Else
            Select Case adapterID
                Case "533680"
                    offset = 10
                Case Else
                    offset = 8
            End Select
        End If

        Return offset
    End Function

    Shared Function GetSVID(pressure As Integer) As String
        Dim SVID As String

        If pressure < 49 Then
            SVID = "107560"
        Else
            SVID = "791913"
        End If

        Return SVID
    End Function

    Shared Function GetVentnormals(conside As String) As Boolean()
        Dim normals() As Boolean

        If conside = "right" Then
            normals = {False, False, True}
        Else
            normals = {False, True, False}
        End If
        Return normals
    End Function

    Shared Function GetVentoffset(fintype As String, conside As String, headermaterial As String, overhang As Double) As Double()
        Dim x, y, z, offset() As Double
        Dim height, depth As Double

        height = General.currentunit.Coillist.First.FinnedHeight
        depth = General.currentunit.Coillist.First.FinnedDepth

        'if finno > 1 then origin top right 
        If fintype = "M" Or fintype = "N" Then
            z = height - 25
            If conside = "left" Then
                x = 25
            Else
                x = depth - 25
            End If
        Else
            z = height - 12.5
            If conside = "right" Then
                x = depth - 37.5
            Else
                x = 37.5
            End If
        End If
        If headermaterial = "C" Then
            If conside = "right" Then
                x -= 10
            Else
                x += 10
            End If
        End If

        y = overhang + 15

        offset = {x, y, z}

        Return offset
    End Function

    Shared Function GetVentLength(ventsize As String, material As String) As Double
        Dim length As Double

        If material = "C" Then
            length = 32
        Else
            Select Case ventsize
                Case ""
                    length = 21.5
                Case "G1/8"
                    length = 18
                Case "G1/4"
                    length = 22
                Case "G3/8"
                    length = 27
                Case "G1/2"
                    length = 32
                Case "G3/4"
                    length = 33
                Case "G1"
                    length = 44
                Case "G5/4"
                    length = 55
                Case "G3/2"
                    length = 60
                Case Else
                    length = 75
            End Select
        End If
        Return length
    End Function

    Shared Function TubeClosingOffset(tubediameter As Double, hasSV As Boolean, plant As String, wallthickness As Double) As Double
        Dim offset As Double = 0

        If plant = "Sibiu" Then
            If hasSV Then
                Select Case tubediameter
                    Case 16
                        offset = 2.1
                    Case 18
                        offset = 2.7
                    Case 22
                        offset = 4
                    Case 28
                        offset = 5
                    Case 35
                        offset = 6.7
                    Case 42
                        offset = 8
                    Case 54
                        offset = 12
                End Select
            Else
                Select Case tubediameter
                    Case 16
                        offset = 4.5
                    Case 18
                        offset = 5
                    Case 22
                        offset = 6
                    Case 28
                        offset = 7.5
                    Case 35
                        offset = 9.5
                    Case 42
                        offset = 11
                    Case 54
                        offset = 15
                End Select
            End If
        Else
            If hasSV Then
                Select Case tubediameter
                    Case 16
                        offset = 0.5
                    Case 18
                        offset = 1.5
                    Case 22
                        If wallthickness = 1 Then
                            offset = 2.6
                        Else
                            offset = 2.3
                        End If
                    Case 22.23
                        offset = 5
                    Case 28
                        offset = 3.3
                    Case 28.57
                        offset = 6
                    Case 34.92
                        offset = 7
                    Case 35
                        If wallthickness = 2 Then
                            offset = 3.5
                        Else
                            offset = 5.5
                        End If
                    Case 41.27
                        offset = 8
                    Case 42
                        Select Case wallthickness
                            Case 1.6
                                offset = 8.3
                            Case 1.8
                                offset = 6
                            Case Else
                                offset = 6.5
                        End Select
                    Case 53.97
                        offset = 8
                    Case 54
                        Select Case wallthickness
                            Case 2
                                offset = 12.8
                            Case 2.4
                                offset = 8.6
                            Case Else
                                offset = 10
                        End Select
                End Select
            Else
                Select Case tubediameter
                    Case 16
                        offset = 2.5
                    Case 18
                        offset = 3
                    Case 22
                        If wallthickness = 1 Then
                            offset = 5.1
                        Else
                            offset = 4.8
                        End If
                    Case 22.23
                        offset = 5
                    Case 28
                        offset = 6
                    Case 28.57
                        offset = 6
                    Case 34.92
                        offset = 7
                    Case 35
                        offset = 7.4
                    Case 41.27
                        offset = 8
                    Case 42
                        If wallthickness = 2.6 Then
                            offset = 11
                        Else
                            offset = 10
                        End If
                    Case 53.97
                        offset = 8
                    Case 54
                        Select Case wallthickness
                            Case 2
                                offset = 15
                            Case 2.4
                                offset = 13
                            Case Else
                                offset = 8
                        End Select
                End Select
            End If
        End If

        Return offset
    End Function

    Shared Function GetTubeOffset(tubediameter As Double, material As String, plant As String) As Double
        Dim offset As Double

        If plant = "Sibiu" Then
            If material = "C" Or material = "D" Then
                Select Case tubediameter
                    Case 16
                        offset = 4
                    Case 18
                        offset = 4
                    Case 22
                        offset = 5
                    Case 28
                        offset = 7
                    Case 35
                        offset = 9
                    Case 42
                        offset = 9
                    Case 54
                        offset = 15
                    Case 64
                        offset = 20
                    Case 76.1
                        offset = 25
                    Case 88.9
                        offset = 30
                    Case 104
                        offset = 35
                    Case Else
                        offset = 0
                End Select
            Else
                Select Case tubediameter
                    Case 21.3
                        offset = 7
                    Case 26.9
                        offset = 9
                    Case 33.7
                        offset = 11
                    Case 42.4
                        offset = 13
                    Case 48.3
                        offset = 15
                    Case 60.3
                        offset = 17
                    Case Else
                        offset = 0
                End Select
            End If
        Else

        End If

        Return offset
    End Function

    Shared Function CoverSheetOrigin(conside As String) As Double()
        Dim dy, dz As Double
        Dim fandiameter As Integer = CDbl(PCFData.GetValue("AxialFan", "FanDiameter"))
        Dim tsthickness As Double = General.currentunit.TubeSheet.Thickness 'UnitProps.CBTSWT.Text)

        If fandiameter = 630 Or fandiameter = 710 Then
            dz = 1.5
        ElseIf fandiameter = 315 Then
            dz = 4
        Else
            dz = 0.5
        End If

        If fandiameter < 600 Then
            dy = 30
        Else
            dy = 50
        End If

        If conside = "left" Then
            dy = tsthickness
        End If

        Return {dy, dz}
    End Function

    Shared Function CutoutSize(diameter As Double, mrsuffix As String, hasSV As Boolean, contype As Integer) As Double()
        Dim a, b As Double

        If mrsuffix.Substring(1, 1) = "X" Or hasSV Or contype = 1 Then
            If diameter < 20 Then
                a = 50
                b = 50
            ElseIf diameter < 30 Then
                a = 50
                b = 60
            ElseIf diameter < 40 Then
                a = 50
                b = 70
            ElseIf diameter < 48 Then
                a = 50
                b = 80
            ElseIf diameter < 60 Then
                a = 72
                b = 90
            ElseIf diameter < 76 Then
                a = 72
                b = 100
            ElseIf diameter < 88 Then
                a = 100
                b = 115
            Else
                a = 100
                b = 130
            End If
        Else
            If diameter < 48 Then
                a = 50
            ElseIf diameter < 76 Then
                a = 72
            ElseIf diameter < 100 Then
                a = 100
            ElseIf diameter = 104 Then
                a = 120
            Else
                a = 130
            End If
            b = a
        End If

        Return {a, b}
    End Function

    Shared Function CheckIfTemplate(itemID As String, itemtype As String) As Boolean
        Dim istemplate As Boolean = False

        If itemtype = "Stutzen" Then
            If itemID = Library.TemplateParts.STUTZEN4 OrElse itemID = Library.TemplateParts.STUTZEN5 OrElse itemID = Library.TemplateParts.STUTZEN45IN OrElse itemID = Library.TemplateParts.STUTZEN45OUT Then
                istemplate = True
            End If
        Else

        End If
        Return istemplate
    End Function

    Shared Function GetGADCGap(RR As Integer) As Double
        Dim gap As Double

        Select Case RR
            Case 4
                gap = 700
            Case 6
                gap = 600
            Case 8
                gap = 500
        End Select
        Return gap
    End Function

End Class
