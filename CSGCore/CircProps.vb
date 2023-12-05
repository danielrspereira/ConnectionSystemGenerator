Public Class CircProps

    Shared Function CheckforMirror(conside As String, circtype As String, unitdescription As String) As Boolean
        Dim check As Boolean = False

        If unitdescription = "VShape" Then

        Else
            If (conside = "left" And Not circtype.Contains("defrost")) Or (conside = "right" And circtype.Contains("defrost")) Then
                check = True
            End If
        End If

        Return check
    End Function

    Shared Function CheckCoords(coords() As Double, pitchx As Double, pitchy As Double, xoffset As Double, yoffset As Double) As Integer
        Dim matchinggrid As Boolean
        Dim gridcount As Integer = 0
        Dim i As Integer = 1
        Dim dx, dy As Double
        Dim crdigit As Integer
        Dim strdelta As String
        Dim dlist() As Double

        dx = Math.Round(coords(0) + pitchx / 2 + xoffset, 3, MidpointRounding.AwayFromZero)
        dy = Math.Round(coords(1) + pitchy / 2 + yoffset, 3, MidpointRounding.AwayFromZero)

        dlist = {dx, dy}
        For j As Integer = 0 To dlist.Count - 1
            strdelta = CStr(dlist(j))
            If strdelta.Contains(",") Then
                strdelta = strdelta.Substring(strdelta.LastIndexOf(",") + 1)
                If strdelta.Length > 2 Then
                    crdigit = strdelta.Substring(strdelta.Length - 2, 1)
                    If crdigit Mod 2 = 0 Then
                        dlist(j) = Math.Round(dlist(j) + 0.005, 3, MidpointRounding.AwayFromZero)
                    End If
                End If
            End If
        Next

        For j As Integer = 0 To dlist.Count - 1
            dlist(j) = Math.Round(dlist(j), 2)
        Next

        For Each delta As Double In dlist
            If delta > 0 Then
                If i Mod 2 = 0 Then
                    If General.IntegerMod(delta, pitchy) = 0 Then
                        matchinggrid = True
                    Else
                        matchinggrid = False
                    End If
                Else
                    If General.IntegerMod(delta, pitchx) = 0 Then
                        matchinggrid = True
                    Else
                        matchinggrid = False
                    End If
                End If
                If matchinggrid Then
                    gridcount += 1
                End If
            End If
            i += 1
        Next

        Return gridcount
    End Function

    Shared Function CheckFin(coilposition As String, fin As String, objsheet As SolidEdgeDraft.Sheet) As String
        Dim coordlist(), xlist, ylist As List(Of Double)
        Dim dxlist, dylist As New List(Of Double)
        Dim errorx, errory As New List(Of Boolean)
        Dim finlist As New List(Of String) From {"D", "F", "N", "K"}
        Dim fakefin As String
        Dim finpitch(), pitchx, pitchy, modx, mody As Double
        Dim finmatching As Boolean
        Dim i As Integer
        Dim fincounter As Integer = 0

        'Get the coords of the ct circles from the drawing
        coordlist = SEDraft.GetCTPositions(objsheet)
        xlist = coordlist(0)
        ylist = coordlist(1)

        If xlist.Count > 0 Then
            'Create dxlist and dylist 
            i = 1
            Do
                If xlist(i) <> xlist(0) Then
                    dxlist.Add(Math.Round(Math.Abs(xlist(i) - xlist(0)), 3))
                Else
                    dxlist.Add(0)
                End If
                If ylist(i) <> ylist(0) Then
                    dylist.Add(Math.Round(Math.Abs(ylist(i) - ylist(0)), 3))
                Else
                    dylist.Add(0)
                End If
                i += 1
            Loop Until i >= xlist.Count

            Do
                'Get the properties of the control fin
                fakefin = finlist(fincounter)
                finpitch = GNData.GetFinPitch(coilposition, fakefin)
                pitchx = finpitch(0)
                pitchy = finpitch(1)
                i = 0
                Do
                    'nominal - actual comparison for the ct positions
                    If dxlist(i) <> 0 Then
                        modx = General.IntegerMod(dxlist(i), pitchx)
                        If modx <> 0 Then
                            errorx.Add(True)
                        End If
                    End If
                    If dylist(i) <> 0 Then
                        mody = General.IntegerMod(dylist(i), pitchy)
                        If mody <> 0 Then
                            errory.Add(True)
                        End If
                    End If
                    i += 1
                Loop Until i >= dxlist.Count - 1

                If errorx.Count = 0 And errory.Count = 0 Then
                    finmatching = True
                Else
                    finmatching = False
                    errorx.Clear()
                    errory.Clear()
                End If
                fincounter += 1
            Loop Until finmatching Or fincounter = finlist.Count
        Else
            fakefin = fin
        End If

        Return fakefin
    End Function

    Shared Function ChangeBowprops(ByRef bowprops() As List(Of Double), circfinpitch() As Double, finpitch() As Double, Optional type As String = "bow") As List(Of Double)()
        Dim oldx1list, oldy1list, oldx2list, oldy2list, posx1list, newlengthlist, posy1list, posx2list, posy2list As New List(Of Double)
        Dim circpitchx, circpitchy, finpitchx, finpitchy As Double
        Dim x0, y0, currentx1, currenty1, currentx2, currenty2, posx1, posy1, posx2, posy2 As Double

        'Setting up the parameters
        oldx1list = bowprops(0)
        oldy1list = bowprops(1)
        oldx2list = bowprops(2)
        oldy2list = bowprops(3)

        circpitchx = circfinpitch(0)
        circpitchy = circfinpitch(1)

        finpitchx = finpitch(0)
        finpitchy = finpitch(1)

        'Get the start point on the circ
        x0 = circpitchx / 2
        y0 = oldy1list.Max

        For i As Integer = 0 To oldx1list.Count - 1
            currentx1 = oldx1list(i)
            currenty1 = oldy1list(i)
            currentx2 = oldx2list(i)
            currenty2 = oldy2list(i)

            'Scale the values down to 1,2,3...
            posx1 = Math.Round((currentx1 - x0) / circpitchx)
            posy1 = Math.Round((y0 - currenty1) / circpitchy)
            posx2 = Math.Round((currentx2 - x0) / circpitchx)
            posy2 = Math.Round((y0 - currenty2) / circpitchy)

            posx1list.Add(posx1)
            posy1list.Add(posy1)
            posx2list.Add(posx2)
            posy2list.Add(posy2)

        Next

        'Get the start point on the fin
        x0 = finpitchx / 2
        y0 = Math.Round((oldy1list.Max + circpitchy / 2) / circpitchy * finpitchy - finpitchy / 2, 4)

        For i As Integer = 0 To oldx1list.Count - 1
            'Rescale the values to the correct position on the fin
            posx1 = posx1list(i)
            posy1 = posy1list(i)
            posx2 = posx2list(i)
            posy2 = posy2list(i)

            currentx1 = Math.Round(x0 + posx1 * finpitchx, 4)
            currenty1 = Math.Round(y0 - posy1 * finpitchy, 4)
            currentx2 = Math.Round(x0 + posx2 * finpitchx, 4)
            currenty2 = Math.Round(y0 - posy2 * finpitchy, 4)

            bowprops(0)(i) = currentx1
            bowprops(1)(i) = currenty1
            bowprops(2)(i) = currentx2
            bowprops(3)(i) = currenty2

        Next

        If type = "bow" Then
            'Recalc the length
            newlengthlist = ChangeLength(bowprops)

            bowprops(4) = newlengthlist
        End If

        Return bowprops
    End Function

    Shared Function ChangeLength(bowprops() As List(Of Double)) As List(Of Double)
        Dim x1list, y1list, x2list, y2list, lengthlist As New List(Of Double)
        Dim dx, dy, length As Double

        x1list = bowprops(0)
        y1list = bowprops(1)
        x2list = bowprops(2)
        y2list = bowprops(3)

        For i As Integer = 0 To x1list.Count - 1
            dx = Math.Abs(x1list(i) - x2list(i))
            dy = Math.Abs(y1list(i) - y2list(i))
            length = Math.Round(Math.Sqrt(dx ^ 2 + dy ^ 2), 2)
            lengthlist.Add(length)
        Next

        Return lengthlist
    End Function

    Shared Function ChangeFrame(ByRef oldframe() As Double, circfinpitch() As Double, finpitch() As Double) As Double()
        Dim xmax As Double = oldframe(0)
        Dim ymax As Double = oldframe(1)
        Dim newx, newy, newframe() As Double

        newx = Math.Round(xmax / circfinpitch(0) * finpitch(0), 3)
        newy = Math.Round(ymax / circfinpitch(1) * finpitch(1), 3)

        newframe = {newx, newy}

        Return newframe
    End Function

    Shared Function CheckBowvsCP(bowpoint() As Double, inoutlist() As List(Of Double)) As Boolean
        Dim bendneeded As Boolean = False
        Dim wrapcount As Integer

        wrapcount = GetWrapCount(bowpoint, inoutlist)

        If wrapcount > 0 Then
            bendneeded = True
        End If

        Return bendneeded

    End Function

    Shared Function GetWrapCount(bowpoint() As Double, valuelists() As List(Of Double), Optional skipvalue As Integer = -1) As Integer
        Dim x1b1, y1b1, x1b2, y1b2, x2b1, y2b1, x2b2, y2b2, minb1, maxb1, dxb1, dyb1, dxb2, dyb2, l1, l2 As Double
        Dim orientation1, orientation2 As String
        Dim doublecheck As Boolean
        Dim wrapcount As Integer = 0

        'b1 = bow1 - bow with type 0
        x1b1 = bowpoint(0)
        y1b1 = bowpoint(1)
        x2b1 = bowpoint(2)
        y2b1 = bowpoint(3)

        If skipvalue > -1 Then
            l1 = valuelists(4)(skipvalue)
        End If

        'b2 values are either (1)bowpoints from bow vs bow check or (2)cp points from bow vs cp check
        'in case of (2) first 2 values = inlet point, last 2 values = outlet point

        'check orientation of this bow - options: horizontal, vertical or diagonal
        orientation1 = GetBowOrientation(x1b1, y1b1, x2b1, y2b1)

        If orientation1 = "horizontal" Then
            minb1 = Math.Min(x1b1, x2b1)
            maxb1 = Math.Max(x1b1, x2b1)
        Else
            minb1 = Math.Min(y1b1, y2b1)
            maxb1 = Math.Max(y1b1, y2b1)
        End If

        'check all other bows / cps besides the one for comparison
        For i As Integer = 0 To valuelists(0).Count - 1
            Dim relevantbow As Boolean = False
            If i <> skipvalue Then      'dont compare a bow with itself
                'b2 = bow2 - bow from the bowlist
                x1b2 = valuelists(0)(i)
                y1b2 = valuelists(1)(i)
                x2b2 = valuelists(2)(i)
                y2b2 = valuelists(3)(i)

                If skipvalue > -1 Then
                    l2 = valuelists(4)(i)
                End If

                orientation2 = GetBowOrientation(x1b2, y1b2, x2b2, y2b2)

                'select check option
                Select Case orientation1
                    Case "horizontal"
                        'check if bow2 is between start-end of bow1
                        If y1b2 = y1b1 Then
                            If x1b2 > minb1 And x1b2 < maxb1 Then
                                relevantbow = True
                                wrapcount += 1
                            End If
                        ElseIf y2b2 = y1b1 Then
                            If x2b2 > minb1 And x2b2 < maxb1 Then
                                relevantbow = True
                                wrapcount += 1
                            End If
                        End If
                    Case "vertical"
                        'check if bow2 is between start-end of bow1
                        If x1b2 = x1b1 Then
                            If y1b2 > minb1 And y1b2 < maxb1 Then
                                relevantbow = True
                                wrapcount += 1
                            End If
                        ElseIf x2b2 = x1b1 Then
                            If y2b2 > minb1 And y2b2 < maxb1 Then
                                relevantbow = True
                                wrapcount += 1
                            End If
                        End If
                    Case Else 'diagonal
                        'reset variable
                        doublecheck = True
                        'compare the ratio of dx/dy of bow1 with ratio of dx/dy from each point
                        dxb1 = Math.Abs(x1b1 - x2b1)
                        dyb1 = Math.Abs(y1b1 - y2b1)
                        'first point of bow2
                        dxb2 = Math.Abs(x1b2 - x1b1)
                        dyb2 = Math.Abs(y1b2 - y1b1)
                        If dyb2 > 0 And y1b2 < Math.Max(y1b1, y2b1) And y1b2 > Math.Min(y1b1, y2b1) Then
                            If Math.Round(dxb1 / dyb1, 2) = Math.Round(dxb2 / dyb2, 2) And x1b2 < Math.Max(x1b1, x2b1) And x1b2 > Math.Min(x1b1, x2b1) Then
                                wrapcount += 1
                                doublecheck = False
                            End If
                        End If
                        If doublecheck Then
                            'second point of bow 2
                            dxb2 = Math.Abs(x2b2 - x1b1)
                            dyb2 = Math.Abs(y2b2 - y1b1)
                            If dyb2 > 0 And y1b2 < Math.Max(y1b1, y2b1) And y1b2 > Math.Min(y1b1, y2b1) Then
                                If Math.Round(dxb1 / dyb1, 2) = Math.Round(dxb2 / dyb2, 2) And x1b2 < Math.Max(x1b1, x2b1) And x1b2 > Math.Min(x1b1, x2b1) Then
                                    wrapcount += 1
                                    relevantbow = True
                                End If
                            End If
                        End If
                End Select
                If skipvalue > -1 Then
                    'diagonal bow longer than non diagonal bow 
                    If ((orientation1 = "diagonal" And orientation2 <> "diagonal" And l1 > l2) Or (orientation2 = "diagonal" And orientation1 <> "diagonal" And l2 > l1)) And wrapcount > 0 And relevantbow Then
                        'bend
                        wrapcount = 0
                        Exit For
                    End If
                End If
            End If
        Next

        Return wrapcount
    End Function

    Shared Function GetBowOrientation(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As String
        Dim orientation As String
        Dim minb1, maxb1 As Double

        'check orientation of this bow - options: horizontal, vertical or diagonal
        If y1 = y2 Then
            If x1 = x2 Then
                orientation = ""
            Else
                orientation = "horizontal"
                minb1 = Math.Min(x1, x2)
                maxb1 = Math.Max(x1, x2)
            End If
        ElseIf x1 = x2 Then
            If y1 = y2 Then
                orientation = ""
            Else
                orientation = "vertical"
                minb1 = Math.Min(y1, y2)
                maxb1 = Math.Max(y1, y2)
            End If
        Else
            orientation = "diagonal"
        End If
        Return orientation
    End Function

    Shared Function CheckBowinBow(bownumber As Integer, bowprops() As List(Of Double)) As Boolean
        Dim x1b1, y1b1, x2b1, y2b1 As Double
        Dim bendneeded As Boolean = False
        Dim wrapcount As Integer

        'Info: no check for length (typA/B/C) needed because this function is only used for longer lines

        'b1 = bow1 - bow with type 0
        x1b1 = bowprops(0)(bownumber)
        y1b1 = bowprops(1)(bownumber)
        x2b1 = bowprops(2)(bownumber)
        y2b1 = bowprops(3)(bownumber)

        wrapcount = GetWrapCount({x1b1, y1b1, x2b1, y2b1}, bowprops, bownumber)

        If wrapcount = 0 Then
            bendneeded = True
        End If

        Return bendneeded
    End Function

    Shared Function GetBowLevels(ByRef bowprops() As List(Of Double), circtype As String, side As String, fintype As String) As List(Of Double)()
        Dim x1b1, y1b1, x2b1, y2b1, x1b2, y1b2, x2b2, y2b2, xmin, xmax, ymin, ymax, zIn_up, zOut_up As Double
        Dim skipthis As Boolean
        Dim selectedbow, pairname As String
        Dim pairlist As New List(Of String)
        Dim partnerlist As New List(Of Integer)
        Dim bowpairlist As New List(Of List(Of Integer))
        Dim partnerline, maxlevel, currentlevel As Integer
        Dim bowdatalist As New List(Of BowData)

        'bowprops = {x1, y1, x2, y2, length, type, default level}

        Try
            'Loop over all bows, avoid double checking pair of bows
            For i As Integer = 0 To bowprops(0).Count - 2
                x1b1 = bowprops(0)(i)
                y1b1 = bowprops(1)(i)
                x2b1 = bowprops(2)(i)
                y2b1 = bowprops(3)(i)

                xmax = Math.Max(x1b1, x2b1)
                xmin = Math.Min(x1b1, x2b1)
                ymax = Math.Max(y1b1, y2b1)
                ymin = Math.Min(y1b1, y2b1)

                For j As Integer = i + 1 To bowprops(0).Count - 1
                    skipthis = False
                    x1b2 = bowprops(0)(j)
                    y1b2 = bowprops(1)(j)
                    x2b2 = bowprops(2)(j)
                    y2b2 = bowprops(3)(j)

                    'Check if second bow is too far away from first bow
                    If x1b2 > xmax And x2b2 > xmax Then
                        skipthis = True
                    ElseIf x1b2 < xmin And x2b2 < xmin Then
                        skipthis = True
                    ElseIf y1b2 > ymax And y2b2 > ymax Then
                        skipthis = True
                    ElseIf y1b2 < ymin And y2b2 < ymin Then
                        skipthis = True
                    End If

                    'Compare if bows are range 
                    If skipthis = False Then
                        'Get start value for level of the longer bow
                        If bowprops(4)(i) > bowprops(4)(j) Then
                            zIn_up = bowprops(6)(i)
                            selectedbow = "i"
                        Else
                            zIn_up = bowprops(6)(j)
                            selectedbow = "j"
                        End If
                        'Calculate the new level
                        zOut_up = GetNewLevel(x1b1, y1b1, x2b1, y2b1, x1b2, y1b2, x2b2, y2b2, zIn_up, fintype)
                        If zOut_up > zIn_up Then
                            'Add the bow pair
                            pairname = i.ToString + "\" + j.ToString
                            pairlist.Add(pairname)
                            'Raise the bowlevel
                            Select Case selectedbow
                                Case "i"
                                    bowprops(6)(i) = zOut_up
                                Case "j"
                                    bowprops(6)(j) = zOut_up
                            End Select
                        End If
                    End If
                Next
            Next

            'Check if a line is crossing 
            For i As Integer = 0 To bowprops(0).Count - 1
                Dim newbow As New BowData With {.Ps = {bowprops(0)(i), bowprops(1)(i)}, .Pe = {bowprops(2)(i), bowprops(3)(i)}}
                bowdatalist.Add(newbow)
                bowpairlist.Add(New List(Of Integer))

                For Each linepair As String In pairlist
                    partnerline = GetPartnerLine(linepair, i)
                    If partnerline <> -1 Then
                        partnerlist.Add(partnerline)
                        bowpairlist.Last.Add(partnerline)
                    End If
                Next
                maxlevel = 0
                For Each partner As Integer In partnerlist
                    bowdatalist(i).PLHot.Add(partner)
                    If bowprops(4)(partner) < bowprops(4)(i) Then
                        currentlevel = bowprops(6)(partner)
                        maxlevel = Math.Max(maxlevel, currentlevel)
                    End If
                Next
                If partnerlist.Count > 0 Then
                    If bowprops(4)(i) <> bowprops(4)(partnerlist.First) Then    'use existing level if both have same length
                        bowprops(6)(i) = maxlevel + 1
                    End If
                End If
                partnerlist.Clear()
            Next
            If circtype = "defrost" Then
                If side = "back" Then
                    Unit.brinebackbows.AddRange(bowdatalist.ToArray)
                    Unit.bpairlistback = bowpairlist
                Else
                    Unit.brinefrontbows.AddRange(bowdatalist.ToArray)
                    Unit.bpairlistfront = bowpairlist
                End If
            Else
                If side = "front" Then
                    Unit.pairlistfront = bowpairlist
                Else
                    Unit.pairlistback = bowpairlist
                End If
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return bowprops
    End Function

    Shared Function GetPartnerLine(linepair As String, linenumber As Integer) As Integer
        Dim lineinpair As Boolean = False
        Dim position, partnerline As Integer
        Dim linestr As String = linenumber.ToString
        Dim leftpart, rightpart As String

        If linepair.IndexOf(linestr) >= 0 Then
            position = linepair.IndexOf("\")

            leftpart = linepair.Substring(0, position)
            rightpart = linepair.Substring(position + 1)

            If leftpart = linestr Or rightpart = linestr Then
                lineinpair = True
            End If
        End If

        If lineinpair Then
            If CInt(leftpart) = linestr Then
                partnerline = CInt(rightpart)
            Else
                partnerline = CInt(leftpart)
            End If
        Else
            partnerline = -1
        End If

        Return partnerline
    End Function

    Shared Function GetNewLevel(P1x As Double, P1y As Double, P2x As Double, P2y As Double, Q1x As Double, Q1y As Double, Q2x As Double, Q2y As Double, zLevel As Integer, fintype As String) As Double
        Dim xp, yp, crossingpoint(), m1, m2, distance1, distance2, diameter As Double
        Dim xList(), yList(), bow1(), bow2(), bows()() As Double
        Dim pointinside As Boolean = True

        xList = {P1x, P2x, Q1x, Q2x}
        yList = {P1y, P2y, Q1y, Q2y}
        bow1 = {P1x, P1y, P2x, P2y}
        bow2 = {Q1x, Q1y, Q2x, Q2y}
        bows = {bow1, bow2}
        'Calculate the point of intersection
        crossingpoint = ContactPoint(P1x, P1y, P2x, P2y, Q1x, Q1y, Q2x, Q2y)
        xp = crossingpoint(0)
        yp = crossingpoint(1)

        If P1y = P2y And Q1y = Q2y And P1y = Q1y Then
            'xp = (P1x + P2x) / 2
            'yp = P1y
            pointinside = True
        ElseIf P1x = P2x And Q1x = Q2x And P1x = Q1x Then
            'xp = P1x
            'yp = (P1y + P2y) / 2
            pointinside = True
        Else
            'If points don't intersect in the area, check if one is wrapping the other

            For Each singlebow() As Double In bows
                If CheckPostion(singlebow, xp, yp) = False Then
                    pointinside = False
                End If
            Next
            'If not, then check for minimum distance between them, must be bigger than Ø
            If pointinside = False Then
                'check m for both lines, must be different
                m1 = CheckM(P2x - P1x, P2y - P1y)
                m2 = CheckM(Q2x - Q1x, Q2y - Q1y)
                If m1 <> m2 Then
                    diameter = GNData.GetTubeDiameter(fintype)
                    distance1 = CheckDistance(bow1, bow2, crossingpoint, diameter)
                    distance2 = CheckDistance(bow2, bow1, crossingpoint, diameter)
                    'min distance must be bigger than diameter
                    If Math.Min(distance1, distance2) < 15 Then
                        pointinside = True
                    End If
                End If
            End If
        End If

        If pointinside Then
            zLevel += 1
        End If

        Return zLevel
    End Function

    Shared Function ContactPoint(P1x As Double, P1y As Double, P2x As Double, P2y As Double, Q1x As Double, Q1y As Double, Q2x As Double, Q2y As Double) As Double()
        Dim SXY(), Sx, Sy, s As Double
        Dim DPx, DPy, DQx, DQy As Double
        Dim enumerator, denominator As Double

        DPx = P2x - P1x
        DPy = P2y - P1y
        DQx = Q2x - Q1x
        DQy = Q2y - Q1y

        enumerator = (Q1x - P1x) * DPy - (Q1y - P1y) * DPx
        denominator = -1 * DPy * DQx + DQy * DPx

        If denominator <> 0 Then
            s = enumerator / denominator
            Sx = Q1x + s * DQx
            Sy = Q1y + s * DQy
        Else
            Sx = 0
            Sy = 0
        End If

        SXY = {Sx, Sy}
        Return SXY
    End Function

    Shared Function CheckPostion(bow() As Double, xp As Double, yp As Double) As Boolean
        Dim position As Boolean = True
        Dim xmin, ymin, xmax, ymax As Double
        Dim x1, y1, x2, y2 As Double

        x1 = bow(0)
        y1 = bow(1)
        x2 = bow(2)
        y2 = bow(3)
        xmin = Math.Min(x1, x2)
        xmax = Math.Max(x1, x2)
        ymin = Math.Min(y1, y2)
        ymax = Math.Max(y1, y2)

        If xp < xmin Then
            position = False
        ElseIf xp > xmax Then
            position = False
        ElseIf yp < ymin Then
            position = False
        ElseIf yp > ymax Then
            position = False
        End If

        Return position
    End Function

    Shared Function CheckM(dx As Double, dy As Double) As Double
        Dim m As Double

        If dx = 0 Then
            m = dy
        Else
            m = Math.Round(dy / dx, 3)
        End If

        Return m
    End Function

    Shared Function CheckDistance(bow1 As Double(), bow2 As Double(), crossingpoint As Double(), diameter As Double) As Double
        Dim xs, ys, x1, y1, x2, y2, dx, dy, tempx, tempy, distance, enumerator, denominator As Double
        Dim loopexit As Boolean = False
        Dim i As Integer = 0

        Try
            ''a 
            x1 = bow2(0)
            y1 = bow2(1)

            x2 = bow2(2)
            y2 = bow2(3)

            '´b
            dx = x2 - x1
            dy = y2 - y1

            Do
                xs = bow1(i)
                ys = bow1(i + 1)
                If xs <> crossingpoint(0) Or ys <> crossingpoint(1) Then
                    'step 1 P - à
                    tempx = xs - x1
                    tempy = ys - y1

                    'Kreuzprodukt of á and ´b - abs value because 3rd dimension = 0 → vector has only 1 entry
                    enumerator = Math.Abs(tempx * dy - tempy * dx)
                    denominator = Math.Sqrt(dx ^ 2 + dy ^ 2)

                    distance = Math.Round(enumerator / denominator, 3)
                    If distance < diameter + 2 Then
                        loopexit = True
                    End If
                Else
                    'crossingpoint is always on a line, so distance would be 0
                    distance = 25
                End If

                i += 2
                If i = 4 Then
                    loopexit = True
                End If
            Loop Until loopexit

        Catch ex As Exception
            distance = 25
        End Try

        Return distance
    End Function

    Shared Function CreateBowkeys(bowprops() As List(Of Double)) As List(Of String)
        Dim lengthlist, typelist As List(Of Double)
        Dim levellist As New List(Of Double)
        Dim bowkey As String
        Dim bowkeylist As New List(Of String)

        If bowprops.Count = 8 Then
            For i As Integer = 0 To bowprops(6).Count - 1
                levellist.Add(Math.Max(bowprops(6)(i), bowprops(7)(i)))
            Next
        Else
            levellist.AddRange(bowprops(6).ToArray)
        End If

        lengthlist = bowprops(4)
        typelist = bowprops(5)

        For i As Integer = 0 To levellist.Count - 1
            bowkey = levellist(i).ToString + "\" + lengthlist(i).ToString + "\" + typelist(i).ToString
            bowkeylist.Add(bowkey)
        Next

        Return bowkeylist
    End Function

    Shared Sub ExtendBowProps(ByRef bowprops() As List(Of Double), bowids As List(Of String), l1levels As List(Of Double), circtype As String) 'As List(Of Double)()
        Dim value As Double
        Dim l1list As New List(Of Double)

        Try
            For i As Integer = 0 To bowids.Count - 1
                If bowids(i).Contains("0774198") = False And bowids(i).Contains("0808178") = False And bowids(i).Contains("0808394") = False Then
                    'get L1 from database
                    value = General.TextToDouble(Database.GetValue("CSG.DB_Bows", "L1", "Article_Number", bowids(i)))
                Else
                    'L1 will be based on level and max l1
                    value = l1levels(bowprops(6)(i) - 1)
                End If
                l1list.Add(value)
            Next

            If circtype = "defrost" Then
                bowprops = {bowprops(0), bowprops(1), bowprops(2), bowprops(3), bowprops(4), bowprops(5), bowprops(6), bowprops(7), l1list}
            Else
                bowprops = {bowprops(0), bowprops(1), bowprops(2), bowprops(3), bowprops(4), bowprops(5), bowprops(6), l1list}
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        'Return bowprops
    End Sub

    Shared Function RecalcL1(pairlist As List(Of List(Of Integer)), ByRef bowprops() As List(Of Double), bowids As List(Of String), levelcount As Integer, ctdiameter As Double,
                             pressure As Integer, material As String, fintype As String, orbitalwelding As Boolean) As List(Of Double)()
        Dim value, wallthickness, newl1 As Double
        Dim spezi, materialcode As String
        Dim nbowids() As List(Of String)

        Try
            For j As Integer = 1 To levelcount
                For i As Integer = 0 To bowprops(0).Count - 1
                    If bowprops(6)(i) = j Then
                        'only care for new items
                        If (bowids(i).Contains("0774198") Or bowids(i).Contains("0808178") Or bowids(i).Contains("0808394")) And pairlist(i).Count > 0 Then
                            'check L1 of all pair bows and take the biggest value
                            Dim templ1list As New List(Of Double)
                            For Each pair In pairlist(i)
                                templ1list.Add(bowprops(7)(pair))
                            Next
                            value = templ1list.Max
                            If bowprops(7)(i) > value + ctdiameter + 15 Then
                                newl1 = value + 10 + ctdiameter
                                bowprops(7)(i) = newl1

                                'search again for a bow in DB
                                spezi = GNData.GetSpecification(pressure, material.Substring(0, 2), "bow", fintype)
                                materialcode = GNData.GetMaterialcode(material, "bow")
                                wallthickness = Database.GetTubeThickness("Bow", ctdiameter.ToString, materialcode, pressure)
                                nbowids = Database.GetBowID(2, ctdiameter, wallthickness, spezi, bowprops(4)(i), orbitalwelding)

                                If nbowids(0).Count > 0 Then
                                    Dim bowlist As New List(Of BowData)
                                    For k As Integer = 0 To nbowids(0).Count - 1
                                        Dim newbow As New BowData With {.ID = nbowids(0)(k), .L1 = nbowids(1)(k), .Length = bowprops(4)(i), .Wallthickness = nbowids(2)(k), .Level = j}
                                        bowlist.Add(newbow)
                                    Next

                                    Dim sortedbows = From plist In bowlist Where plist.L1 >= newl1 And plist.L1 <= newl1 + 5 Order By plist.L1

                                    If sortedbows.ToList.Count > 0 Then
                                        bowids(i) = sortedbows.ToList.First.ID
                                        bowprops(7)(i) = sortedbows.ToList.First.L1
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return bowprops
    End Function

    Shared Function RecalcL1B(pairlist As List(Of List(Of Integer)), ByRef brineprops() As List(Of Double), bowprops() As List(Of Double), brineids As List(Of String), ctdiameter As Double) As List(Of Double)()
        Dim value As Double

        Try

            For i As Integer = 0 To brineids.Count - 1
                If (brineids(i).Contains("0774198") Or brineids(i).Contains("0808178")) And pairlist(i).Count > 0 Then
                    'check L1 of all pair bows and take the biggest value
                    Dim templ1list As New List(Of Double)
                    For Each pair In pairlist(i)
                        templ1list.Add(bowprops(7)(pair))
                    Next
                    value = templ1list.Max
                    brineprops(8)(i) = value + 10 + ctdiameter + 2
                End If
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return brineprops
    End Function

    Shared Function RecalcL1B2(pairlist As List(Of List(Of Integer)), ByRef brineprops() As List(Of Double), brineids As List(Of String), levelcount As Integer) As List(Of Double)()
        Dim value As Double

        Try
            If levelcount > 1 Then
                For j As Integer = 2 To levelcount
                    For i As Integer = 0 To brineids.Count - 1
                        If brineprops(6)(i) = j Then
                            If (brineids(i).Contains("0774198") Or brineids(i).Contains("0808178")) And pairlist(i).Count > 0 Then
                                'check L1 of all pair bows and take the biggest value
                                Dim templ1list As New List(Of Double)
                                For Each pair In pairlist(i)
                                    templ1list.Add(brineprops(8)(pair))
                                Next
                                value = templ1list.Max
                                brineprops(8)(i) = Math.Max(value + 17, brineprops(8)(i))
                            End If
                        End If
                    Next
                Next
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return brineprops
    End Function

    Shared Function CheckBTLayer(noblindtubes As Integer, finneddepth As Double, finpitch() As Double, coilposition As String) As Integer
        Dim tubecount, layercount As Integer
        Dim pitch As Double

        'check only necessary for condensers → staggered alignment
        If coilposition = "horizontal" Then
            pitch = finpitch(1)
        Else
            pitch = finpitch(0)
        End If

        tubecount = Math.Round(finneddepth / pitch)

        If noblindtubes > tubecount Then
            If General.IntegerMod(noblindtubes, 2 * tubecount) = 0 Then
                '50/50 splitting is possible
                layercount = Math.Round(noblindtubes / 2 / tubecount)
            Else
                'can be 33/66 splitting, enough blindtubes for more than 1 layer
                layercount = 1
            End If
        Else
            'no splitting, only 1 layer max
            layercount = 0
        End If

        'layercount * pitch is the value of displacement for the coil
        Return layercount
    End Function

    Shared Function GetCircOffset(circuit As CircuitData, coil As CoilData) As Double()
        Dim offset() As Double

        'calculate position of first core tube of the circuit, affected by circuit number & blind tubes
        'exclude sandwich design from offset check
        If General.currentunit.ApplicationType = "Evaporator" Or circuit.CircuitType = "Sandwich" Then
            offset = {0, 0}
        Else
            'consider total number of circuits (blind tubes in the middle)
            If circuit.CircuitNumber = 1 Then
                If coil.Circuits.Count = 1 Then
                    offset = {0, 0}
                Else
                    'depending of coil alignment, always staggered tube arrangement → 2x pitch
                    If coil.Alignment = "horizontal" Then
                        offset = {2 * circuit.PitchX * coil.NoBlindTubeLayers, 0}
                    Else
                        offset = {0, 2 * circuit.PitchY * coil.NoBlindTubeLayers}
                    End If
                End If
            Else
                'position depends of total no. of circuits, noblindtubes, circnumber
                If circuit.CircuitNumber = coil.Circuits.Count Then
                    'can be 2 or 4
                    If coil.Alignment = "horizontal" Then
                        offset = {coil.FinnedHeight - circuit.CircuitSize(0), 0}
                    Else
                        offset = {0, coil.FinnedHeight - circuit.CircuitSize(1)}
                    End If
                Else
                    '2 and 3 are left for 4 circuits
                    If circuit.CircuitNumber = 2 Then
                        If coil.Alignment = "horizontal" Then
                            offset = {coil.Circuits(0).CircuitSize(0), 0}
                        Else
                            offset = {0, coil.Circuits(0).CircuitSize(1)}
                        End If
                    Else
                        Dim nobttl As Integer = coil.NoBlindTubeLayers / coil.NoRows
                        If coil.Alignment = "horizontal" Then
                            offset = {coil.Circuits(0).CircuitSize(0) + coil.Circuits(1).CircuitSize(0) + circuit.PitchX * 2 * nobttl, 0}
                        Else
                            offset = {0, coil.Circuits(0).CircuitSize(0) + coil.Circuits(1).CircuitSize(0) + circuit.PitchY * 2 * nobttl}
                        End If
                    End If
                End If
            End If
        End If

        Return offset
    End Function

    Shared Function GetBrineLevels(ByRef brineprops() As List(Of Double), ByVal circprops() As List(Of Double), side As String, fintype As String) As List(Of Double)()
        Dim x1b1, y1b1, x2b1, y2b1, x1b2, y1b2, x2b2, y2b2, xmin, xmax, ymin, ymax, zIn_up, zOut_up, xmb1, ymb1, xmb2, ymb2, repaircoords() As Double
        Dim skipthis As Boolean
        Dim pairname As String
        Dim maxlevel, currentlevel As Integer
        Dim pairlist As New List(Of String)
        Dim partnerlist As New List(Of Integer)
        Dim newlevel As New List(Of Double)
        Dim bowpairlist As New List(Of List(Of Integer))

        Try

            For i As Integer = 0 To brineprops(0).Count - 1
                x1b1 = brineprops(0)(i)
                y1b1 = brineprops(1)(i)
                x2b1 = brineprops(2)(i)
                y2b1 = brineprops(3)(i)

                xmax = Math.Max(x1b1, x2b1)
                xmin = Math.Min(x1b1, x2b1)
                ymax = Math.Max(y1b1, y2b1)
                ymin = Math.Min(y1b1, y2b1)
                newlevel.Add(1)

                For j As Integer = 0 To circprops(0).Count - 1
                    skipthis = False
                    x1b2 = circprops(0)(j)
                    y1b2 = circprops(1)(j)
                    x2b2 = circprops(2)(j)
                    y2b2 = circprops(3)(j)

                    'Check if second bow is too far away from first bow
                    If x1b2 > xmax And x2b2 > xmax Then
                        skipthis = True
                    ElseIf x1b2 < xmin And x2b2 < xmin Then
                        skipthis = True
                    ElseIf y1b2 > ymax And y2b2 > ymax Then
                        skipthis = True
                    ElseIf y1b2 < ymin And y2b2 < ymin Then
                        skipthis = True
                    Else
                        If x1b2 <> x2b2 And y1b2 <> y2b2 Then
                            'diagonal but not cropped cooling bows check
                            xmb1 = Math.Round((x1b1 + x2b1) / 2)
                            ymb1 = Math.Round((y1b1 + y2b1) / 2)
                            xmb2 = Math.Round((x1b2 + x2b2) / 2)
                            ymb2 = Math.Round((y1b2 + y2b2) / 2)
                            If xmb1 = xmb2 And (y1b1 = ymb2 Or y2b1 = ymb2) And ymb2 <> ymb1 Then
                                skipthis = True
                                repaircoords = {x1b2, y1b2, x2b2, y2b2}
                                If side = "front" Then
                                    Unit.repairfront.Add(repaircoords)
                                Else
                                    Unit.repairback.Add(repaircoords)
                                End If
                            ElseIf ymb1 = ymb2 And (x1b1 = xmb2 Or x2b1 = xmb2) And xmb1 <> xmb2 Then
                                skipthis = True
                                repaircoords = {x1b2, y1b2, x2b2, y2b2}
                                If side = "front" Then
                                    Unit.repairfront.Add(repaircoords)
                                Else
                                    Unit.repairback.Add(repaircoords)
                                End If
                            End If
                        End If
                    End If

                    If skipthis = False Then
                        zIn_up = brineprops(6)(i)

                        'Calculate the new level
                        zOut_up = GetNewLevel(x1b1, y1b1, x2b1, y2b1, x1b2, y1b2, x2b2, y2b2, zIn_up, fintype)
                        zOut_up = Math.Max(zOut_up, circprops(6)(j) + 1)
                        If zOut_up > zIn_up Then
                            'Add the bow pair
                            pairname = i.ToString + "\" + j.ToString
                            pairlist.Add(pairname)
                            'Raise the bowlevel
                            newlevel(i) = zOut_up
                        End If
                    End If
                Next
            Next

            'Check if a line is crossing multiple times
            For i As Integer = 0 To brineprops(0).Count - 1
                Dim searchline As String = i.ToString + "\"
                bowpairlist.Add(New List(Of Integer))
                For Each linepair As String In pairlist
                    If linepair.Substring(0, searchline.Length) = searchline Then
                        partnerlist.Add(linepair.Substring(searchline.Length))
                        bowpairlist.Last.Add(linepair.Substring(searchline.Length))
                    End If
                Next
                maxlevel = 0
                For Each partner As Integer In partnerlist
                    If side = "back" Then
                        Unit.brinebackbows(i).PLCool.Add(partner)
                    Else
                        Unit.brinefrontbows(i).PLCool.Add(partner)
                    End If
                    currentlevel = circprops(6)(partner) + 1
                    maxlevel = Math.Max(maxlevel, currentlevel)
                Next
                If partnerlist.Count > 0 Then
                    newlevel(i) = maxlevel
                End If
                partnerlist.Clear()
            Next

            If side = "front" Then
                Unit.cpairlistfront = bowpairlist
            Else
                Unit.cpairlistback = bowpairlist
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        brineprops = {brineprops(0), brineprops(1), brineprops(2), brineprops(3), brineprops(4), brineprops(5), brineprops(6), newlevel}
        'possible to save line pairs of heating + cooling

        Return brineprops
    End Function

    Shared Function MoveCoordsbyRotation(finnedheight As Double, coords As List(Of Double), offset As Double) As List(Of Double)
        Dim newposlist As New List(Of Double)

        For i As Integer = 0 To coords.Count - 1
            Dim newpos As Double = finnedheight - coords(i) - offset
            newposlist.Add(newpos)
        Next
        Return newposlist
    End Function

    Shared Function CreateStutzenkey(figure As Integer, header As HeaderData, ctoverhang As Double, abv As Double, minangle As Double, brine As Boolean) As String
        Dim l1, l2, angle, angles(), parameters() As Double
        Dim key As String

        Select Case figure
            Case 4
                l1 = Math.Round(header.Dim_a - ctoverhang + 8 + header.Tube.WallThickness)
                l2 = abv
                angle = 180
            Case 5
                'calc l1 / l2 / angle
                angles = GNData.AllowedAngles("C", figure)
                Dim tempangles As New List(Of Double)
                For Each a In angles
                    If a >= minangle Then
                        tempangles.Add(a)
                    End If
                Next
                If brine Then
                    tempangles = New List(Of Double) From {90, 45}
                ElseIf Math.Abs(abv) > 50 Then
                    If ctoverhang >= 55 Then
                        tempangles = New List(Of Double) From {90, 45}
                    Else
                        tempangles = New List(Of Double) From {45, 60, 75, 90}
                    End If
                End If
                parameters = Calculation.Fig5Parameters(header.Dim_a, abv, header.Tube.Diameter, header.Tube.WallThickness, tempangles.ToArray, ctoverhang)
                l1 = parameters(0)
                l2 = parameters(1)
                angle = parameters(2)
            Case 8
                l1 = Math.Round(header.Dim_a - ctoverhang + header.Tube.WallThickness + 8, 1)
                l2 = 0
                angle = 0
            Case Else
                'combination figure 4 and 5
                If header.Tube.HeaderType = "outlet" Or abv = 25 Then
                    angles = {90}
                Else
                    angles = {45}
                End If
                parameters = Calculation.Fig5Parameters(header.Dim_a, abv, header.Tube.Diameter, header.Tube.WallThickness, angles, ctoverhang)
                l1 = parameters(0)
                l2 = parameters(1)
                angle = parameters(2)
        End Select

        key = l1.ToString + "\" + l2.ToString + "\" + angle.ToString

        Return key
    End Function

    Shared Sub ReplaceBows(ByRef bowids As List(Of String), bowprops As List(Of Double)(), ByRef hairpins As List(Of HairpinData))

        Try
            For i As Integer = 0 To bowids.Count - 1
                For Each hp In hairpins
                    If hp.RefBow = "" Then
                        If Math.Abs(bowprops(4)(i) - hp.Pitch) < 1 AndAlso bowprops(5)(i) = 1 AndAlso bowprops(6)(i) = 1 Then
                            hp.RefBow = bowids(i)
                        End If
                    End If
                Next
            Next

            For i As Integer = 0 To hairpins.Count - 1
                If hairpins(i).PDMID <> "NULL" And hairpins(i).PDMID <> "" Then
                    'check if 3D for hairpin exists
                    If General.GetFullFilename(General.currentjob.Workspace, hairpins(i).PDMID, ".par") <> "" Then
                        For j As Integer = 0 To bowids.Count - 1
                            If bowids(j) = hairpins(i).RefBow Then
                                bowids(j) = hairpins(i).PDMID
                            End If
                        Next
                    End If
                End If
            Next
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub ExtendElementaryCirc(ByRef bowids As List(Of String), ByRef bowprops() As List(Of Double), quantity As Integer, circsize() As Double, alignment As String)
        Dim extbowids As New List(Of String)
        Dim x1l, x2l, y1l, y2l, l1list, levellist, typelist As New List(Of Double)

        Try
            x1l.AddRange(bowprops(0).ToArray)
            y1l.AddRange(bowprops(1).ToArray)
            x2l.AddRange(bowprops(2).ToArray)
            y2l.AddRange(bowprops(3).ToArray)
            typelist.AddRange(bowprops(5).ToArray)
            levellist.AddRange(bowprops(6).ToArray)
            l1list.AddRange(bowprops(7).ToArray)

            For j As Integer = 0 To bowids.Count - 1
                For i As Integer = 1 To quantity - 1
                    extbowids.Add(bowids(j))
                    typelist.Add(bowprops(5)(j))
                    levellist.Add(bowprops(6)(j))
                    l1list.Add(bowprops(7)(j))

                    If alignment = "horizontal" Then
                        'add up x values
                        x1l.Add(Math.Round(bowprops(0)(j) + circsize(0) * i / quantity, 3))
                        y1l.Add(bowprops(1)(j))
                        x2l.Add(Math.Round(bowprops(2)(j) + circsize(0) * i / quantity, 3))
                        y2l.Add(bowprops(3)(j))
                    Else
                        'add up y values
                        x1l.Add(bowprops(0)(j))
                        x2l.Add(bowprops(2)(j))
                        If bowprops(1)(j) > circsize(1) Then
                            y1l.Add(Math.Round(bowprops(1)(j) - circsize(1) * i / quantity, 3))
                            y2l.Add(Math.Round(bowprops(3)(j) - circsize(1) * i / quantity, 3))
                        Else
                            y1l.Add(Math.Round(bowprops(1)(j) + circsize(1) * i / quantity, 3))
                            y2l.Add(Math.Round(bowprops(3)(j) + circsize(1) * i / quantity, 3))
                        End If
                    End If
                Next
            Next

            bowprops = {x1l, y1l, x2l, y2l, bowprops(4), typelist, levellist, l1list}
            bowids.AddRange(extbowids.ToArray)
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub
End Class
