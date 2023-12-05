Imports System.Data.SqlClient
Imports Oracle.ManagedDataAccess.Client
Public Class Database

    Shared Function GetConnectionString() As String
        Return "Server=deffbfcsq13.europe.guentner-corp.com\smalldbs;Integrated Security=SSPI;database=CADKonfigurator;User Id=EUROPE\mlewin"
    End Function

    Shared Function GetTubeThickness(tubetyp As String, diameter As String, materialcode As String, pressure As String) As Double
        Dim builder As New SqlConnectionStringBuilder(GetConnectionString)
        Dim sqlConnection1 As New SqlConnection(builder.ConnectionString)
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim tablename As String = "CSG.DB_Tubes"
        Dim maxpressurelist As New List(Of String)
        Dim wtlist As New List(Of Double)
        Dim pressurelist As New List(Of Integer)
        Dim minwt, wt As Double
        Dim sptype As Integer = General.currentunit.TubeSheet.PunchingType.ToString
        Dim tsmat As String = General.currentunit.TubeSheet.MaterialCodeLetter
        Dim tswt As Double
        Dim i As Integer = 0
        Dim loopexit As Boolean = False

        Try
            tswt = General.currentunit.TubeSheet.Thickness

            If sptype = 3 And tsmat = "S" And tswt > 1.5 And tubetyp = "CoreTube" Then
                minwt = 0.6
            Else
                minwt = 0.1
            End If

            cmd.CommandText = "Select * From " + tablename + " WHERE Typ='" + tubetyp + "' AND Diameter=" + diameter.Replace(",", ".") +
                " AND Materialcode='" + materialcode + "' AND maxPressure>=" + pressure

            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader

            While reader.Read
                wtlist.Add(reader.Item("Wallthickness"))
                maxpressurelist.Add(reader.Item("maxPressure"))
            End While

            sqlConnection1.Close()

            For Each value In maxpressurelist
                pressurelist.Add(value)
            Next

            If tubetyp = "CoreTube" Or tubetyp = "Headertube" Or tubetyp = "Bow" Or tubetyp = "Stub" Then
                Do
                    If minwt <= wtlist(i) Then
                        wt = wtlist(i)
                        If pressurelist(i) = pressurelist.Min Then
                            loopexit = True
                        End If
                    End If
                    i += 1
                    If i = pressurelist.Count Then
                        loopexit = True
                    End If
                Loop Until loopexit
            Else
                wt = wtlist(0)
            End If

        Catch ex As Exception

        End Try

        If wt <= 0.1 Then
            wt = 1
        End If

        Return wt
    End Function

    Shared Function GetBowID(type As Integer, diameter As Double, wallthickness As Double, spezi As String, ABV As Double, orbitalwelding As Boolean) As List(Of String)()
        Dim tablename As String = "CSG.DB_Bows"
        Dim builder As New SqlConnectionStringBuilder(GetConnectionString)
        Dim sqlConnection1 As New SqlConnection(builder.ConnectionString)
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim bowIDlist() As List(Of String)
        Dim idlist, l1list, wtlist As New List(Of String)
        Dim abvmax, abvmin As Double
        Dim searchtype As String
        Dim nullcounter As Integer = 0

        abvmax = ABV + 1
        abvmin = ABV - 1

        Try
            If type = 9 Then
                searchtype = "Typ=9"
            Else
                searchtype = "Typ<9"
            End If
            cmd.CommandText = "Select * From " + tablename + " WHERE " + searchtype + " AND Diameter=" + diameter.ToString.Replace(",", ".") +
                " AND Specification='" + spezi + "' AND         ABV<" + abvmax.ToString.Replace(",", ".") + " AND ABV>" +
                abvmin.ToString.Replace(",", ".") + " AND Wallthickness>=" + wallthickness.ToString.Replace(",", ".")
            If orbitalwelding Then
                cmd.CommandText += " AND ERPCode like '%.1' and ERPCode like 'BV%'"
            End If

            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader

            While reader.Read
                idlist.Add(reader.Item("Article_Number"))
                l1list.Add(reader.Item("L1"))
                wtlist.Add(reader.Item("Wallthickness"))
            End While

            For Each acnumber In idlist
                If acnumber = "NULL" Then
                    nullcounter += 1
                End If
            Next

            If nullcounter = idlist.Count Then
                idlist.Clear()
                l1list.Clear()
                wtlist.Clear()
            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        Finally
            sqlConnection1.Close()
        End Try

        bowIDlist = {idlist, l1list, wtlist}

        Return bowIDlist
    End Function

    Shared Function GetHairpinID(finnedlength As Double, pitch As Double, diameter As Double, wallthickness As Double, materialcode As String) As String
        Dim tablename As String = "CSG.DB_Hairpins"
        Dim builder As New SqlConnectionStringBuilder(GetConnectionString)
        Dim sqlConnection1 As New SqlConnection(builder.ConnectionString)
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim idlist As New List(Of String)
        Dim hairpinID As String = ""

        Try

            cmd.CommandText = "Select * From " + tablename + " WHERE Diameter=" + diameter.ToString.Replace(",", ".") + " AND FinnedLength=" +
                finnedlength.ToString + " AND Pitch=" + pitch.ToString + " AND Materialcode='" + materialcode + "' AND Wallthickness=" + wallthickness.ToString.Replace(",", ".")

            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader

            While reader.Read
                idlist.Add(reader.Item("Article_Number"))
            End While

            If idlist.Count > 0 Then
                hairpinID = idlist(0)
            End If
        Catch ex As Exception
            hairpinID = "error"
            General.CreateLogEntry(ex.ToString)
        Finally
            sqlConnection1.Close()
        End Try

        Return hairpinID
    End Function

    Shared Function GetStutzenData(figure As String, diameter As String, specification As String, wallthickness As String, pressure As Integer) As List(Of String)()
        Dim builder As New SqlConnectionStringBuilder(GetConnectionString)
        Dim sqlConnection1 As New SqlConnection(builder.ConnectionString)
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim tablename As String = "CSG.DB_CPs"
        Dim l1list, l2list, anglelist, abvlist, idlist As New List(Of String)
        Dim datalists() As List(Of String)

        Try

            cmd.CommandText = "Select * From " + tablename + " WHERE Figure=" + figure + " AND Diameter=" + diameter.Replace(",", ".") +
                " AND Specification='" + specification + "' AND Wallthickness>=" + wallthickness.Replace(",", ".") + "AND maxpressure >= " + pressure.ToString

            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader

            'add the needed parameters into the lists depending of figure typ
            While reader.Read
                idlist.Add(reader.Item("Article_Number"))
                l1list.Add(reader.Item("L1"))
                If figure <> "8" Then
                    l2list.Add(reader.Item("L2"))
                    anglelist.Add(reader.Item("Angle"))
                    If figure = "4" Or figure = "45" Then
                        abvlist.Add(reader.Item("ABV"))
                    ElseIf figure = "3" Then
                        abvlist.Add(reader.Item("L3"))
                    End If
                End If
            End While

            sqlConnection1.Close()

        Catch ex As Exception

        End Try

        Select Case figure
            Case "8"
                datalists = {idlist, l1list}
            Case "5"
                datalists = {idlist, l1list, l2list, anglelist}
            Case Else
                datalists = {idlist, l1list, l2list, anglelist, abvlist}
        End Select

        Return datalists
    End Function

    Shared Function GetCSData(modelrange As String, suffix As String, aligment As String, pressure As String, material As String, fintype As String,
                              diameter As String, headertype As String, RR As String, passes As String) As Double()
        Dim builder As New SqlConnectionStringBuilder(GetConnectionString)
        Dim sqlConnection1 As New SqlConnection(builder.ConnectionString)
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim tablename As String = "CSG.DB_CSData"
        Dim alist, dishorlist, disverlist, overtoplist, overbottomlist As New List(Of String)
        Dim csdata() As Double = {0, 0, 0, 0, 0}
        Dim RRtext As String = ""
        Dim passestext As String = ""
        Dim tempfin As String = fintype

        Try

            If RR <> "NULL" Then
                RRtext = " AND NoCoreTubeRows=" + RR
                If pressure = "46" And tempfin = "E" And modelrange.Substring(2, 1) <> "D" Then
                    tempfin = "F"
                End If
                If modelrange = "GFDV" And tempfin = "E" Then
                    tempfin = "F"
                End If
            End If
            If passes <> "NULL" Then
                passestext = " AND NoPasses=" + passes
            End If
            cmd.CommandText = "Select * From " + tablename + " WHERE ModelRangeName='" + modelrange + "' AND ModelRangeSuffix='" + suffix + "'" + RRtext + passestext +
                " AND Alignment='" + aligment + "' AND Pressure=" + pressure + " AND Material='" + material + "' AND FinType='" + tempfin +
                "' AND Diameter=" + diameter.Replace(",", ".") + " AND HeaderType='" + headertype + "'"

            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader

            'add the needed parameters into the lists, should only be one match anyway
            While reader.Read
                alist.Add(reader.Item("a"))
                dishorlist.Add(reader.Item("displacehor"))
                disverlist.Add(reader.Item("displacever"))
                overtoplist.Add(reader.Item("OverhangTop"))
                overbottomlist.Add(reader.Item("OverhangBottom"))
            End While

            sqlConnection1.Close()

            If alist.Count > 0 Then
                csdata = {General.TextToDouble(alist(0)), General.TextToDouble(dishorlist(0)), General.TextToDouble(disverlist(0)), overtoplist(0), overbottomlist(0)}
            Else
                'MsgBox("No data for " + headertyp + " position of this unittype")
            End If


        Catch ex As Exception

        End Try

        Return csdata
    End Function

    Shared Function GetCapID(diameter As String, materialcode As String, pressure As String, Optional holeneeded As Boolean = False) As String
        Dim builder As New SqlConnectionStringBuilder(GetConnectionString)
        Dim sqlConnection1 As New SqlConnection(builder.ConnectionString)
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim tablename As String = "CSG.DB_Tubes"
        Dim material, capID, capstring As String
        Dim wtlist As New List(Of Double)
        Dim idlist As New List(Of String)
        Dim index As Integer = 0

        If materialcode <> "C" Then
            material = "W"
        Else
            material = materialcode
        End If

        capID = ""

        Try
            If holeneeded = False Then
                capstring = "(TYP = 'Cap' OR TYP ='Cap_F')"
            Else
                capstring = "TYP = 'Cap_L'"
            End If
            cmd.CommandText = "Select * From " + tablename + " WHERE " + capstring + " AND Diameter=" + diameter.Replace(",", ".") + " AND Materialcode='" +
                material + "' AND maxPressure>=" + pressure

            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader

            While reader.Read = True
                idlist.Add(reader.Item("Article_Number"))
                wtlist.Add(reader.Item("Wallthickness"))
            End While

            sqlConnection1.Close()

            If idlist.Count > 0 Then
                index = wtlist.IndexOf(wtlist.Min)
                capID = idlist(index)
            End If

        Catch ex As Exception

        End Try

        Return capID
    End Function

    Shared Function GetTubeERP(diameter As String, tubetype As String, pressure As String, materialcode As String) As String
        Dim builder As New SqlConnectionStringBuilder(GetConnectionString)
        Dim sqlConnection1 As New SqlConnection(builder.ConnectionString)
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim wtlist As New List(Of Double)
        Dim erplist As New List(Of String)
        Dim erpcode As String = ""

        Try
            cmd.CommandText = "Select * from csg.DB_Tubes where Typ='" + tubetype + "' and diameter=" + diameter.Replace(",", ".") + " and maxPressure >=" + pressure + " and Materialcode='" + materialcode + "'"
            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader

            While reader.Read
                wtlist.Add(reader.Item("Wallthickness"))
                erplist.Add(reader.Item("Article_Number"))
            End While

            If wtlist.Count > 0 Then
                erpcode = erplist(wtlist.IndexOf(wtlist.Min))
            End If

            sqlConnection1.Close()

        Catch ex As Exception

        End Try

        Return erpcode
    End Function

    Shared Function GetValue(tablename As String, column As String, refcolumn As String, refvalue As String, Optional datatype As String = "string") As String
        Dim builder As New SqlConnectionStringBuilder(GetConnectionString)
        Dim sqlConnection1 As New SqlConnection(builder.ConnectionString)
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim value As String = ""

        Try
            cmd.CommandText = "Select " + column + " from " + tablename + " where " + refcolumn + " ='" + refvalue + "'"
            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader

            While reader.Read
                value = reader.Item(column)
            End While

            sqlConnection1.Close()

            If datatype = "double" And value = "" Then
                value = "0"
            End If

        Catch ex As Exception

        End Try

        Return value
    End Function

    Shared Function GetFlangeID(flangetext As String, diameter As String, nipplematerial As String) As String
        Dim builder As New SqlConnectionStringBuilder(GetConnectionString)
        Dim sqlConnection1 As New SqlConnection(builder.ConnectionString)
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim tablename As String = "CSG.DB_Flanges"
        Dim flangeID As String = ""
        Dim flangetype, flangematerial, category, flangeprops() As String


        Try
            If nipplematerial = "W" Then
                nipplematerial = "V"
            End If

            flangeprops = flangetext.Split({"/"}, 0)

            flangetype = flangeprops(0).Replace(" ", "")
            flangematerial = flangeprops(1).Replace(" ", "")
            category = flangeprops(2).Replace(" ", "")


            cmd.CommandText = "Select * From " + tablename + " WHERE Flangetype='" + flangetype + "' AND Category=" + category +
                 " AND FlangeMaterial='" + flangematerial + "' AND Diameter=" + diameter.Replace(",", ".")

            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader

            While reader.Read
                flangeID = reader.Item("Article_Number")
            End While

            sqlConnection1.Close()

        Catch ex As Exception

        End Try

        Return flangeID
    End Function

    Shared Function GetFlangeIDByERP(flangeerp As String) As String
        Dim builder As New SqlConnectionStringBuilder(GetConnectionString)
        Dim sqlConnection1 As New SqlConnection(builder.ConnectionString)
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim tablename As String = "CSG.DB_Flanges"
        Dim flangeID As String = ""

        Try

            cmd.CommandText = "Select * From " + tablename + " WHERE ERPCode='" + flangeerp + "'"

            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader

            While reader.Read
                flangeID = reader.Item("Article_Number")
            End While


        Catch ex As Exception

        Finally
            sqlConnection1.Close()
        End Try

        Return flangeID
    End Function

    Shared Function CheckNewBatch(currentjob As JobInfo) As List(Of JobInfo)
        Dim builder As New SqlConnectionStringBuilder(GetConnectionString)
        Dim sqlConnection1 As New SqlConnection(builder.ConnectionString)
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim tablename As String = "csg.batch_csg"
        Dim oldjoblist, orderedlist As New List(Of JobInfo)

        Try
            cmd.CommandText = "Select * From " + tablename + " WHERE ((Status = 100 AND prio = 1) or prio = 3) and OrderNumber='" + currentjob.OrderNumber + "' AND OrderPosition='" + currentjob.OrderPosition + "'"
            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader

            While reader.Read
                oldjoblist.Add(ConvertNULL(reader, {"Prio", "Status", "Path", "Plant", "Users", "RequestTime", "FinishedTime", "ModelRange", "PDMID"}, reader("Uid")))
            End While

            sqlConnection1.Close()

            If oldjoblist.Count > 0 Then
                Dim joblist = From plist In oldjoblist Order By plist.RequestTime Descending

                orderedlist = joblist.ToList
            End If

        Catch ex As Exception
        Finally
            sqlConnection1.Close()
        End Try

        Return oldjoblist
    End Function

    Shared Function GetJobInfo(uid As String) As JobInfo
        Dim builder As New SqlConnectionStringBuilder(GetConnectionString)
        Dim sqlConnection1 As New SqlConnection(builder.ConnectionString)
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim tablename As String = "csg.batch_csg"
        Dim currentjob As JobInfo = Nothing

        Try
            cmd.CommandText = "Select * From " + tablename + " where uid =" + uid

            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader

            While reader.Read
                currentjob = New JobInfo With {.Uid = reader("Uid"), .OrderNumber = reader("OrderNumber"), .OrderPosition = reader("OrderPosition"), .Plant = reader("Plant"),
                    .Path = reader("Path"), .ProjectNumber = reader("ProjectNumber"), .Users = reader("Users")}
            End While

            sqlConnection1.Close()

        Catch ex As Exception
            sqlConnection1.Close()
        End Try

        Return currentjob
    End Function

    Shared Function GetAllJobs(prio As Integer) As Integer
        Dim builder As New SqlConnectionStringBuilder(GetConnectionString)
        Dim sqlConnection1 As New SqlConnection(builder.ConnectionString)
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim tablename As String = "csg.batch_csg"
        Dim i As Integer = 0
        Try
            cmd.CommandText = "Select * From " + tablename + " WHERE Prio =" + prio.ToString + " AND Status =0"
            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader
            While reader.Read
                i += 1
            End While
            sqlConnection1.Close()

        Catch ex As Exception

        End Try

        Return i
    End Function

    Shared Function AddJob(newjob As JobInfo) As Integer
        Dim builder As New SqlConnectionStringBuilder(GetConnectionString)
        Dim sqlConnection1 As New SqlConnection(builder.ConnectionString)
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim tablename As String = "csg.batch_csg"
        Dim commandtxt, columnnames, totaltext, entrytime, xmlpath As String
        Dim uid As Integer = 0

        Try

            entrytime = ConvertDatetoStr(Date.UtcNow)
            columnnames = "OrderNumber, OrderPosition, ProjectNumber, Prio, Status, Path, Plant, Users, RequestTime"

            If newjob.Path <> "" Then
                xmlpath = (newjob.Path + "\").Substring(8)
                xmlpath = xmlpath.Substring(xmlpath.IndexOf("\"), xmlpath.LastIndexOf("\") - xmlpath.IndexOf("\"))
            Else
                xmlpath = ""
            End If
            'If newjob.Prio <> 5 Then
            'Else
            '    xmlpath = newjob.Path
            'End If

            With newjob
                totaltext = "'" + .OrderNumber + "', '" + .OrderPosition + "', '" + .ProjectNumber + "', '" + .Prio.ToString + "', '0', '" + xmlpath + "', '" + .Plant + "', '" + .Users + "', '" + entrytime + "'"
            End With

            commandtxt = "INSERT INTO " + tablename + " (" + columnnames + ") VALUES (" + totaltext + ")"

            cmd.CommandText = commandtxt
            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()
            cmd.ExecuteNonQuery()
            sqlConnection1.Close()

            'return uid based on entry time and username

            commandtxt = "Select * FROM " + tablename + " WHERE RequestTime='" + entrytime + "' and Users='" + newjob.Users + "'"
            cmd.CommandText = commandtxt
            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()
            reader = cmd.ExecuteReader

            While reader.Read
                uid = reader.Item("Uid")
            End While

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        Finally
            sqlConnection1.Close()
        End Try

        Return uid
    End Function

    Shared Sub UpdateJob(currentjob As JobInfo, columns() As String, ByRef values() As String)
        Dim builder As New SqlConnectionStringBuilder(GetConnectionString)
        Dim tablename As String = "csg.batch_csg"

        Try
            If currentjob.Uid > 0 Then
                For i As Integer = 0 To columns.Count - 1
                    If values(i) <> "" Then
                        Using con As New SqlConnection(builder.ConnectionString)
                            Using cmd As New SqlCommand("Update " + tablename + " set " + columns(i) + "=@statusno where uid=@uid", con)
                                cmd.Parameters.Add("@statusno", SqlDbType.VarChar).Value = values(i)
                                cmd.Parameters.Add("@uid", SqlDbType.Int).Value = currentjob.Uid
                                con.Open()
                                cmd.ExecuteNonQuery()
                            End Using
                            con.Close()
                        End Using
                    End If
                Next
            End If

        Catch ex As Exception

        End Try

    End Sub

    Shared Sub UpdateLog(uid As String, modelrange As String, pdmid As String)
        Dim builder As New SqlConnectionStringBuilder(GetConnectionString)
        Dim sqlConnection1 As New SqlConnection(builder.ConnectionString)
        Dim cmd, command As New SqlCommand
        Dim ds As New DataSet()
        Dim tablename As String = "CSG.DB_Log"
        Dim cmdtext As String = "Select * From " + tablename + " WHERE uid = " + uid.ToString

        Try

            cmd.CommandText = cmdtext
            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            If modelrange = "" Then
                modelrange = "NULL"
            End If

            Dim da As New SqlDataAdapter(cmd)

            sqlConnection1.Open()

            command = New SqlCommand("UPDATE " + tablename + " SET Saved = @Saved , PDMID = @PDMID , ModelRange = @ModelRange WHERE uid = @olduid", sqlConnection1)
            command.Parameters.Add("@Saved", SqlDbType.VarChar).Value = "true"
            command.Parameters.Add("@PDMID", SqlDbType.VarChar).Value = pdmid
            command.Parameters.Add("@ModelRange", SqlDbType.VarChar).Value = modelrange
            command.Parameters.Add("@olduid", SqlDbType.Int).Value = uid

            command.ExecuteNonQuery()

        Catch ex As Exception

        Finally
            sqlConnection1.Close()
        End Try

    End Sub

    Shared Function ConvertDatetoStr(ordertime As Date) As String
        Dim dateyear, datemonth, dateday, datehour, datemin, datesec As String
        Dim fulldate As String

        dateyear = ordertime.Year.ToString
        dateyear = dateyear.Replace(" ", "")
        datemonth = ordertime.Month.ToString
        datemonth = datemonth.Replace(" ", "")
        dateday = ordertime.Day.ToString
        dateday = dateday.Replace(" ", "")
        datehour = ordertime.Hour.ToString
        datemin = ordertime.Minute.ToString
        datesec = ordertime.Second.ToString

        fulldate = dateyear + "-" + datemonth + "-" + dateday + " " + datehour + ":" + datemin + ":" + datesec

        Return fulldate
    End Function

    Shared Function ConvertNULL(reader As SqlDataReader, columnnames() As String, uid As Integer) As JobInfo
        Dim oldjob As New JobInfo With {.Uid = uid}
        Dim type As Type = oldjob.GetType
        Dim props As Reflection.PropertyInfo() = type.GetProperties
        Dim prop As Reflection.PropertyInfo

        Try
            For i As Integer = 4 To props.Count - 1
                prop = props(i)
                If prop.Name = columnnames(i - 4) Then
                    Debug.Print(prop.Name)
                    If Not reader.IsDBNull(i) Then
                        prop.SetValue(oldjob, reader(columnnames(i - 4)))
                    Else
                        Debug.Print("NULL")
                    End If
                End If
            Next

        Catch ex As Exception

        End Try

        Return oldjob
    End Function

    Shared Function SearchJobs(names() As String, values() As String) As List(Of JobInfo)
        Dim builder As New SqlConnectionStringBuilder(GetConnectionString)
        Dim sqlConnection1 As New SqlConnection(builder.ConnectionString)
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim tablename As String = "csg.batch_csg"
        Dim joblist As New List(Of JobInfo)


        Try
            cmd.CommandText = "Select * from " + tablename + " Where "
            For i As Integer = 0 To names.Count - 1
                cmd.CommandText += names(i) + " = '" + values(i) + "' "
                If i < names.Count - 1 Then
                    cmd.CommandText += " and "
                End If
            Next

            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader

            While reader.Read
                Dim newjob As New JobInfo With {.Uid = reader("Uid"),
                    .OrderNumber = reader("OrderNumber"),
                    .OrderPosition = reader("OrderPosition"),
                    .Prio = reader("Prio"),
                    .Status = reader("Status"),
                    .ProjectNumber = reader("ProjectNumber"),
                    .Users = reader("Users"),
                    .RequestTime = reader("RequestTime")}
                joblist.Add(newjob)
            End While

            sqlConnection1.Close()
        Catch ex As Exception

        End Try

        Return joblist
    End Function

    Shared Function GetTNfromERP(erpcode As String) As String
        Dim teilenr As String = ""
        Dim conn As New OracleConnection
        Dim cmd As New OracleCommand
        Dim reader As OracleDataReader

        Try
            conn.ConnectionString = GetPDMDBConnectionString
            cmd.CommandText = "Select TEILENUMMER FROM CDBPROD.ZEICHNUNG_V WHERE GUE_BAANARTNR like '" + erpcode + "' and Z_ART  <> 'cad_drawing'"
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            conn.Open()

            reader = cmd.ExecuteReader

            While reader.Read
                teilenr = reader("teilenummer")
            End While

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        Finally
            conn.Close()
        End Try

        Return teilenr
    End Function

    Shared Function GetPDMDBConnectionString() As String
        Return "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=deffbslor30.europe.guentner-corp.com)(PORT=1521))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=pdm)));User Id=CDBREAD; Password=8MzLFGKe;"
    End Function

    Shared Function GetStatusfromDB(teilenummer As String, zart As String) As List(Of Integer)
        Dim conn As New OracleConnection
        Dim cmd As New OracleCommand
        Dim reader As OracleDataReader
        Dim slist As New List(Of Integer)

        Try
            conn.ConnectionString = GetPDMDBConnectionString()
            cmd.CommandText = "select Z_Status from CDBPROD.ZEICHNUNG_V where Teilenummer='" + teilenummer + "' and Z_ART = '" + zart + "' and Z_Status <> 180"
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            conn.Open()

            reader = cmd.ExecuteReader
            While reader.Read
                slist.Add(reader("Z_Status"))
            End While

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        Finally
            conn.Close()
        End Try
        Return slist
    End Function

    Shared Function CancelJob(tablename As String, uid As String) As Integer
        Dim builder As New SqlConnectionStringBuilder(GetConnectionString)
        Dim sqlConnection1 As New SqlConnection(builder.ConnectionString)
        Dim command As New SqlCommand
        Dim rowsaffected As Integer

        Using con As New SqlConnection(builder.ConnectionString)
            Using cmd As New SqlCommand("Update " + tablename + " set status=@statusno where uid=@path", con)
                cmd.Parameters.Add("@statusno", SqlDbType.Int).Value = -10
                cmd.Parameters.Add("@path", SqlDbType.VarChar).Value = uid
                con.Open()
                rowsaffected = cmd.ExecuteNonQuery()
            End Using
            con.Close()
        End Using
        Return rowsaffected
    End Function

End Class
