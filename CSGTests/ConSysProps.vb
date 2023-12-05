Imports System.ComponentModel
Imports CSGCore

Public Class ConSysProps
    Public selectedconsys As ConSysData

    Private Sub RefreshData()
        Try
            CheckValve.Visible = False
            If UnitProps.selectedcirc.Pressure <= 16 Then
                CheckFlange.Visible = True
                CheckValve.Visible = True
            End If
            selectedconsys = UnitProps.selectedcoil.ConSyss(CInt(UnitProps.CBCircuit.Text) - 1)
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Private Sub DisableButtons()
        Dim blist As New List(Of Button) From {BBuild, BDrawing, BReset}
        For Each b In blist
            b.Enabled = False
        Next
    End Sub

    Private Sub ConSysProps_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If CBConSys.Items.Count > 0 Then
            If UnitProps.CBCircuit.Text <> "" Then
                Debug.Print(UnitProps.CBCircuit.Text)
                CBConSys.SelectedItem = UnitProps.CBCircuit.Text
                selectedconsys = UnitProps.selectedcoil.ConSyss(UnitProps.CBCircuit.Text - 1)
            Else
                CBConSys.SelectedIndex = 0
                selectedconsys = UnitProps.selectedcoil.ConSyss(0)
            End If
        Else
            selectedconsys = UnitProps.selectedcoil.ConSyss(UnitProps.CBCircuit.Text - 1)
        End If

        If General.currentunit.ApplicationType = "Evaporator" Then
            LEvap.Visible = Not General.isProdavit
            CBEvap.Visible = Not General.isProdavit
        End If
        If OrderProps.CBCSGMode.Text.Contains("Continue") Then
            BReset.Visible = True
        End If
        If General.currentjob.Prio = 5 Then
            DisableButtons()
        End If
        If General.currentunit.UnitDescription = "Dual" Then
            LInCon.Visible = True
            LOutCon.Visible = True
            LPot.Visible = True
            LaPot.Visible = True
            LOverhangBPot.Visible = True
            LOverhangTPot.Visible = True
            CBConInD.Visible = True
            CBConOutD.Visible = True
            CBPotD.Visible = True
            TBConInWT.Visible = True
            TBConOutWT.Visible = True
            TBPotWT.Visible = True
            TBaPot.Visible = True
            TBOverhangTPot.Visible = True
            TBOverhangBPot.Visible = True
        End If
        Text = "Order: " + General.currentjob.OrderNumber + " Pos.: " + General.currentjob.OrderPosition
    End Sub

    Private Sub BAddConSys_Click(sender As Object, e As EventArgs) Handles BAddConSys.Click
        Dim circcount As Integer = UnitProps.CBCircuit.Items.Count
        Dim consyscount As Integer = CBConSys.Items.Count + 1

        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        If circcount > 0 Or circcount = consyscount - 1 Then
            Dim circno As String = UnitProps.CBCircuit.Text

            Dim brine As Boolean = False

            If UnitProps.CBCType.Text.Contains("Brine") Then
                brine = True
            End If

            Dim outheader As New TubeData With {
                .Materialcodeletter = CBHeaderMat.Text,
                .Diameter = General.TextToDouble(CBHeaderOutD.Text),
                .WallThickness = General.TextToDouble(TBHeaderOutWT.Text),
                .IsBrine = brine,
                .HeaderType = "outlet",
                .Quantity = 1}

            Dim outdata As New HeaderData With {
                .Tube = outheader,
                .Dim_a = General.TextToDouble(TBaOut.Text),
                .Displacehor = General.TextToDouble(TBDishorOut.Text),
                .Displacever = General.TextToDouble(TBDisverOut.Text),
                .Overhangbottom = General.TextToDouble(TBOverhangBOut.Text),
                .Overhangtop = General.TextToDouble(TBOverhangTOut.Text),
                .Nippletubes = TBNippleOutQ.Text}

            Dim outnipple As New TubeData With {
                .Diameter = General.TextToDouble(CBNippleOutD.Text),
                .WallThickness = General.TextToDouble(TBNippleOutWT.Text),
                .Quantity = TBNippleOutQ.Text,
                .Length = General.TextToDouble(TBNippleOutL.Text),
                .HeaderType = "outlet",
                .IsBrine = brine,
                .Materialcodeletter = CBNippleMat.Text}

            Dim newconsys As New ConSysData With {
            .Circnumber = circno,
            .HasFTCon = CheckFlange.Checked,
            .HeaderAlignment = CBAlignment.Text,
            .HeaderMaterial = CBHeaderMat.Text,
            .OutletHeaders = New List(Of HeaderData) From {outdata},
            .OutletNipples = New List(Of TubeData) From {outnipple}}

            GetVType(newconsys)

            If newconsys.VType <> "X" Or UnitProps.selectedcirc.NoDistributions = 1 Then
                Dim inheader As New TubeData With {
                    .Materialcodeletter = CBHeaderMat.Text,
                    .Diameter = General.TextToDouble(CBHeaderInD.Text),
                    .WallThickness = General.TextToDouble(TBHeaderInWT.Text),
                    .IsBrine = brine,
                    .HeaderType = "inlet",
                    .Quantity = 1}

                Dim indata As New HeaderData With {
                    .Tube = inheader,
                    .Dim_a = General.TextToDouble(TBaIn.Text),
                    .Displacehor = General.TextToDouble(TBDishorIn.Text),
                    .Displacever = General.TextToDouble(TBDisverIn.Text),
                    .Overhangbottom = General.TextToDouble(TBOverhangBIn.Text),
                    .Overhangtop = General.TextToDouble(TBOverhangTIn.Text),
                    .Nippletubes = TBNippleInQ.Text}

                Dim innipple As New TubeData With {
                .Diameter = General.TextToDouble(CBNippleInD.Text),
                .WallThickness = General.TextToDouble(TBNippleInWT.Text),
                .Quantity = TBNippleInQ.Text,
                .Length = General.TextToDouble(TBNippleInL.Text),
                .HeaderType = "inlet",
                .IsBrine = brine,
                .Materialcodeletter = CBNippleMat.Text}

                newconsys.InletHeaders = New List(Of HeaderData) From {indata}
                newconsys.InletNipples = New List(Of TubeData) From {innipple}
            Else
                newconsys.InletHeaders = New List(Of HeaderData) From {New HeaderData}
                newconsys.InletNipples = New List(Of TubeData) From {New TubeData}
            End If
            UnitProps.selectedcoil.ConSyss.Add(newconsys)
            selectedconsys = newconsys
            General.consyslist.Add(selectedconsys)
            CBConSys.Items.Add(consyscount.ToString)
            CBConSys.Text = consyscount.ToString

        Else
            MsgBox("Please add a circuit first")
        End If

    End Sub

    Private Sub BEditConSys_Click(sender As Object, e As EventArgs) Handles BEditConSys.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        Try
            With selectedconsys
                .HasFTCon = CheckFlange.Checked
                .HeaderAlignment = CBAlignment.Text
                .HeaderMaterial = CBHeaderMat.Text
                If CBEvap.Text <> "" And Not General.isProdavit And General.currentunit.ApplicationType = "Evaporator" Then
                    If CBEvap.Text = "Flooded" Then
                        .VType = "P"
                    Else
                        .VType = "X"
                    End If
                End If
                If .OutletHeaders.Count = 0 Then
                    .OutletHeaders.Add(New HeaderData With {.Tube = New TubeData})
                End If
                With .OutletHeaders.First
                    .Dim_a = General.TextToDouble(TBaOut.Text)
                    .Displacehor = General.TextToDouble(TBDishorOut.Text)
                    .Displacever = General.TextToDouble(TBDisverOut.Text)
                    .Overhangtop = General.TextToDouble(TBOverhangTOut.Text)
                    .Overhangbottom = General.TextToDouble(TBOverhangBOut.Text)
                    .Nippletubes = General.TextToDouble(TBNippleOutQ.Text)
                    With .Tube
                        .Diameter = General.TextToDouble(CBHeaderOutD.Text)
                        .WallThickness = General.TextToDouble(TBHeaderOutWT.Text)
                        .TubeType = "header"
                        .HeaderType = "outlet"
                        .Materialcodeletter = CBHeaderMat.Text
                        .Quantity = 1
							If UnitProps.selectedcirc.CircuitType = "Defrost" Then
                            .IsBrine = True
                        End If  
                    End With
                End With
                If .OutletNipples.Count = 0 Then
                    .OutletNipples.Add(New TubeData)
                End If
                With .OutletNipples.First
                    .Diameter = General.TextToDouble(CBNippleOutD.Text)
                    .WallThickness = General.TextToDouble(TBNippleOutWT.Text)
                    .Quantity = General.TextToDouble(TBNippleOutQ.Text)
                    .Length = General.TextToDouble(TBNippleOutL.Text)
                    .Materialcodeletter = CBNippleMat.Text
                    If UnitProps.selectedcirc.CircuitType = "Defrost" Then
                        .IsBrine = True
                    End If
                    If CBAngleOut.Text <> "" Then
                        .Angle = CInt(CBAngleOut.Text)
                    Else
                        .Angle = 0
                    End If
                End With

                If CBHeaderInD.Text <> "" Then
                    If .InletHeaders.Count = 0 Then
                        .InletHeaders.Add(New HeaderData With {.Tube = New TubeData})
                    End If
                    With .InletHeaders.First
                        .Dim_a = General.TextToDouble(TBaIn.Text)
                        .Displacehor = General.TextToDouble(TBDishorIn.Text)
                        .Displacever = General.TextToDouble(TBDisverIn.Text)
                        .Overhangtop = General.TextToDouble(TBOverhangTIn.Text)
                        .Overhangbottom = General.TextToDouble(TBOverhangBIn.Text)
                        .Nippletubes = General.TextToDouble(TBNippleInQ.Text)
                        With .Tube
                            .Diameter = General.TextToDouble(CBHeaderInD.Text)
                            .WallThickness = General.TextToDouble(TBHeaderInWT.Text)
                            .TubeType = "header"
                            .HeaderType = "inlet"
                            .Materialcodeletter = CBHeaderMat.Text
                            If .Diameter > 17.5 Or selectedconsys.HeaderAlignment = "horizontal" Then
                                .Quantity = 1
                            End If
                            If UnitProps.selectedcirc.CircuitType = "Defrost" Then
                                .IsBrine = True
                            End If
                        End With
                    End With
                    If .InletNipples.Count = 0 Then
                        .InletNipples.Add(New TubeData)
                    End If
                    With .InletNipples.First
                        .Diameter = General.TextToDouble(CBNippleInD.Text)
                        .WallThickness = General.TextToDouble(TBNippleInWT.Text)
                        .Quantity = General.TextToDouble(TBNippleInQ.Text)
                        .Length = General.TextToDouble(TBNippleInL.Text)
                        .Materialcodeletter = CBNippleMat.Text
                        If UnitProps.selectedcirc.CircuitType = "Defrost" Then
                            .IsBrine = True
                        End If
                        If CBAngleIn.Text <> "" Then
                            .Angle = CInt(CBAngleIn.Text)
                        Else
                            .Angle = 0
                        End If
                    End With
                End If

                If General.currentunit.UnitDescription = "Dual" Then
                    If .OutletConjunctions.Count = 0 Then
                        .OutletConjunctions.Add(New TubeData)
                        If UnitProps.selectedcirc.Pressure > 16 Then
                            .OutletConjunctions.Add(New TubeData)
                        End If
                    End If
                    For Each conjtube In .OutletConjunctions
                        conjtube.Diameter = General.TextToDouble(CBConOutD.Text)
                        conjtube.WallThickness = General.TextToDouble(TBConOutWT.Text)
                    Next
                    If .InletConjunctions.Count = 0 Then
                        .InletConjunctions.Add(New TubeData)
                        If UnitProps.selectedcirc.Pressure > 16 Then
                            .InletConjunctions.Add(New TubeData)
                        End If
                    End If
                    For Each conjtube In .InletConjunctions
                        conjtube.Diameter = General.TextToDouble(CBConInD.Text)
                        conjtube.WallThickness = General.TextToDouble(TBConInWT.Text)
                    Next
                    If .OilSifons.Count = 0 Then
                        .OilSifons.Add(New HeaderData)
                    End If
                    For Each sifon In .OilSifons
                        sifon.Tube.Diameter = General.TextToDouble(CBPotD.Text)
                        sifon.Tube.WallThickness = General.TextToDouble(TBPotWT.Text)
                        sifon.Dim_a = General.TextToDouble(TBaPot.Text)
                        sifon.Overhangtop = General.TextToDouble(TBOverhangTPot.Text)
                        sifon.Overhangbottom = General.TextToDouble(TBOverhangBPot.Text)
                    Next
                End If
            End With
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Private Sub CBConSys_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBConSys.SelectedIndexChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBConSys.Text)
        Try
            If CBConSys.Text <> "" Then
                UnitProps.CBCircuit.Text = CBConSys.Text
                RefreshData()
                If UnitProps.selectedcirc.NoPasses = 1 Or UnitProps.selectedcirc.NoPasses = 3 Then
                    RBInletPos.Checked = True
                    RBOutletPos.Visible = True
                    RBInletPos.Visible = True
                Else
                    RBOutletPos.Visible = False
                    RBInletPos.Visible = False
                End If
                With selectedconsys
					If .VType = "P" Then
                        CBEvap.Text = "Flooded"
                    ElseIf .VType = "X" Then
                        CBEvap.Text = "Direct Exp"
                    End If	  
                    If .OutletHeaders.First.Tube.Quantity > 0 Then
                        CBAlignment.Text = .HeaderAlignment
                        CBHeaderMat.Text = .HeaderMaterial
                        CBHeaderOutD.Text = .OutletHeaders.First.Tube.Diameter.ToString.Replace(",", ".")
                        TBHeaderOutWT.Text = .OutletHeaders.First.Tube.WallThickness.ToString
                        TBaOut.Text = .OutletHeaders.First.Dim_a.ToString.Replace(",", ".")
                        TBDishorOut.Text = .OutletHeaders.First.Displacehor.ToString.Replace(",", ".")
                        TBDisverOut.Text = .OutletHeaders.First.Displacever.ToString.Replace(",", ".")
                        TBOverhangBOut.Text = .OutletHeaders.First.Overhangbottom.ToString.Replace(",", ".")
                        TBOverhangTOut.Text = .OutletHeaders.First.Overhangtop.ToString.Replace(",", ".")
                        If .OutletNipples.First.Quantity > 0 Then
                            CBNippleMat.Text = .OutletNipples.First.Materialcodeletter
                            CBNippleOutD.Text = .OutletNipples.First.Diameter.ToString.Replace(",", ".")
                            TBNippleOutWT.Text = .OutletNipples.First.WallThickness.ToString
                            TBNippleOutL.Text = .OutletNipples.First.Length.ToString.Replace(",", ".")
                            TBNippleOutQ.Text = .OutletNipples.First.Quantity
                            CBAngleOut.Text = .OutletNipples.First.Angle
                        End If
                        If General.currentunit.UnitDescription = "Dual" Then
                            If .OutletConjunctions.First.Quantity > 0 Then
                                CBConOutD.Text = .OutletConjunctions.First.Diameter.ToString.Replace(",", ".")
                                TBConOutWT.Text = .OutletConjunctions.First.WallThickness.ToString
                            End If
                            If .OilSifons.First.Tube.Quantity > 0 Then
                                CBPotD.Text = .OilSifons.First.Tube.Diameter.ToString.Replace(",", ".")
                                TBPotWT.Text = .OilSifons.First.Tube.WallThickness.ToString
                                TBaPot.Text = .OilSifons.First.Dim_a.ToString.Replace(",", ".")
                                TBOverhangTPot.Text = .OilSifons.First.Overhangtop.ToString.Replace(",", ".")
                                TBOverhangBPot.Text = .OilSifons.First.Overhangbottom.ToString.Replace(",", ".")
                            End If
                        End If
                    Else
                        CBHeaderMat.Text = UnitProps.selectedcirc.CoreTube.Materialcodeletter
                        CBAlignment.Text = "vertical"
                        CBNippleMat.Text = CBHeaderMat.Text
                        If General.currentunit.UnitDescription = "Dual" Then
                            If UnitProps.selectedcirc.IsOnebranchEvap Then
                                .OutletNipples.First.Quantity = 0
                                TBNippleOutWT.Text = "0"
                                TBNippleOutL.Text = "0"
                                TBNippleOutQ.Text = "0"
                                TBConOutWT.Text = "0"
                            Else
                                If .OutletNipples.First.Quantity > 0 Then
                                    CBNippleMat.Text = .OutletNipples.First.Materialcodeletter
                                    CBNippleOutD.Text = .OutletNipples.First.Diameter.ToString.Replace(",", ".")
                                    TBNippleOutWT.Text = .OutletNipples.First.WallThickness.ToString
                                    TBNippleOutL.Text = .OutletNipples.First.Length.ToString.Replace(",", ".")
                                    TBNippleOutQ.Text = .OutletNipples.First.Quantity
                                    CBAngleOut.Text = .OutletNipples.First.Angle
                                End If
                                If .OutletConjunctions.First.Diameter > 0 Then
                                    CBConOutD.Text = .OutletConjunctions.First.Diameter.ToString.Replace(",", ".")
                                    TBConOutWT.Text = .OutletConjunctions.First.WallThickness.ToString
                                End If

                            End If
                        Else
                            If UnitProps.selectedcirc.IsOnebranchEvap Or UnitProps.selectedcirc.NoDistributions = 1 Then
                                CBHeaderOutD.Text = .OutletNipples.First.Diameter.ToString.Replace(",", ".")
                                TBHeaderOutWT.Text = .OutletNipples.First.WallThickness.ToString.Replace(",", ".")
                            Else
                                CBHeaderOutD.Text = UnitProps.selectedcirc.CoreTube.Diameter.ToString.Replace(",", ".")
                                TBHeaderOutWT.Text = UnitProps.selectedcirc.CoreTube.WallThickness.ToString.Replace(",", ".")
                            End If
                            .OutletNipples.First.Quantity = 0
                            TBNippleOutWT.Text = "0"
                            TBNippleOutL.Text = "0"
                            TBNippleOutQ.Text = "0"
                        End If
                    End If
                    If .InletHeaders.First.Tube.Quantity > 0 Then
                        CBHeaderInD.Text = .InletHeaders.First.Tube.Diameter.ToString.Replace(",", ".")
                        TBHeaderInWT.Text = .InletHeaders.First.Tube.WallThickness.ToString
                        TBaIn.Text = .InletHeaders.First.Dim_a.ToString.Replace(",", ".")
                        TBDishorIn.Text = .InletHeaders.First.Displacehor.ToString.Replace(",", ".")
                        TBDisverIn.Text = .InletHeaders.First.Displacever.ToString.Replace(",", ".")
                        TBOverhangBIn.Text = .InletHeaders.First.Overhangbottom.ToString.Replace(",", ".")
                        TBOverhangTIn.Text = .InletHeaders.First.Overhangtop.ToString.Replace(",", ".")
                        If .InletNipples.First.Quantity > 0 Then
                            CBNippleInD.Text = .InletNipples.First.Diameter.ToString.Replace(",", ".")
                            TBNippleInWT.Text = .InletNipples.First.WallThickness.ToString
                            TBNippleInL.Text = .InletNipples.First.Length.ToString.Replace(",", ".")
                            TBNippleInQ.Text = .InletNipples.First.Quantity
                            CBAngleIn.Text = .InletNipples.First.Angle
                        End If
                        If General.currentunit.UnitDescription = "Dual" Then
                            If .InletConjunctions.First.Quantity > 0 Then
                                CBConInD.Text = .InletConjunctions.First.Diameter.ToString.Replace(",", ".")
                                TBConInWT.Text = .InletConjunctions.First.WallThickness.ToString
                            End If
                        End If
                    Else
                        If General.currentunit.UnitDescription = "Dual" Then
                            If UnitProps.selectedcirc.IsOnebranchEvap Then
                                .InletNipples.First.Quantity = 0
                                TBNippleInWT.Text = "0"
                                TBNippleInL.Text = "0"
                                TBNippleInQ.Text = "0"
                                .InletConjunctions.First.Quantity = 0
                                TBConInWT.Text = "0"
                            Else
                                If .InletNipples.First.Quantity > 0 Then
                                    CBNippleInD.Text = .InletNipples.First.Diameter.ToString.Replace(",", ".")
                                    TBNippleInWT.Text = .InletNipples.First.WallThickness.ToString
                                    TBNippleInL.Text = .InletNipples.First.Length.ToString.Replace(",", ".")
                                    TBNippleInQ.Text = .InletNipples.First.Quantity
                                    CBAngleIn.Text = .InletNipples.First.Angle
                                End If
                                If .InletConjunctions.First.Diameter > 0 Then
                                    CBConInD.Text = .InletConjunctions.First.Diameter.ToString.Replace(",", ".")
                                    TBConInWT.Text = .InletConjunctions.First.WallThickness.ToString
                                End If
                            End If
                        Else
                            If UnitProps.selectedcirc.IsOnebranchEvap Or UnitProps.selectedcirc.NoDistributions = 1 Then
                                CBHeaderInD.Text = .InletNipples.First.Diameter.ToString.Replace(",", ".")
                                TBHeaderInWT.Text = .InletNipples.First.WallThickness.ToString.Replace(",", ".")
                            Else
                                CBHeaderInD.Text = UnitProps.selectedcirc.CoreTube.Diameter.ToString.Replace(",", ".")
                                TBHeaderInWT.Text = UnitProps.selectedcirc.CoreTube.WallThickness.ToString.Replace(",", ".")
                            End If
                            .InletNipples.First.Quantity = 0
                            TBNippleInWT.Text = "0"
                            TBNippleInL.Text = "0"
                            TBNippleInQ.Text = "0"
                        End If
                    End If
                    CheckFlange.Checked = .HasFTCon

                    .ConSysFile.Orderno = General.currentjob.OrderNumber
                    .ConSysFile.Orderpos = General.currentjob.OrderPosition
                    .ConSysFile.Projectno = General.currentjob.ProjectNumber
                End With

                If General.isProdavit Then
                    BCSData.PerformClick()
                End If
                If OrderProps.CBCSGMode.Text.Contains("Continue") Then
                    BEditConSys.PerformClick()
                End If
            End If
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try

    End Sub

    Private Sub BCSData_Click(sender As Object, e As EventArgs) Handles BCSData.Click
        Dim dosearch As Boolean = False

        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        If Not General.isProdavit Then
            If General.Buttonlist.IndexOf(Me.Name + "_BEditConSys") = -1 Then
                MsgBox("Save input first")
            Else
                dosearch = True
            End If
        End If

        If General.isProdavit Or dosearch Then
            If General.currentunit.ModelRangeName = "NNNN" Then
                MRName.Show()
                MRName.Activate()
            ElseIf CBConSys.Text <> "" Then
                CheckOilCooler.Visible = False
                CheckOilCooler.Checked = False
                CheckNH3.Visible = False
                CheckNH3.Checked = False
                If UnitProps.selectedcirc.CircuitType.Contains("Defrost") Then
                    CSData.GetBrineCSData(selectedconsys, UnitProps.selectedcirc)
                Else
                    CSData.GetCSData(selectedconsys, UnitProps.selectedcirc, selectedconsys.OutletHeaders.First, UnitProps.selectedcoil)
                End If
                With selectedconsys.OutletHeaders.First
                    TBaOut.Text = .Dim_a.ToString.Replace(".", ",")
                    TBDishorOut.Text = .Displacehor.ToString.Replace(".", ",")
                    TBDisverOut.Text = .Displacever.ToString.Replace(".", ",")
                    TBOverhangTOut.Text = .Overhangtop.ToString.Replace(".", ",")
                    TBOverhangBOut.Text = .Overhangbottom.ToString.Replace(".", ",")
                End With
                TBNippleOutQ.Text = selectedconsys.OutletNipples.First.Quantity.ToString
                TBNippleOutL.Text = selectedconsys.OutletNipples.First.Length.ToString.Replace(".", ",")

                If selectedconsys.VType <> "X" Or UnitProps.selectedcirc.NoDistributions = 1 Then
                    If Not UnitProps.selectedcirc.CircuitType.Contains("Defrost") Then
                        If Not (selectedconsys.VType = "X" And UnitProps.selectedcirc.NoDistributions > 1) Then
                            CSData.GetCSData(selectedconsys, UnitProps.selectedcirc, selectedconsys.InletHeaders.First, UnitProps.selectedcoil)
                        End If
                    End If

                    With selectedconsys.InletHeaders.First
                        TBaIn.Text = .Dim_a.ToString.Replace(".", ",")
                        TBDishorIn.Text = .Displacehor.ToString.Replace(".", ",")
                        TBDisverIn.Text = .Displacever.ToString.Replace(".", ",")
                        TBOverhangTIn.Text = .Overhangtop.ToString.Replace(".", ",")
                        TBOverhangBIn.Text = .Overhangbottom.ToString.Replace(".", ",")
                    End With
                    If selectedconsys.InletNipples.Count > 0 Then
                        TBNippleInQ.Text = selectedconsys.InletNipples.First.Quantity.ToString
                        TBNippleInL.Text = selectedconsys.InletNipples.First.Length.ToString.Replace(".", ",")
                    End If
                End If
                If Not General.isProdavit Then
                    BClear.Visible = True
                End If
                If General.currentunit.ModelRangeSuffix.Substring(0, 1) = "O" Then
                    CheckOilCooler.Checked = True
                    CheckOilCooler.Visible = True
                ElseIf General.currentunit.ApplicationType = "Condenser" And General.currentunit.ModelRangeSuffix.Substring(1, 1) = "A" Then
                    CheckNH3.Checked = True
                    CheckNH3.Visible = True
                End If
            Else
                MsgBox("Select or create a connection system first")
            End If
        End If
    End Sub

    Private Sub CBHeaderMat_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBHeaderMat.SelectedIndexChanged
        ChangeCBValues(CBHeaderOutD, CBHeaderMat.Text)
        ChangeCBValues(CBHeaderInD, CBHeaderMat.Text)
    End Sub

    Private Sub CBNippleMat_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBNippleMat.SelectedIndexChanged
        ChangeCBValues(CBNippleOutD, CBNippleMat.Text)
        ChangeCBValues(CBNippleInD, CBNippleMat.Text)
        ChangeCBValues(CBConInD, CBHeaderMat.Text)
        ChangeCBValues(CBConOutD, CBHeaderMat.Text)
    End Sub

    Private Sub CBFlange_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBFlange.SelectedIndexChanged
        CSData.FlangeData(selectedconsys, CBFlange.Text)

        'update nipple lengths
        With selectedconsys
            If .OutletNipples.Count > 0 Then
                TBNippleOutL.Text = .OutletNipples.First.Length.ToString.Replace(",", ".")
            End If
            If .InletNipples.Count > 0 Then
                TBNippleInL.Text = .InletNipples.First.Length.ToString.Replace(",", ".")
            End If
            If .ConType = 0 And CBFlange.Text <> "" Then
                If CBFlange.Text.Contains("Loose") Then
                    .ConType = 2
                ElseIf CBFlange.Text.Contains("Thread") Then
                    .ConType = 1
                Else
                    .ConType = 3
                End If
            End If
        End With

    End Sub

    Shared Sub ChangeCBValues(objcb As ComboBox, materialcode As String)
        Dim valuelist() As Double

        objcb.Items.Clear()

        Select Case materialcode
            Case "C"
                valuelist = {9.52, 12, 15, 16, 18, 22, 28, 35, 42, 54, 64, 76.1, 88.9, 104, 133, 159, 200}
            Case "D"
                valuelist = {22.23, 28.57, 34.92, 41.27, 53.97}
            Case Else
                valuelist = {9.52, 12, 15, 17.2, 21.3, 26.9, 33.7, 42.4, 48.3, 60.3, 76.1, 88.9, 114.3, 139.7, 168.3}
        End Select

        For Each diameter As Double In valuelist
            objcb.Items.Add(diameter.ToString.Replace(",", "."))
        Next

    End Sub

    Shared Sub GetVType(ByRef consys As ConSysData)

        If General.currentunit.ApplicationType = "Condenser" Then
            consys.VType = "D"
        ElseIf consys.HeaderAlignment = "horizontal" Then
            consys.VType = "P"
        ElseIf UnitProps.selectedcirc.Pressure < 17 Then
            consys.VType = "P"
        Else
            consys.VType = General.currentunit.ModelRangeSuffix.Substring(1, 1)
        End If
        If consys.VType = "G" Then
            consys.VType = "P"
        End If
    End Sub

    Private Sub BClear_Click(sender As Object, e As EventArgs) Handles BClear.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        General.currentunit.ModelRangeName = "NNNN"
    End Sub

    Private Sub BBuild_Click(sender As Object, e As EventArgs) Handles BBuild.Click
        Dim docreate As Boolean = False
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        Try
            If Not General.isProdavit Then
                If General.Buttonlist.IndexOf(Me.Name + "_BEditConSys") = -1 Then
                    MsgBox("Save input first")
                Else
                    docreate = True
                End If
            End If
            If General.isProdavit Or docreate Then
                BEditConSys.PerformClick()

                If UnitProps.CBCoil.Text <> "" Then
                    If UnitProps.selectedcirc.CircuitFile.Fullfilename = "" Then
                        'check out circuit drawing
                        WSM.CheckoutCircs(UnitProps.selectedcirc.PDMID, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                    End If

                    If UnitProps.selectedcoil.CoilFile.Fullfilename = "" Then
                        Unit.CreateCoil(UnitProps.selectedcoil, UnitProps.selectedcoil.Circuits(0))

                        'assemble coils in same assembly
                        General.currentunit.UnitFile.Fullfilename = General.currentjob.Workspace + "\Masterassembly.asm"
                        If Not IO.File.Exists(General.currentunit.UnitFile.Fullfilename) Then
                            'create the assembly
                            Unit.CreateMasterAssembly()
                        End If

                        'add coil
                        Unit.AddAssembly(General.currentunit.UnitFile.Fullfilename, UnitProps.selectedcoil.CoilFile.Fullfilename, General.currentunit.Occlist)
                    End If

                    'check if adjust fin is needed
                    If UnitProps.selectedcirc.CircuitType.Contains("Defrost") Then
                        If Not IO.File.Exists(General.currentjob.Workspace + "\CoilSupporttube" + UnitProps.selectedcoil.Number.ToString + ".par") Then
                            'check out circuit drawing
                            WSM.CheckoutCircs(UnitProps.selectedcirc.PDMID, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                            Unit.AdjustCoilDefrost(UnitProps.selectedcoil, UnitProps.selectedcirc)
                        End If
                    End If

                    If selectedconsys.ConSysFile.Fullfilename = "" Then
                        selectedconsys.BOMItem = Order.GetConsysBOMItem(selectedconsys, UnitProps.selectedcoil, UnitProps.selectedcirc.CircuitType)
                        Unit.CreateConsys(UnitProps.selectedcoil, UnitProps.selectedcirc, selectedconsys)
                    End If

                    Unit.AddAssembly(UnitProps.selectedcoil.CoilFile.Fullfilename, selectedconsys.ConSysFile.Fullfilename, UnitProps.selectedcoil.Occlist)

                    'Update UI in case of changed values for overhang and a
                    UpdateHeaderInfo(selectedconsys.OutletHeaders.First)

                    If selectedconsys.InletHeaders.First.Tube.Quantity > 0 Then
                        UpdateHeaderInfo(selectedconsys.InletHeaders.First)
                    End If

                End If
            End If
        Catch ex As Exception
            Debug.Print(ex.ToString)
        Finally
            BReset.Visible = True
            BReset.Enabled = True
        End Try

    End Sub

    Private Sub UpdateHeaderInfo(header As HeaderData)

        With header
            If .Tube.HeaderType = "outlet" Then
                TBaOut.Text = .Dim_a.ToString.Replace(".", ",")
                TBDishorOut.Text = .Displacehor.ToString.Replace(".", ",")
                TBDisverOut.Text = .Displacever.ToString.Replace(".", ",")
                TBOverhangTOut.Text = .Overhangtop.ToString.Replace(".", ",")
                TBOverhangBOut.Text = .Overhangbottom.ToString.Replace(".", ",")
            Else
                TBaIn.Text = .Dim_a.ToString.Replace(".", ",")
                TBDishorIn.Text = .Displacehor.ToString.Replace(".", ",")
                TBDisverIn.Text = .Displacever.ToString.Replace(".", ",")
                TBOverhangTIn.Text = .Overhangtop.ToString.Replace(".", ",")
                TBOverhangBIn.Text = .Overhangbottom.ToString.Replace(".", ",")
            End If
        End With
        Update()
    End Sub

    Private Sub ButtonHeaderWT_Click(sender As Object, e As EventArgs) Handles ButtonHeaderWT.Click
        Dim wallthickness As Double
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")

        Try
            If CBHeaderMat.Text <> "" Then
                If CBHeaderOutD.Text <> "" Then
                    wallthickness = Database.GetTubeThickness("Headertube", CBHeaderOutD.Text, CBHeaderMat.Text, UnitProps.CBOP.Text)
                    TBHeaderOutWT.Text = wallthickness.ToString.Replace(".", ",")
                End If
                If CBHeaderInD.Text <> "" Then
                    wallthickness = Database.GetTubeThickness("Headertube", CBHeaderInD.Text, CBHeaderMat.Text, UnitProps.CBOP.Text)
                    TBHeaderInWT.Text = wallthickness.ToString.Replace(".", ",")
                End If
                If CBConOutD.Text <> "" Then
                    wallthickness = Database.GetTubeThickness("Headertube", CBConOutD.Text, CBHeaderMat.Text, UnitProps.CBOP.Text)
                    TBConOutWT.Text = wallthickness.ToString.Replace(".", ",")
                End If
                If CBConInD.Text <> "" Then
                    wallthickness = Database.GetTubeThickness("Headertube", CBConInD.Text, CBHeaderMat.Text, UnitProps.CBOP.Text)
                    TBConInWT.Text = wallthickness.ToString.Replace(".", ",")
                End If
                If CBPotD.Text <> "" Then
                    wallthickness = Database.GetTubeThickness("Headertube", CBPotD.Text, CBHeaderMat.Text, UnitProps.CBOP.Text)
                    TBPotWT.Text = wallthickness.ToString.Replace(".", ",")
                End If
            Else
                MsgBox("Missing data for wallthickness of header tubes")
            End If
            If CBNippleMat.Text <> "" Then
                If CBNippleOutD.Text <> "" Then
                    wallthickness = Database.GetTubeThickness("Headertube", CBNippleOutD.Text, CBNippleMat.Text, UnitProps.CBOP.Text)
                    TBNippleOutWT.Text = wallthickness.ToString.Replace(".", ",")
                End If
                If CBNippleInD.Text <> "" Then
                    wallthickness = Database.GetTubeThickness("Headertube", CBNippleInD.Text, CBNippleMat.Text, UnitProps.CBOP.Text)
                    TBNippleInWT.Text = wallthickness.ToString.Replace(".", ",")
                End If
            Else
                MsgBox("Missing data for wallthickness of connection tubes")
            End If

        Catch ex As Exception
            MsgBox("Error getting wall thickness of header tube")
        End Try
    End Sub

    Private Sub CheckFlange_CheckedChanged(sender As Object, e As EventArgs) Handles CheckFlange.CheckedChanged
        CBFlange.Visible = CheckFlange.Checked
        selectedconsys.HasFTCon = CheckFlange.Checked
        If General.isProdavit And CheckFlange.Checked Then
            Dim flangetype As String = CSData.GetFlangeTypebyERP(PCFData.GetValue("ThreadFlangeConnectionFC" + UnitProps.selectedcirc.CircuitNumber.ToString, "ERPCode"), selectedconsys.HeaderMaterial)
            CBFlange.Text = flangetype
        End If
    End Sub

    Private Sub BDrawing_Click(sender As Object, e As EventArgs) Handles BDrawing.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        If selectedconsys.ConSysFile.Fullfilename <> "" Then
            SEDrawing.CreateDrawingConSys(selectedconsys, UnitProps.selectedcoil, UnitProps.selectedcirc)

            BReset.Visible = True
        End If
    End Sub

    Private Sub BReset_Click(sender As Object, e As EventArgs) Handles BReset.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        'delete consys.dft files and reset the information in the consysdata class
        Try
            IO.File.Delete(selectedconsys.ConSysFile.Fullfilename)
            If IO.File.Exists(selectedconsys.ConSysFile.Fullfilename.Replace(".asm", ".dft")) Then
                IO.File.Delete(selectedconsys.ConSysFile.Fullfilename.Replace(".asm", ".dft"))
            End If
            Threading.Thread.Sleep(400)
            'delete all header and nipple tube files
            For Each fname As String In IO.Directory.GetFiles(General.currentjob.Workspace)
                If (fname.Contains("letHeader") Or fname.Contains("letNipple")) And fname.Contains("_" + selectedconsys.Circnumber.ToString + "_" + UnitProps.selectedcoil.Number.ToString + ".par") Then
                    IO.File.Delete(fname)
                    Threading.Thread.Sleep(200)
                End If
            Next

            'clear consys class
            For Each h In selectedconsys.InletHeaders
                h.StutzenDatalist.Clear()
                h.Xlist.Clear()
                h.Ylist.Clear()
                h.Nipplepositions.Clear()
                h.Nippletubes = CInt(TBNippleInQ.Text)
            Next

            For Each h In selectedconsys.OutletHeaders
                h.StutzenDatalist.Clear()
                h.Xlist.Clear()
                h.Ylist.Clear()
                h.Nipplepositions.Clear()
                h.Nippletubes = CInt(TBNippleOutQ.Text)
            Next

            If selectedconsys.InletHeaders.Count > 1 Then
                Do
                    selectedconsys.InletHeaders.RemoveAt(selectedconsys.InletHeaders.Count - 1)
                Loop Until selectedconsys.InletHeaders.Count = 1
            End If

            If selectedconsys.OutletHeaders.Count > 1 Then
                Do
                    selectedconsys.OutletHeaders.RemoveAt(selectedconsys.OutletHeaders.Count - 1)
                Loop Until selectedconsys.OutletHeaders.Count = 1
            End If

            If selectedconsys.InletNipples.Count > 1 Then
                Do
                    selectedconsys.InletNipples.RemoveAt(selectedconsys.InletNipples.Count - 1)
                Loop Until selectedconsys.InletNipples.Count = 1
            End If

            If selectedconsys.OutletNipples.Count > 1 Then
                Do
                    selectedconsys.OutletNipples.RemoveAt(selectedconsys.OutletNipples.Count - 1)
                Loop Until selectedconsys.OutletNipples.Count = 1
            End If

            selectedconsys.Occlist.Clear()
            selectedconsys.CoverSheetCutouts.Clear()

        Catch ex As Exception
            Debug.Print(ex.ToString)
        Finally
            selectedconsys.ConSysFile.Fullfilename = ""
        End Try
    End Sub

    Private Sub RBInletPos_CheckedChanged(sender As Object, e As EventArgs) Handles RBInletPos.CheckedChanged
        General.CreateActionLogEntry("ConSysProps", "InletANS", "changed", RBInletPos.Checked.ToString)
        RBOutletPos.Checked = Not RBInletPos.Checked
        If RBInletPos.Checked Then
            selectedconsys.InletHeaders.First.OddLocation = "front"
            selectedconsys.OutletHeaders.First.OddLocation = "back"
            If selectedconsys.OutletNipples.Count > 0 Then
                If UnitProps.selectedcirc.Pressure <= 16 Then
                    selectedconsys.OutletNipples.First.Angle = 180
                    CBAngleOut.Text = "180"
                End If
            End If
            If selectedconsys.InletNipples.Count > 0 Then
                selectedconsys.InletNipples.First.Angle = 0
                CBAngleIn.Text = "0"
            End If
        Else
            selectedconsys.OutletHeaders.First.OddLocation = "front"
            selectedconsys.InletHeaders.First.OddLocation = "back"
            If selectedconsys.OutletNipples.Count > 0 Then
                If UnitProps.selectedcirc.Pressure <= 16 Then
                    selectedconsys.OutletNipples.First.Angle = 0
                    CBAngleOut.Text = "0"
                End If
            End If
            If selectedconsys.InletNipples.Count > 0 Then
                selectedconsys.InletNipples.First.Angle = 180
                CBAngleIn.Text = "180"
            End If
        End If
    End Sub

    Private Sub RBOutletPos_CheckedChanged(sender As Object, e As EventArgs) Handles RBOutletPos.CheckedChanged
        General.CreateActionLogEntry("ConSysProps", "OutletANS", "changed", RBOutletPos.Checked.ToString)
        RBInletPos.Checked = Not RBOutletPos.Checked
        If RBOutletPos.Checked Then
            selectedconsys.OutletHeaders.First.OddLocation = "front"
            selectedconsys.InletHeaders.First.OddLocation = "back"
            If selectedconsys.OutletNipples.Count > 0 Then
                If UnitProps.selectedcirc.Pressure <= 16 Then
                    selectedconsys.OutletNipples.First.Angle = 0
                    CBAngleOut.Text = "0"
                End If
            End If
            If selectedconsys.InletNipples.Count > 0 Then
                selectedconsys.InletNipples.First.Angle = 180
                CBAngleIn.Text = "180"
            End If
        Else
            selectedconsys.InletHeaders.First.OddLocation = "front"
            selectedconsys.OutletHeaders.First.OddLocation = "back"
            If selectedconsys.OutletNipples.Count > 0 Then
                If UnitProps.selectedcirc.Pressure <= 16 Then
                    selectedconsys.OutletNipples.First.Angle = 180
                    CBAngleOut.Text = "180"
                End If
            End If
            If selectedconsys.InletNipples.Count > 0 Then
                selectedconsys.InletNipples.First.Angle = 0
                CBAngleIn.Text = "0"
            End If
        End If
    End Sub

    Private Sub CBHeaderMat_TextChanged(sender As Object, e As EventArgs) Handles CBHeaderMat.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBHeaderMat.Text)
    End Sub

    Private Sub CBNippleMat_TextChanged(sender As Object, e As EventArgs) Handles CBNippleMat.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBNippleMat.Text)
    End Sub

    Private Sub CBAlignment_TextChanged(sender As Object, e As EventArgs) Handles CBAlignment.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBAlignment.Text)
    End Sub

    Private Sub CBEvap_TextChanged(sender As Object, e As EventArgs) Handles CBEvap.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBEvap.Text)
    End Sub

    Private Sub CBFlange_TextChanged(sender As Object, e As EventArgs) Handles CBFlange.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBFlange.Text)
    End Sub

    Private Sub CBHeaderOutD_TextChanged(sender As Object, e As EventArgs) Handles CBHeaderOutD.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBHeaderOutD.Text)
    End Sub

    Private Sub CBHeaderInD_TextChanged(sender As Object, e As EventArgs) Handles CBHeaderInD.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBHeaderInD.Text)
    End Sub

    Private Sub CBNippleOutD_TextChanged(sender As Object, e As EventArgs) Handles CBNippleOutD.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBNippleOutD.Text)
    End Sub

    Private Sub CBNippleInD_TextChanged(sender As Object, e As EventArgs) Handles CBNippleInD.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBNippleInD.Text)
    End Sub

    Private Sub CheckLength_CheckedChanged(sender As Object, e As EventArgs) Handles CheckLength.CheckedChanged
        If CBConSys.Text <> "" Then
            selectedconsys.ControlNipples = CheckLength.Checked
        End If
        General.CreateActionLogEntry("ConSysProps", "CheckLength", "changed", CheckLength.Checked.ToString)
    End Sub

    Private Sub CBConSys_TextChanged(sender As Object, e As EventArgs) Handles CBConSys.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBConSys.Text)
    End Sub

    Private Sub CBAngleOut_TextChanged(sender As Object, e As EventArgs) Handles CBAngleOut.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBAngleOut.Text)
    End Sub

    Private Sub CBAngleIn_TextChanged(sender As Object, e As EventArgs) Handles CBAngleIn.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBAngleIn.Text)
    End Sub

    Private Sub CheckValve_CheckedChanged(sender As Object, e As EventArgs) Handles CheckValve.CheckedChanged
        General.CreateActionLogEntry("ConSysProps", "CheckValve", "changed to ", CheckValve.Checked.ToString)
        selectedconsys.HasBallValve = CheckValve.Checked
        CBValveSize.Visible = CheckValve.Checked
        If General.isProdavit And CheckValve.Checked Then
            Dim size As String = CSData.GetValveSizeByERP(PCFData.GetValue("VentilationDrainBallValve", "ERPCode"))
            CBValveSize.Text = size
        ElseIf CheckValve.Checked = False Then
            selectedconsys.Valvesize = ""
        End If
    End Sub

    Private Sub CheckSensor_CheckedChanged(sender As Object, e As EventArgs) Handles CheckSensor.CheckedChanged
        General.CreateActionLogEntry("ConSysProps", "CheckSensor", "changed to ", CheckSensor.Checked.ToString)
        selectedconsys.HasSensor = CheckSensor.Checked
    End Sub

    Private Sub CBValveSize_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBValveSize.SelectedIndexChanged
        If CBValveSize.Text <> "" Then
            selectedconsys.Valvesize = CBValveSize.Text
        End If
    End Sub

    Private Sub CBValveSize_TextChanged(sender As Object, e As EventArgs) Handles CBValveSize.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBValveSize.Text)
    End Sub

End Class