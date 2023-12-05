<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class OrderProps
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.BGo = New System.Windows.Forms.Button()
        Me.LCSG = New System.Windows.Forms.Label()
        Me.CBCSGMode = New System.Windows.Forms.ComboBox()
        Me.RBAPO = New System.Windows.Forms.RadioButton()
        Me.RBEU = New System.Windows.Forms.RadioButton()
        Me.TBNr = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TBPos = New System.Windows.Forms.TextBox()
        Me.TBOrder = New System.Windows.Forms.TextBox()
        Me.CBUnit = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.BJob = New System.Windows.Forms.Button()
        Me.CheckTest = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'BGo
        '
        Me.BGo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BGo.Location = New System.Drawing.Point(160, 173)
        Me.BGo.Name = "BGo"
        Me.BGo.Size = New System.Drawing.Size(121, 47)
        Me.BGo.TabIndex = 6
        Me.BGo.Text = "GO"
        Me.BGo.UseVisualStyleBackColor = True
        '
        'LCSG
        '
        Me.LCSG.AutoSize = True
        Me.LCSG.Location = New System.Drawing.Point(21, 149)
        Me.LCSG.Name = "LCSG"
        Me.LCSG.Size = New System.Drawing.Size(67, 13)
        Me.LCSG.TabIndex = 72
        Me.LCSG.Text = "Select Mode"
        '
        'CBCSGMode
        '
        Me.CBCSGMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBCSGMode.FormattingEnabled = True
        Me.CBCSGMode.Items.AddRange(New Object() {"Desktop", "Batch", "Copy Order", "Continue local Order"})
        Me.CBCSGMode.Location = New System.Drawing.Point(160, 146)
        Me.CBCSGMode.Name = "CBCSGMode"
        Me.CBCSGMode.Size = New System.Drawing.Size(121, 21)
        Me.CBCSGMode.TabIndex = 5
        '
        'RBAPO
        '
        Me.RBAPO.AutoSize = True
        Me.RBAPO.Location = New System.Drawing.Point(89, 12)
        Me.RBAPO.Name = "RBAPO"
        Me.RBAPO.Size = New System.Drawing.Size(47, 17)
        Me.RBAPO.TabIndex = 70
        Me.RBAPO.TabStop = True
        Me.RBAPO.Text = "APO"
        Me.RBAPO.UseVisualStyleBackColor = True
        '
        'RBEU
        '
        Me.RBEU.AutoSize = True
        Me.RBEU.Location = New System.Drawing.Point(24, 12)
        Me.RBEU.Name = "RBEU"
        Me.RBEU.Size = New System.Drawing.Size(59, 17)
        Me.RBEU.TabIndex = 69
        Me.RBEU.TabStop = True
        Me.RBEU.Text = "Europe"
        Me.RBEU.UseVisualStyleBackColor = True
        '
        'TBNr
        '
        Me.TBNr.Location = New System.Drawing.Point(160, 120)
        Me.TBNr.Name = "TBNr"
        Me.TBNr.Size = New System.Drawing.Size(121, 20)
        Me.TBNr.TabIndex = 4
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(21, 123)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 13)
        Me.Label6.TabIndex = 68
        Me.Label6.Text = "Project Number"
        '
        'TBPos
        '
        Me.TBPos.Location = New System.Drawing.Point(160, 94)
        Me.TBPos.Name = "TBPos"
        Me.TBPos.Size = New System.Drawing.Size(121, 20)
        Me.TBPos.TabIndex = 3
        '
        'TBOrder
        '
        Me.TBOrder.Location = New System.Drawing.Point(160, 67)
        Me.TBOrder.Name = "TBOrder"
        Me.TBOrder.Size = New System.Drawing.Size(121, 20)
        Me.TBOrder.TabIndex = 2
        '
        'CBUnit
        '
        Me.CBUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBUnit.FormattingEnabled = True
        Me.CBUnit.Items.AddRange(New Object() {"GxH C/V - Family", "GxD C/V - Family", "GACV - Family", "GADC - Family", "Condenser Flat/Vertical", "Condenser VShape", "Evaporator", "Evaporator (2 Coils)"})
        Me.CBUnit.Location = New System.Drawing.Point(160, 40)
        Me.CBUnit.Name = "CBUnit"
        Me.CBUnit.Size = New System.Drawing.Size(121, 21)
        Me.CBUnit.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(21, 97)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(44, 13)
        Me.Label4.TabIndex = 52
        Me.Label4.Text = "Position"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(21, 70)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(73, 13)
        Me.Label3.TabIndex = 51
        Me.Label3.Text = "Order Number"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(21, 43)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 13)
        Me.Label2.TabIndex = 50
        Me.Label2.Text = "Unit Type"
        '
        'BJob
        '
        Me.BJob.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BJob.Location = New System.Drawing.Point(24, 173)
        Me.BJob.Name = "BJob"
        Me.BJob.Size = New System.Drawing.Size(112, 47)
        Me.BJob.TabIndex = 73
        Me.BJob.Text = "Job List"
        Me.BJob.UseVisualStyleBackColor = True
        '
        'CheckTest
        '
        Me.CheckTest.AutoSize = True
        Me.CheckTest.Location = New System.Drawing.Point(220, 12)
        Me.CheckTest.Name = "CheckTest"
        Me.CheckTest.Size = New System.Drawing.Size(61, 17)
        Me.CheckTest.TabIndex = 74
        Me.CheckTest.Text = "LNTest"
        Me.CheckTest.UseVisualStyleBackColor = True
        Me.CheckTest.Visible = False
        '
        'OrderProps
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(296, 232)
        Me.Controls.Add(Me.CheckTest)
        Me.Controls.Add(Me.BJob)
        Me.Controls.Add(Me.LCSG)
        Me.Controls.Add(Me.CBCSGMode)
        Me.Controls.Add(Me.RBAPO)
        Me.Controls.Add(Me.RBEU)
        Me.Controls.Add(Me.TBNr)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.TBPos)
        Me.Controls.Add(Me.TBOrder)
        Me.Controls.Add(Me.CBUnit)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.BGo)
        Me.Name = "OrderProps"
        Me.Text = "CSG Main"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents BGo As Button
    Friend WithEvents LCSG As Label
    Friend WithEvents CBCSGMode As ComboBox
    Friend WithEvents RBAPO As RadioButton
    Friend WithEvents RBEU As RadioButton
    Friend WithEvents TBNr As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents TBPos As TextBox
    Friend WithEvents TBOrder As TextBox
    Friend WithEvents CBUnit As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents BJob As Button
    Friend WithEvents CheckTest As CheckBox
End Class
