<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MRName
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.LMRC = New System.Windows.Forms.Label()
        Me.BConfirm = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CBMRName = New System.Windows.Forms.ComboBox()
        Me.CBMRSuffix = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'LMRC
        '
        Me.LMRC.AutoSize = True
        Me.LMRC.Location = New System.Drawing.Point(12, 22)
        Me.LMRC.Name = "LMRC"
        Me.LMRC.Size = New System.Drawing.Size(102, 13)
        Me.LMRC.TabIndex = 4
        Me.LMRC.Text = "Model Range Name"
        '
        'BConfirm
        '
        Me.BConfirm.Location = New System.Drawing.Point(15, 83)
        Me.BConfirm.Name = "BConfirm"
        Me.BConfirm.Size = New System.Drawing.Size(193, 23)
        Me.BConfirm.TabIndex = 3
        Me.BConfirm.Text = "OK"
        Me.BConfirm.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 51)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Model Range Suffix"
        '
        'CBMRName
        '
        Me.CBMRName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBMRName.FormattingEnabled = True
        Me.CBMRName.Items.AddRange(New Object() {"GACV", "GxHV", "GxVV", "GxHC", "GxVC", "GxDV"})
        Me.CBMRName.Location = New System.Drawing.Point(120, 19)
        Me.CBMRName.Name = "CBMRName"
        Me.CBMRName.Size = New System.Drawing.Size(88, 21)
        Me.CBMRName.TabIndex = 7
        '
        'CBMRSuffix
        '
        Me.CBMRSuffix.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBMRSuffix.FormattingEnabled = True
        Me.CBMRSuffix.Location = New System.Drawing.Point(120, 46)
        Me.CBMRSuffix.Name = "CBMRSuffix"
        Me.CBMRSuffix.Size = New System.Drawing.Size(88, 21)
        Me.CBMRSuffix.TabIndex = 8
        '
        'MRName
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(230, 122)
        Me.Controls.Add(Me.CBMRSuffix)
        Me.Controls.Add(Me.CBMRName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LMRC)
        Me.Controls.Add(Me.BConfirm)
        Me.Name = "MRName"
        Me.Text = "MRName"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LMRC As Label
    Friend WithEvents BConfirm As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents CBMRName As ComboBox
    Friend WithEvents CBMRSuffix As ComboBox
End Class
