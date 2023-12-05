<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ConfirmSave
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
        Me.BYes = New System.Windows.Forms.Button()
        Me.BNo = New System.Windows.Forms.Button()
        Me.LInfo = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'BYes
        '
        Me.BYes.Location = New System.Drawing.Point(51, 70)
        Me.BYes.Name = "BYes"
        Me.BYes.Size = New System.Drawing.Size(75, 23)
        Me.BYes.TabIndex = 0
        Me.BYes.Text = "Yes"
        Me.BYes.UseVisualStyleBackColor = True
        '
        'BNo
        '
        Me.BNo.Location = New System.Drawing.Point(149, 70)
        Me.BNo.Name = "BNo"
        Me.BNo.Size = New System.Drawing.Size(75, 23)
        Me.BNo.TabIndex = 1
        Me.BNo.Text = "No"
        Me.BNo.UseVisualStyleBackColor = True
        '
        'LInfo
        '
        Me.LInfo.AutoSize = True
        Me.LInfo.Location = New System.Drawing.Point(24, 13)
        Me.LInfo.Name = "LInfo"
        Me.LInfo.Size = New System.Drawing.Size(248, 13)
        Me.LInfo.TabIndex = 2
        Me.LInfo.Text = "No connection system drawing for Coil 1 - Consys 1"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(110, 45)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Proceed?"
        '
        'ConfirmSave
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(278, 104)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LInfo)
        Me.Controls.Add(Me.BNo)
        Me.Controls.Add(Me.BYes)
        Me.Name = "ConfirmSave"
        Me.Text = "ConfirmSave"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents BYes As Button
    Friend WithEvents BNo As Button
    Friend WithEvents LInfo As Label
    Friend WithEvents Label1 As Label
End Class
