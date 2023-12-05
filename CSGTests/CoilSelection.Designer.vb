<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CoilSelection
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.UnitTree = New System.Windows.Forms.TreeView()
        Me.BConfirm = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(29, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(225, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Select Coil Item from BOM of Production Order"
        '
        'UnitTree
        '
        Me.UnitTree.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.UnitTree.Location = New System.Drawing.Point(32, 48)
        Me.UnitTree.Name = "UnitTree"
        Me.UnitTree.Size = New System.Drawing.Size(396, 143)
        Me.UnitTree.TabIndex = 1
        '
        'BConfirm
        '
        Me.BConfirm.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BConfirm.Location = New System.Drawing.Point(350, 16)
        Me.BConfirm.Name = "BConfirm"
        Me.BConfirm.Size = New System.Drawing.Size(78, 23)
        Me.BConfirm.TabIndex = 2
        Me.BConfirm.Text = "OK"
        Me.BConfirm.UseVisualStyleBackColor = True
        '
        'CoilSelection
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(459, 203)
        Me.Controls.Add(Me.BConfirm)
        Me.Controls.Add(Me.UnitTree)
        Me.Controls.Add(Me.Label1)
        Me.Name = "CoilSelection"
        Me.Text = "Select Coil Item"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents UnitTree As TreeView
    Friend WithEvents BConfirm As Button
End Class
