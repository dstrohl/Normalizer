<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class byColAdvConfigControl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.byColHeaderRowsControl = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.byColHeaderColControl = New System.Windows.Forms.NumericUpDown()
        Me.ColorDialog1 = New System.Windows.Forms.ColorDialog()
        Me.byColOldDataHeadersControl = New System.Windows.Forms.TextBox()
        Me.ByColDataHeadersControl = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btnBcNorm = New System.Windows.Forms.Button()
        Me.btnBcPreview = New System.Windows.Forms.CheckBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.grpbox1 = New System.Windows.Forms.GroupBox()
        Me.byColGrpSize = New System.Windows.Forms.NumericUpDown()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.PreviewColumn = New System.Windows.Forms.TrackBar()
        Me.btnRefreshPreview = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.ColGrpSizeControl = New System.Windows.Forms.TextBox()
        CType(Me.byColHeaderRowsControl, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.byColHeaderColControl, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.grpbox1.SuspendLayout()
        CType(Me.byColGrpSize, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.PreviewColumn, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'byColHeaderRowsControl
        '
        Me.byColHeaderRowsControl.Location = New System.Drawing.Point(106, 21)
        Me.byColHeaderRowsControl.Name = "byColHeaderRowsControl"
        Me.byColHeaderRowsControl.Size = New System.Drawing.Size(82, 22)
        Me.byColHeaderRowsControl.TabIndex = 0
        Me.byColHeaderRowsControl.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 53)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Columns"
        '
        'byColHeaderColControl
        '
        Me.byColHeaderColControl.Location = New System.Drawing.Point(106, 51)
        Me.byColHeaderColControl.Name = "byColHeaderColControl"
        Me.byColHeaderColControl.Size = New System.Drawing.Size(82, 22)
        Me.byColHeaderColControl.TabIndex = 3
        Me.byColHeaderColControl.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'byColOldDataHeadersControl
        '
        Me.byColOldDataHeadersControl.Dock = System.Windows.Forms.DockStyle.Fill
        Me.byColOldDataHeadersControl.Location = New System.Drawing.Point(232, 117)
        Me.byColOldDataHeadersControl.Multiline = True
        Me.byColOldDataHeadersControl.Name = "byColOldDataHeadersControl"
        Me.byColOldDataHeadersControl.Size = New System.Drawing.Size(569, 108)
        Me.byColOldDataHeadersControl.TabIndex = 6
        Me.byColOldDataHeadersControl.Text = "Moved Headers"
        '
        'ByColDataHeadersControl
        '
        Me.ByColDataHeadersControl.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ByColDataHeadersControl.Location = New System.Drawing.Point(232, 3)
        Me.ByColDataHeadersControl.Multiline = True
        Me.ByColDataHeadersControl.Name = "ByColDataHeadersControl"
        Me.ByColDataHeadersControl.Size = New System.Drawing.Size(569, 108)
        Me.ByColDataHeadersControl.TabIndex = 7
        Me.ByColDataHeadersControl.Text = "Moved Data"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label3.Location = New System.Drawing.Point(48, 114)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(178, 114)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Old Data Headers Label(s)"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label4.Location = New System.Drawing.Point(120, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(106, 114)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Data Header(s)"
        '
        'btnBcNorm
        '
        Me.btnBcNorm.Enabled = False
        Me.btnBcNorm.Location = New System.Drawing.Point(15, 101)
        Me.btnBcNorm.Name = "btnBcNorm"
        Me.btnBcNorm.Size = New System.Drawing.Size(84, 28)
        Me.btnBcNorm.TabIndex = 11
        Me.btnBcNorm.Text = "Normalize"
        Me.btnBcNorm.UseVisualStyleBackColor = True
        '
        'btnBcPreview
        '
        Me.btnBcPreview.AccessibleRole = System.Windows.Forms.AccessibleRole.TitleBar
        Me.btnBcPreview.Appearance = System.Windows.Forms.Appearance.Button
        Me.btnBcPreview.AutoSize = True
        Me.btnBcPreview.Location = New System.Drawing.Point(25, 23)
        Me.btnBcPreview.Name = "btnBcPreview"
        Me.btnBcPreview.Size = New System.Drawing.Size(67, 27)
        Me.btnBcPreview.TabIndex = 12
        Me.btnBcPreview.Text = "Preview"
        Me.btnBcPreview.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 4
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 200.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 120.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 120.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.grpbox1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.GroupBox1, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.GroupBox2, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.GroupBox3, 3, 0)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1256, 255)
        Me.TableLayoutPanel1.TabIndex = 13
        '
        'grpbox1
        '
        Me.grpbox1.Controls.Add(Me.ColGrpSizeControl)
        Me.grpbox1.Controls.Add(Me.byColGrpSize)
        Me.grpbox1.Controls.Add(Me.Label5)
        Me.grpbox1.Controls.Add(Me.byColHeaderRowsControl)
        Me.grpbox1.Controls.Add(Me.Label1)
        Me.grpbox1.Controls.Add(Me.Label2)
        Me.grpbox1.Controls.Add(Me.byColHeaderColControl)
        Me.grpbox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpbox1.Location = New System.Drawing.Point(3, 3)
        Me.grpbox1.Name = "grpbox1"
        Me.grpbox1.Size = New System.Drawing.Size(194, 249)
        Me.grpbox1.TabIndex = 0
        Me.grpbox1.TabStop = False
        Me.grpbox1.Text = "Headers"
        '
        'byColGrpSize
        '
        Me.byColGrpSize.Location = New System.Drawing.Point(147, 111)
        Me.byColGrpSize.Name = "byColGrpSize"
        Me.byColGrpSize.Size = New System.Drawing.Size(41, 22)
        Me.byColGrpSize.TabIndex = 6
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(11, 83)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(103, 17)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Col Group Size"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 23)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(42, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Rows"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.AutoSize = True
        Me.GroupBox1.Controls.Add(Me.TableLayoutPanel2)
        Me.GroupBox1.Location = New System.Drawing.Point(203, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(810, 249)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "New Header Text"
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.TableLayoutPanel2.ColumnCount = 2
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 28.59281!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 71.40719!))
        Me.TableLayoutPanel2.Controls.Add(Me.ByColDataHeadersControl, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.byColOldDataHeadersControl, 1, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.Label4, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Label3, 0, 1)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(3, 18)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 2
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(804, 228)
        Me.TableLayoutPanel2.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.PreviewColumn)
        Me.GroupBox2.Controls.Add(Me.btnRefreshPreview)
        Me.GroupBox2.Controls.Add(Me.btnBcPreview)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(1019, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(114, 249)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Preview"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(22, 88)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(82, 17)
        Me.Label7.TabIndex = 16
        Me.Label7.Text = "Data Group"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(35, 71)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(57, 17)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "Preview"
        '
        'PreviewColumn
        '
        Me.PreviewColumn.Location = New System.Drawing.Point(10, 108)
        Me.PreviewColumn.Name = "PreviewColumn"
        Me.PreviewColumn.Size = New System.Drawing.Size(98, 53)
        Me.PreviewColumn.TabIndex = 14
        Me.PreviewColumn.Value = 1
        '
        'btnRefreshPreview
        '
        Me.btnRefreshPreview.Location = New System.Drawing.Point(25, 156)
        Me.btnRefreshPreview.Name = "btnRefreshPreview"
        Me.btnRefreshPreview.Size = New System.Drawing.Size(67, 23)
        Me.btnRefreshPreview.TabIndex = 13
        Me.btnRefreshPreview.Text = "Refresh"
        Me.btnRefreshPreview.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.btnBcNorm)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(1139, 3)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(114, 249)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Normalize"
        '
        'ColGrpSizeControl
        '
        Me.ColGrpSizeControl.Enabled = False
        Me.ColGrpSizeControl.Location = New System.Drawing.Point(62, 111)
        Me.ColGrpSizeControl.Name = "ColGrpSizeControl"
        Me.ColGrpSizeControl.Size = New System.Drawing.Size(61, 22)
        Me.ColGrpSizeControl.TabIndex = 7
        '
        'byColAdvConfigControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "byColAdvConfigControl"
        Me.Size = New System.Drawing.Size(1256, 255)
        CType(Me.byColHeaderRowsControl, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.byColHeaderColControl, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.grpbox1.ResumeLayout(False)
        Me.grpbox1.PerformLayout()
        CType(Me.byColGrpSize, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.PreviewColumn, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents byColHeaderRowsControl As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents byColHeaderColControl As System.Windows.Forms.NumericUpDown
    Friend WithEvents ColorDialog1 As System.Windows.Forms.ColorDialog
    Friend WithEvents byColOldDataHeadersControl As System.Windows.Forms.TextBox
    Friend WithEvents ByColDataHeadersControl As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnBcNorm As System.Windows.Forms.Button
    Friend WithEvents btnBcPreview As System.Windows.Forms.CheckBox
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents grpbox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnRefreshPreview As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents byColGrpSize As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents PreviewColumn As System.Windows.Forms.TrackBar
    Friend WithEvents ColGrpSizeControl As System.Windows.Forms.TextBox
End Class
