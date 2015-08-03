Partial Class Normalize
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim RibbonDropDownItemImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl2 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl3 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Me.tbNormalize = Me.Factory.CreateRibbonTab
        Me.rgStep1 = Me.Factory.CreateRibbonGroup
        Me.Label2 = Me.Factory.CreateRibbonLabel
        Me.BySelector = Me.Factory.CreateRibbonDropDown
        Me.rgStep2_byCol = Me.Factory.CreateRibbonGroup
        Me.ByColFDC = Me.Factory.CreateRibbonBox
        Me.Label1 = Me.Factory.CreateRibbonLabel
        Me.EditBox8 = Me.Factory.CreateRibbonEditBox
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.btnByColAdv = Me.Factory.CreateRibbonToggleButton
        Me.btnByColPreview = Me.Factory.CreateRibbonToggleButton
        Me.btnNext2 = Me.Factory.CreateRibbonButton
        Me.rgStep2_bydel = Me.Factory.CreateRibbonGroup
        Me.ByDelFDC = Me.Factory.CreateRibbonEditBox
        Me.byDelDelimeter = Me.Factory.CreateRibbonEditBox
        Me.btnNext2a = Me.Factory.CreateRibbonButton
        Me.rgStep3 = Me.Factory.CreateRibbonGroup
        Me.btnNormalize = Me.Factory.CreateRibbonButton
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.TestCopy = Me.Factory.CreateRibbonButton
        Me.Restore = Me.Factory.CreateRibbonButton
        Me.tbNormalize.SuspendLayout()
        Me.rgStep1.SuspendLayout()
        Me.rgStep2_byCol.SuspendLayout()
        Me.ByColFDC.SuspendLayout()
        Me.rgStep2_bydel.SuspendLayout()
        Me.rgStep3.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tbNormalize
        '
        Me.tbNormalize.Groups.Add(Me.rgStep1)
        Me.tbNormalize.Groups.Add(Me.rgStep2_byCol)
        Me.tbNormalize.Groups.Add(Me.rgStep2_bydel)
        Me.tbNormalize.Groups.Add(Me.rgStep3)
        Me.tbNormalize.Groups.Add(Me.Group1)
        Me.tbNormalize.Label = "NORMALIZE"
        Me.tbNormalize.Name = "tbNormalize"
        '
        'rgStep1
        '
        Me.rgStep1.Items.Add(Me.Label2)
        Me.rgStep1.Items.Add(Me.BySelector)
        Me.rgStep1.Label = "Step 1"
        Me.rgStep1.Name = "rgStep1"
        '
        'Label2
        '
        Me.Label2.Label = "Normalize By"
        Me.Label2.Name = "Label2"
        '
        'BySelector
        '
        RibbonDropDownItemImpl1.Label = "Select One"
        RibbonDropDownItemImpl2.Label = "Column"
        RibbonDropDownItemImpl3.Label = "Delimeter"
        Me.BySelector.Items.Add(RibbonDropDownItemImpl1)
        Me.BySelector.Items.Add(RibbonDropDownItemImpl2)
        Me.BySelector.Items.Add(RibbonDropDownItemImpl3)
        Me.BySelector.Label = "Normalize by:"
        Me.BySelector.Name = "BySelector"
        Me.BySelector.ShowLabel = False
        '
        'rgStep2_byCol
        '
        Me.rgStep2_byCol.Items.Add(Me.ByColFDC)
        Me.rgStep2_byCol.Items.Add(Me.Separator1)
        Me.rgStep2_byCol.Items.Add(Me.btnByColAdv)
        Me.rgStep2_byCol.Items.Add(Me.btnByColPreview)
        Me.rgStep2_byCol.Items.Add(Me.btnNext2)
        Me.rgStep2_byCol.Label = "Step 2"
        Me.rgStep2_byCol.Name = "rgStep2_byCol"
        Me.rgStep2_byCol.Visible = False
        '
        'ByColFDC
        '
        Me.ByColFDC.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical
        Me.ByColFDC.Items.Add(Me.Label1)
        Me.ByColFDC.Items.Add(Me.EditBox8)
        Me.ByColFDC.Name = "ByColFDC"
        '
        'Label1
        '
        Me.Label1.Label = "First Data Column"
        Me.Label1.Name = "Label1"
        '
        'EditBox8
        '
        Me.EditBox8.Label = "byColFDC"
        Me.EditBox8.Name = "EditBox8"
        Me.EditBox8.ScreenTip = "Either Number or Letter"
        Me.EditBox8.ShowLabel = False
        Me.EditBox8.Text = Nothing
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'btnByColAdv
        '
        Me.btnByColAdv.Label = "Advanced Config"
        Me.btnByColAdv.Name = "btnByColAdv"
        '
        'btnByColPreview
        '
        Me.btnByColPreview.Label = "Show Preview"
        Me.btnByColPreview.Name = "btnByColPreview"
        '
        'btnNext2
        '
        Me.btnNext2.Label = "Next >"
        Me.btnNext2.Name = "btnNext2"
        '
        'rgStep2_bydel
        '
        Me.rgStep2_bydel.Items.Add(Me.ByDelFDC)
        Me.rgStep2_bydel.Items.Add(Me.byDelDelimeter)
        Me.rgStep2_bydel.Items.Add(Me.btnNext2a)
        Me.rgStep2_bydel.Label = "Step 2"
        Me.rgStep2_bydel.Name = "rgStep2_bydel"
        Me.rgStep2_bydel.Visible = False
        '
        'ByDelFDC
        '
        Me.ByDelFDC.Label = "Data Column"
        Me.ByDelFDC.Name = "ByDelFDC"
        Me.ByDelFDC.Text = Nothing
        '
        'byDelDelimeter
        '
        Me.byDelDelimeter.Label = "Delimiter"
        Me.byDelDelimeter.Name = "byDelDelimeter"
        Me.byDelDelimeter.Text = Nothing
        '
        'btnNext2a
        '
        Me.btnNext2a.Label = "Next >"
        Me.btnNext2a.Name = "btnNext2a"
        '
        'rgStep3
        '
        Me.rgStep3.Items.Add(Me.btnNormalize)
        Me.rgStep3.Label = "Step 3"
        Me.rgStep3.Name = "rgStep3"
        Me.rgStep3.Visible = False
        '
        'btnNormalize
        '
        Me.btnNormalize.Enabled = False
        Me.btnNormalize.Label = "Normalize"
        Me.btnNormalize.Name = "btnNormalize"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.TestCopy)
        Me.Group1.Items.Add(Me.Restore)
        Me.Group1.Label = "Group1"
        Me.Group1.Name = "Group1"
        '
        'TestCopy
        '
        Me.TestCopy.Label = "Copy"
        Me.TestCopy.Name = "TestCopy"
        '
        'Restore
        '
        Me.Restore.Label = "Restore"
        Me.Restore.Name = "Restore"
        '
        'Normalize
        '
        Me.Name = "Normalize"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.tbNormalize)
        Me.tbNormalize.ResumeLayout(False)
        Me.tbNormalize.PerformLayout()
        Me.rgStep1.ResumeLayout(False)
        Me.rgStep1.PerformLayout()
        Me.rgStep2_byCol.ResumeLayout(False)
        Me.rgStep2_byCol.PerformLayout()
        Me.ByColFDC.ResumeLayout(False)
        Me.ByColFDC.PerformLayout()
        Me.rgStep2_bydel.ResumeLayout(False)
        Me.rgStep2_bydel.PerformLayout()
        Me.rgStep3.ResumeLayout(False)
        Me.rgStep3.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents tbNormalize As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents rgStep2_bydel As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ByDelFDC As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents byDelDelimeter As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents rgStep2_byCol As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ByColFDC As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents EditBox8 As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents Label1 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents btnNormalize As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents rgStep1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnByColPreview As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents TestCopy As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Restore As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnByColAdv As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents BySelector As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents rgStep3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Label2 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents btnNext2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnNext2a As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Normalize
        Get
            Return Me.GetRibbon(Of Normalize)()
        End Get
    End Property
End Class
