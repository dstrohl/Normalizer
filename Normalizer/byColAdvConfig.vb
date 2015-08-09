Public Class byColAdvConfigControl
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    ' **************************************************************************************************
    ' **
    ' **  Control change handlers
    ' **
    ' **************************************************************************************************

    Public Sub byColHeaderColControl_ValueChanged(sender As Object, e As EventArgs) Handles byColHeaderColControl.ValueChanged
        'Column Header Change
        Globals.ThisAddIn.WorkingColumns.HeaderColCount = byColHeaderColControl.Value
        UpdateSelGrpOptions()
        UpdatePreview()
    End Sub


    Private Sub btnRefreshPreview_Click(sender As Object, e As EventArgs) Handles btnRefreshPreview.Click
        ' Refresh Button
        Globals.ThisAddIn.WorkingColumns.AllContent = Globals.ThisAddIn.Application.Selection
        RefreshControlsFromData()
        UpdatePreview()
    End Sub

    Private Sub byColHeaderRowsControl_ValueChanged(sender As Object, e As EventArgs) Handles byColHeaderRowsControl.ValueChanged
        ' Row Headers
        Globals.ThisAddIn.WorkingColumns.HeaderRowCount = byColHeaderRowsControl.Value
        UpdatePreview()
    End Sub

    Private Sub btnBcPreview_CheckedChanged(sender As Object, e As EventArgs) Handles btnBcPreview.CheckedChanged
        ' Preview button clicked
        Globals.ThisAddIn.ByColSamplePane.Visible = TryCast(sender, System.Windows.Forms.CheckBox).Checked
        If Globals.ThisAddIn.ByColSamplePane.Visible Then
            Globals.ThisAddIn.WorkingColumns.AllContent = Globals.ThisAddIn.Application.Selection
            RefreshControlsFromData()
            UpdatePreview()
        End If

    End Sub

    Private Sub DataColumnGroupSize_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataColumnGroupSize.SelectedIndexChanged
        ' Data Column Group Size
        Globals.ThisAddIn.WorkingColumns.DataColGroupSize = Val(DataColumnGroupSize.Text)
        UpdatePreview()
    End Sub

    Private Sub ByColDataHeadersControl_TextChanged(sender As Object, e As EventArgs) Handles ByColDataHeadersControl.TextChanged
        'UpdateArray(Globals.ThisAddIn.WorkingColumns.OldDataColLabel, Split(ByColDataHeadersControl.Text))

    End Sub

    Private Sub byColOldDataHeadersControl_TextChanged(sender As Object, e As EventArgs) Handles byColOldDataHeadersControl.TextChanged
        'UpdateArray(Globals.ThisAddIn.WorkingColumns.OldHeaderColLabel, Split(byColOldDataHeadersControl.Text))

    End Sub

    ' **************************************************************************************************
    ' **
    ' **  Operational Methods
    ' **
    ' **************************************************************************************************

    Public Sub UpdateControlsMax()
        With Globals.ThisAddIn.WorkingColumns
            byColHeaderRowsControl.Maximum = .MaxHeaderRows
            byColHeaderColControl.Maximum = .MaxHeaderColumns
        End With
    End Sub

    Public Sub RefreshControlsFromData()
        With Globals.ThisAddIn.WorkingColumns
            byColHeaderRowsControl.Value = .HeaderRowCount
            byColHeaderColControl.Value = .HeaderColCount
            UpdateSelGrpOptions(.DataColGroupSize.ToString)
        End With
        UpdateControlsMax()
    End Sub

    'Public Sub Updateconfig()

    '    'ColGrpSizeControl.Text = Globals.ThisAddIn.WorkingColumns.ColGroupSizeOptions(byColGrpSize.Value)
    '    With Globals.ThisAddIn.NormConfig
    '        '.AllContent = Globals.ThisAddIn.Application.Selection
    '        '.NormalizeMethod = "Column"
    '        '.HeaderRowCount = byColHeaderRowsControl.Value
    '        '.HeaderColCount = byColHeaderColControl.Value
    '        'UpdateSelGrpOptions()
    '        .DataColGroupCount = Val(ColGrpSizeControl.Text)
    '    End With
    '    'Globals.ThisAddIn.WorkingColumns.Setup()
    '    'byColHeaderRowsControl.Maximum = Globals.ThisAddIn.WorkingColumns.Rows - 1
    '    'byColHeaderColControl.Maximum = Globals.ThisAddIn.WorkingColumns.Columns - 1
    '    'byColGrpSize.Maximum = Globals.ThisAddIn.WorkingColumns.ColGroupSizeOptions.Count

    'End Sub


    Public Sub UpdatePreview()


        If Not Globals.ThisAddIn.ByColSamplePanel Is Nothing Then
            If Globals.ThisAddIn.ByColSamplePanel.Visible And Globals.ThisAddIn.WorkingColumns.DataLoaded Then

                Globals.ThisAddIn.ByColSamplePanel.LoadData()

                'With Globals.ThisAddIn.NormConfig
                '    WriteDebug("Header All Content Range: " & .AllContent.Address)
                '    WriteDebug("Header Header Cols: " & Val(.HeaderRowCount))
                '    WriteDebug("Header Header Rows: " & Val(.HeaderColCount))
                '    WriteDebug("Header Data Groups: " & Val(.DataColGroupCount))
                '    WriteDebug("Old Data Labels: " & .OldDataColLabel.ToString)
                '    WriteDebug("Old Header Labels: " & .OldHeaderColLabel.ToString)
                'End With
                'With Globals.ThisAddIn.WorkingColumns
                '    WriteDebug("all rows: " & Val(Globals.ThisAddIn.WorkingColumns.Rows))
                '    WriteDebug("all cols: " & Val(Globals.ThisAddIn.WorkingColumns.Columns))
                '    WriteDebug("possible col groups: " & Val(Globals.ThisAddIn.WorkingColumns.ColGroupSizeOptions.Count))
                'End With

                'Globals.ThisAddIn.ByColSamplePanel.SetupDisplay(Val(ColGrpSizeControl.Text), byColHeaderRowsControl.Value, byColHeaderColControl.Value)
            End If
        End If

    End Sub

    Public Sub UpdateSelGrpOptions(Optional ByVal OldOption As String = Nothing)
        'Dim OldOption As String = DataColumnGroupSize.Text
        If OldOption = Nothing Then
            If DataColumnGroupSize.Text = Nothing Then
                OldOption = "1"
            Else
                OldOption = DataColumnGroupSize.Text
            End If
        End If

        DataColumnGroupSize.Items.Clear()

        For Each Size As String In Globals.ThisAddIn.WorkingColumns.ColGroupSizeOptions
            DataColumnGroupSize.Items.Add(Size.ToString)
        Next

        If Globals.ThisAddIn.WorkingColumns.ColGroupSizeOptions.Contains(Val(OldOption)) Then
            DataColumnGroupSize.SelectedItem = OldOption
        Else
            DataColumnGroupSize.SelectedItem = "1"
        End If

    End Sub

End Class

