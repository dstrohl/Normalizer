Public Class byColAdvConfigControl
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        'Updateconfig()
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub byColHeaderColControl_ValueChanged(sender As Object, e As EventArgs) Handles byColHeaderColControl.ValueChanged
        Check4UpdateNeeded()
    End Sub


    Public Sub Updateconfig()
        ColGrpSizeControl.Text = Globals.ThisAddIn.WorkingColumns.ColGroupSizeOptions(byColGrpSize.Value)
        With Globals.ThisAddIn.NormConfig
            .AllContent = Globals.ThisAddIn.Application.Selection
            .NormalizeMethod = "Column"
            .HeaderRowCount = byColHeaderRowsControl.Value
            .HeaderColCount = byColHeaderColControl.Value
            .DataColGroupCount = Val(ColGrpSizeControl.Text)
            UpdateArray(.OldDataColLabel, Split(ByColDataHeadersControl.Text))
            UpdateArray(.OldHeaderColLabel, Split(byColOldDataHeadersControl.Text))
        End With
        Globals.ThisAddIn.WorkingColumns.Setup()
        byColHeaderRowsControl.Maximum = Globals.ThisAddIn.WorkingColumns.Rows - 1
        byColHeaderColControl.Maximum = Globals.ThisAddIn.WorkingColumns.Columns - 1
        byColGrpSize.Maximum = Globals.ThisAddIn.WorkingColumns.ColGroupSizeOptions.Count
    End Sub

    Private Sub UpdatePreview()
        Updateconfig()
        Globals.ThisAddIn.WorkingColumns.Preview(PreviewColumn.Value)
    End Sub

    Private Sub Check4UpdateNeeded()
        If btnBcPreview.Checked Then
            Dim update_needed As Boolean = False

            With Globals.ThisAddIn.NormConfig
                If Globals.ThisAddIn.Application.Selection.Equals(.AllContent) Then
                    update_needed = True
                ElseIf .NormalizeMethod <> "Column" Then
                    update_needed = True
                ElseIf .HeaderRowCount <> byColHeaderRowsControl.Value
                    update_needed = True
                ElseIf .HeaderColCount <> byColHeaderColControl.Value
                    update_needed = True
                ElseIf .DataColGroupCount <> val(ColGrpSizeControl.Text)
                    update_needed = True
                ElseIf Globals.ThisAddIn.WorkingColumns.CurDataColGroupNum <> PreviewColumn.Value
                    update_needed = True
                End If
            End With

            If update_needed Then
                btnRefreshPreview.Enabled = True
            End If
        End If

    End Sub

    Private Sub PreviewDataColumn_Scroll(sender As Object, e As EventArgs) Handles PreviewColumn.Scroll

    End Sub

    Protected Overrides Sub Finalize()
        Globals.ThisAddIn.WorkingColumns.CleanUpPreview()
        MyBase.Finalize()
    End Sub

    Private Sub byColAdvConfigControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnRefreshPreview_Click(sender As Object, e As EventArgs) Handles btnRefreshPreview.Click
        UpdatePreview()
    End Sub

    Private Sub byColGrpSize_ValueChanged(sender As Object, e As EventArgs) Handles byColGrpSize.ValueChanged
        ColGrpSizeControl.Text = Globals.ThisAddIn.WorkingColumns.ColGroupSizeOptions(byColGrpSize.Value)
        Check4UpdateNeeded()
    End Sub

    Private Sub byColHeaderRowsControl_ValueChanged(sender As Object, e As EventArgs) Handles byColHeaderRowsControl.ValueChanged
        Check4UpdateNeeded()
    End Sub

    Private Sub ByColDataHeadersControl_TextChanged(sender As Object, e As EventArgs) Handles ByColDataHeadersControl.TextChanged
        Check4UpdateNeeded()
    End Sub

    Private Sub byColOldDataHeadersControl_TextChanged(sender As Object, e As EventArgs) Handles byColOldDataHeadersControl.TextChanged
        Check4UpdateNeeded()
    End Sub

    Private Sub btnBcPreview_CheckedChanged(sender As Object, e As EventArgs) Handles btnBcPreview.CheckedChanged
        If btnBcPreview.Checked Then
            UpdatePreview()
            PreviewColumn.Enabled = True
            btnRefreshPreview.Enabled = True
        Else
            Globals.ThisAddIn.WorkingColumns.CleanUpPreview()
            PreviewColumn.Enabled = False
            btnRefreshPreview.Enabled = False
        End If
    End Sub
End Class

