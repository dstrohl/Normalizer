Imports Microsoft.Office.Tools.Ribbon

Public Class Normalize
    Public ColumnsObj As ColumnsClass = New ColumnsClass
    Public NormConfig As ConfigSettings = New ConfigSettings

    'Private Sub TestCopy_Click(sender As Object, e As RibbonControlEventArgs) Handles TestCopy.Click
    '    NormConfig.HeaderColCount = 1
    '    NormConfig.HeaderRowCount = 1
    '    NormConfig.DataColGroupCount = 1
    '    NormConfig.OldHeaderColLabel = {"TestHdr"}
    '    NormConfig.OldDataColLabel = {"TestDataHdr"}
    '    NormConfig.AllContent = Globals.ThisAddIn.Application.Selection

    '    ColumnsObj.Setup(NormConfig)
    '    ColumnsObj.BackupStyles()
    'End Sub

    'Private Sub Restore_Click(sender As Object, e As RibbonControlEventArgs) Handles Restore.Click
    '    ColumnsObj.restoreStyles()
    'End Sub

    Private Sub ShowHideStepPanes()
        If Globals.ThisAddIn.Application.Selection.rows.countlarge > 1 And Globals.ThisAddIn.Application.Selection.columns.countlarge > 1 Then

            If BySelector.SelectedItemIndex = 1 Then
                'Globals.Ribbons.Ribbon1.rgStep2_byCol.Visible = True
                Globals.ThisAddIn.ByColAdvConfigPane.Visible = True
                'Globals.Ribbons.Ribbon1.rgStep2_bydel.Visible = False
                Globals.ThisAddIn.ByDelAdvConfigPane.Visible = False
            ElseIf BySelector.SelectedItemIndex = 2 Then
                'Globals.Ribbons.Ribbon1.rgStep2_byCol.Visible = False
                'Globals.Ribbons.Ribbon1.rgStep2_bydel.Visible = True
                Globals.ThisAddIn.ByColAdvConfigPane.Visible = False
                Globals.ThisAddIn.ByDelAdvConfigPane.Visible = True
            Else
                Globals.ThisAddIn.ByColAdvConfigPane.Visible = False
                Globals.ThisAddIn.ByDelAdvConfigPane.Visible = False
            End If

        Else
            BySelector.SelectedItemIndex = 0
            MsgBox("Nothing Selected, select a range before trying to normalize.")
        End If

    End Sub

    Private Sub BySelector_SelectionChanged(sender As Object, e As RibbonControlEventArgs) Handles BySelector.SelectionChanged
        ShowHideStepPanes()
    End Sub
End Class

