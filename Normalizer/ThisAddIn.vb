Public Class ThisAddIn
    Private byColTaskPanel As byColAdvConfigControl
    Private byDelTaskPanel As byDelAdvConfigControl
    Private byColSampleTaskPanel As byColSampleControl

    Private WithEvents byColTaskPaneView As Microsoft.Office.Tools.CustomTaskPane
    Private WithEvents byDelTaskPaneView As Microsoft.Office.Tools.CustomTaskPane
    Private WithEvents byColSampleTaskPaneView As Microsoft.Office.Tools.CustomTaskPane

    'Public NormConfig As ConfigSettings = New ConfigSettings
    Public WorkingColumns As ColumnsClass = New ColumnsClass

    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup

        System.Diagnostics.Debug.Write("Normalizer Loaded")

        byColTaskPanel = New byColAdvConfigControl()
        byColTaskPaneView = Me.CustomTaskPanes.Add(byColTaskPanel, "Normalize By Column")
        byColTaskPaneView.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop
        byColTaskPaneView.Height = 200
        byColTaskPaneView.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoVertical

        byDelTaskPanel = New byDelAdvConfigControl()
        byDelTaskPaneView = Me.CustomTaskPanes.Add(byDelTaskPanel, "Normalize By Delimeter")
        byDelTaskPaneView.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop
        byDelTaskPaneView.Height = 120
        byDelTaskPaneView.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoVertical

        byColSampleTaskPanel = New byColSampleControl()
        byColSampleTaskPaneView = Me.CustomTaskPanes.Add(byColSampleTaskPanel, "Sample Normalizer Results")
        byColSampleTaskPaneView.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        byColSampleTaskPaneView.Width = 500
        byColSampleTaskPaneView.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal


        'Set configuration base settings
        'Me.NormConfig.AllContent = Globals.ThisAddIn.Application.Selection

        'By Column Configuration Styles
        'Dim newStyle As Excel.Style = Globals.ThisAddIn.Application.ActiveWorkbook.Styles.Add("NormalizerDataColumn")




    End Sub


    'Private Sub byColTaskPaneValue_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles byColTaskPaneView.VisibleChanged
    '    If byColTaskPaneView.Visible Then
    '        byColTaskPanel.UpdatePreview()
    '    End If

    'End Sub

    Private Sub byColSampleTaskPaneValue_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles byColSampleTaskPaneView.VisibleChanged
        byColTaskPanel.btnBcPreview.Checked = byColSampleTaskPaneView.Visible
    End Sub
    Public ReadOnly Property ByColAdvConfigPane() As Microsoft.Office.Tools.CustomTaskPane
        Get
            Return byColTaskPaneView
        End Get
    End Property

    Public ReadOnly Property ByDelAdvConfigPane() As Microsoft.Office.Tools.CustomTaskPane
        Get
            Return byDelTaskPaneView
        End Get
    End Property
    Public ReadOnly Property ByColSamplePane() As Microsoft.Office.Tools.CustomTaskPane
        Get
            Return byColSampleTaskPaneView
        End Get
    End Property

    Public ReadOnly Property ByColSamplePanel() As byColSampleControl
        Get
            Return byColSampleTaskPanel
        End Get
    End Property


End Class

