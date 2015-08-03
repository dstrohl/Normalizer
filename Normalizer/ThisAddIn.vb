Public Class ThisAddIn
    Private byColTaskPanel As byColAdvConfigControl
    Private byDelTaskPanel As byDelAdvConfigControl

    Private WithEvents byColTaskPaneView As Microsoft.Office.Tools.CustomTaskPane
    Private WithEvents byDelTaskPaneView As Microsoft.Office.Tools.CustomTaskPane

    Public NormConfig As ConfigSettings = New ConfigSettings
    Public WorkingColumns As ColumnsClass = New ColumnsClass

    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        byColTaskPanel = New byColAdvConfigControl()
        byColTaskPaneView = Me.CustomTaskPanes.Add(byColTaskPanel, "Normalize By Column")
        byColTaskPaneView.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop
        byColTaskPaneView.Height = 120
        byColTaskPaneView.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoVertical

        byDelTaskPanel = New byDelAdvConfigControl()
        byDelTaskPaneView = Me.CustomTaskPanes.Add(byDelTaskPanel, "Normalize By Delimeter")
        byDelTaskPaneView.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop
        byDelTaskPaneView.Height = 120
        byDelTaskPaneView.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoVertical



        'Set configuration base settings
        'Me.NormConfig.AllContent = Globals.ThisAddIn.Application.Selection

        'By Column Configuration Styles
        'Dim newStyle As Excel.Style = Globals.ThisAddIn.Application.ActiveWorkbook.Styles.Add("NormalizerDataColumn")




    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub byColTaskPaneValue_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles byColTaskPaneView.VisibleChanged
        If byColTaskPaneView.Visible Then
            byColTaskPanel.Updateconfig()
        End If

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

End Class

