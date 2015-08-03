Public Class ConfigSettings

    'Base Configuration
    Public NormalizeMethod As String = "Column"
    Public AllContent As Excel.Range

    'By Delimeter Configuration
    Public Delimeter As String = ","
    Public DelimeterColumn As Long = 1

    'By Column Configuration
    Public HeaderRowCount As Integer = 1
    Public HeaderColCount As Integer = 1
    Public DataColGroupCount As Integer = 1
    Public OldHeaderColLabel() As String = {"Moved Headers"}
    Public OldDataColLabel() As String = {"Moved Data"}

    'Defining styles for preview
    Public DataColumnStyle As Excel.Style
    Public DataHeaderStyle As Excel.Style
    Public CommonDataStyle As Excel.Style
    Public CommonHeaderStyle As Excel.Style

    'Defining basic colors
    Public DataColumnBkColor As Integer = Drawing.ColorTranslator.ToOle(Drawing.ColorTranslator.FromHtml("#B4C6E7"))
    Public DataHeaderBkColor As Integer = Drawing.ColorTranslator.ToOle(Drawing.ColorTranslator.FromHtml("#C6E0B4"))
    Public CommonDataBkColor As Integer = Drawing.ColorTranslator.ToOle(Drawing.ColorTranslator.FromHtml("#FFE699"))
    Public CommonHeaderBkColor As Integer = Drawing.ColorTranslator.ToOle(Drawing.ColorTranslator.FromHtml("#F8CBAD"))



    Public InsideBorder_Weight As Integer = Excel.XlBorderWeight.xlHairline
    Public InsideBorder_Color As Integer = Drawing.ColorTranslator.ToOle(Drawing.Color.White)

    Public OutsideBorder_Weight As Integer = Excel.XlBorderWeight.xlThick
    Public OutsideBorder_Color As Integer = Drawing.ColorTranslator.ToOle(Drawing.Color.Black)

    'defining borders
    Public NormSampleBorder As Excel.Borders


    'Bacup sheet for formatts
    Public BackupFormatSheet As Excel.Worksheet


End Class
