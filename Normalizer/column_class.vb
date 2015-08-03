Imports Normalizer.normalizerCommon
Imports System.Collections.Generic



Public MustInherit Class ColumnClassBase
    ' Class for individual columns
    Public ColHeaderRange As Excel.Range
    Public ColDataRange As Excel.Range
    Public CurColumn As Long
    Public StartCol As Long
    Public EndCol As Long
    Public ColCount As Long
    Public RowCount As Long
    Public DataRowCount As Long
    Public HeaderRowCount As Long
    'Public MyHeaderStyle As Excel.Style
    'Public MyDataStyle As Excel.Style
    Public Styled As Boolean = False
    Public Sub Setup(Optional ByVal Col As Long = 1)
        SetCols(Col)

        ColCount = EndCol - StartCol
        RowCount = Globals.ThisAddIn.NormConfig.AllContent.Rows.CountLarge
        HeaderRowCount = Globals.ThisAddIn.NormConfig.HeaderRowCount
        DataRowCount = RowCount - HeaderRowCount

        ColHeaderRange = Globals.ThisAddIn.NormConfig.AllContent.Range(getRCText(StartCol, 1, EndCol, Globals.ThisAddIn.NormConfig.HeaderRowCount))
        ColDataRange = Globals.ThisAddIn.NormConfig.AllContent.Range(getRCText(StartCol, 2, EndCol, RowCount))
    End Sub

    Public Sub ApplyStyles()
        ColDataRange.Style = MyDataStyle
        ColDataRange.ApplyOutlineStyles()

        ColDataRange.Borders(Excel.XlBordersIndex.xlEdgeBottom).Color = Globals.ThisAddIn.NormConfig.OutsideBorder_Color
        ColDataRange.Borders(Excel.XlBordersIndex.xlEdgeTop).Color = Globals.ThisAddIn.NormConfig.OutsideBorder_Color
        ColDataRange.Borders(Excel.XlBordersIndex.xlEdgeLeft).Color = Globals.ThisAddIn.NormConfig.OutsideBorder_Color
        ColDataRange.Borders(Excel.XlBordersIndex.xlEdgeRight).Color = Globals.ThisAddIn.NormConfig.OutsideBorder_Color

        ColDataRange.Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Globals.ThisAddIn.NormConfig.OutsideBorder_Weight
        ColDataRange.Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Globals.ThisAddIn.NormConfig.OutsideBorder_Weight
        ColDataRange.Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Globals.ThisAddIn.NormConfig.OutsideBorder_Weight
        ColDataRange.Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Globals.ThisAddIn.NormConfig.OutsideBorder_Weight

        ColHeaderRange.Style = MyHeaderStyle
        ColHeaderRange.ApplyOutlineStyles()

        ColHeaderRange.Borders(Excel.XlBordersIndex.xlEdgeBottom).Color = Globals.ThisAddIn.NormConfig.OutsideBorder_Color
        ColHeaderRange.Borders(Excel.XlBordersIndex.xlEdgeTop).Color = Globals.ThisAddIn.NormConfig.OutsideBorder_Color
        ColHeaderRange.Borders(Excel.XlBordersIndex.xlEdgeLeft).Color = Globals.ThisAddIn.NormConfig.OutsideBorder_Color
        ColHeaderRange.Borders(Excel.XlBordersIndex.xlEdgeRight).Color = Globals.ThisAddIn.NormConfig.OutsideBorder_Color

        ColHeaderRange.Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Globals.ThisAddIn.NormConfig.OutsideBorder_Weight
        ColHeaderRange.Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Globals.ThisAddIn.NormConfig.OutsideBorder_Weight
        ColHeaderRange.Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Globals.ThisAddIn.NormConfig.OutsideBorder_Weight
        ColHeaderRange.Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Globals.ThisAddIn.NormConfig.OutsideBorder_Weight

        Styled = True

    End Sub

    Public Sub ClearStyles()
        'Dim tmpBackupRange As Excel.Range

        With Globals.ThisAddIn.NormConfig
            .BackupFormatSheet.Range(ColHeaderRange.Address).Copy()
            ColHeaderRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats)

            .BackupFormatSheet.Range(ColDataRange.Address).Copy()
            ColDataRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats)

        End With
        Styled = False
    End Sub
    Public Sub ChangeColGroup(col As Long)
        Dim restyle As Boolean = False
        If Styled Then
            ClearStyles()
            restyle = True
        End If
        SetCols(col)
        ColHeaderRange = Globals.ThisAddIn.NormConfig.AllContent.Range(getRCText(StartCol, 1, EndCol, Globals.ThisAddIn.NormConfig.HeaderRowCount))
        ColDataRange = Globals.ThisAddIn.NormConfig.AllContent.Range(getRCText(StartCol, 2, EndCol, RowCount))
        If restyle Then
            ApplyStyles()
        End If
    End Sub

    Public MustOverride ReadOnly Property MyDataStyle() As Excel.Style

    Public MustOverride ReadOnly Property MyHeaderStyle() As Excel.Style

    Overridable Sub SetCols(col As Long)
        StartCol = col
        EndCol = col
    End Sub

    Public ReadOnly Property Columns() As Long
        Get
            Return ColCount
        End Get

    End Property
    Public ReadOnly Property Rows() As Long
        Get
            Return RowCount
        End Get

    End Property

    Public ReadOnly Property HeaderRows() As Long
        Get
            Return HeaderRowCount
        End Get

    End Property

    Public ReadOnly Property DataRows() As Long
        Get
            Return DataRowCount
        End Get

    End Property



End Class

Public Class DataColumClass
    Inherits ColumnClassBase

    Overrides Sub SetCols(col As Long)
        With Globals.ThisAddIn.NormConfig
            StartCol = .HeaderColCount + (col * .DataColGroupCount)
            EndCol = StartCol + .DataColGroupCount - 1
        End With

    End Sub
    Overrides ReadOnly Property MyDataStyle() As Excel.Style
        Get
            Return Globals.ThisAddIn.NormConfig.DataColumnStyle
        End Get
    End Property
    Overrides ReadOnly Property MyHeaderStyle() As Excel.Style
        Get
            Return Globals.ThisAddIn.NormConfig.DataHeaderStyle
        End Get
    End Property

    Public ReadOnly Property DataRange() As Excel.Range
        Get
            Return ColDataRange
        End Get
    End Property

    Public Function OldHeader(col As Integer) As Excel.Range
        Dim row As Integer = 1
        If col > ColCount Then
            col = col Mod ColCount
            row = col / ColCount
        End If

        Return ColHeaderRange.Item(row, col)
    End Function


    Public Function NewDataHeader(col As Integer) As String
        Return Globals.ThisAddIn.NormConfig.OldDataColLabel(col)
    End Function

    Public Function HeaderColHeader(col As Integer) As String
        Return Globals.ThisAddIn.NormConfig.OldHeaderColLabel(col)
    End Function

End Class

Public Class HeaderColumClass
    Inherits ColumnClassBase

    Overrides Sub SetCols(col As Long)
        StartCol = col
        EndCol = Globals.ThisAddIn.NormConfig.HeaderColCount
    End Sub

    Overrides ReadOnly Property MyDataStyle() As Excel.Style
        Get
            Return Globals.ThisAddIn.NormConfig.CommonDataStyle
        End Get
    End Property
    Overrides ReadOnly Property MyHeaderStyle() As Excel.Style
        Get
            Return Globals.ThisAddIn.NormConfig.CommonHeaderStyle
        End Get
    End Property

    Public ReadOnly Property DataRange() As Excel.Range
        Get
            Return ColDataRange
        End Get
    End Property

    Public ReadOnly Property HeaderRange() As Excel.Range
        Get
            Return ColHeaderRange
        End Get
    End Property

End Class

Public Class ColumnsClass
    Public ColCount As Long
    Public RowCount As Long
    Public DataRowCount As Long
    Public DataColCount As Long
    Public HeaderRowCount As Long
    Public HeaderColCount As Long
    Public DataGroupCount As Long
    Public NewHeaderColumnCount As Integer
    Public DataColGroupCount As Integer
    Public HeaderColGroupObj As New HeaderColumClass
    Public DataColGroupObj As New DataColumClass
    Public CurDataColGroupNum As Long = 1
    Public ColGroupSizeOptions As New List(Of Integer) From {1}
    'Public StyleBackupSheet As Excel.Worksheet
    Public PreviewPreped As Boolean = False

    Public Sub New()
        'Dim ColGroupSizeOptions As New List(Of Integer) From {1}
        'ColGroupSizeOptions.Add(1)
    End Sub
    Public Sub Setup()
        ColCount = Globals.ThisAddIn.NormConfig.AllContent.Columns.CountLarge
        RowCount = Globals.ThisAddIn.NormConfig.AllContent.Rows.CountLarge
        DataRowCount = RowCount - Globals.ThisAddIn.NormConfig.HeaderRowCount
        DataColCount = ColCount - Globals.ThisAddIn.NormConfig.HeaderColCount
        DataGroupCount = ColCount / Globals.ThisAddIn.NormConfig.DataColGroupCount
        HeaderRowCount = Globals.ThisAddIn.NormConfig.HeaderRowCount
        HeaderColCount = Globals.ThisAddIn.NormConfig.HeaderColCount
        DataColGroupCount = Globals.ThisAddIn.NormConfig.DataColGroupCount
        NewHeaderColumnCount = HeaderRowCount * DataColGroupCount
        HeaderColGroupObj.Setup()
        DataColGroupObj.Setup(CurDataColGroupNum)
        For i As Integer = 1 To DataColCount
            If DataColCount Mod i = 0 Then
                ColGroupSizeOptions.Add(i)
            End If
        Next
    End Sub

    Public Sub SetColGroup(col As Integer)
        If CurDataColGroupNum <> col Then
            CurDataColGroupNum = col
            DataColGroupObj.ChangeColGroup(CurDataColGroupNum)
        End If
    End Sub

    Public Sub Preview(Col As Long)
        If Not PreviewPreped Then

            With Globals.ThisAddIn.NormConfig
                Dim ActSheet As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
                ActSheet.Copy(After:=ActSheet)
                .BackupFormatSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(ActSheet.Index + 1)

                .BackupFormatSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden

                .DataColumnStyle = Globals.ThisAddIn.Application.ActiveWorkbook.Styles.Add("Normalizer Data Column")
                With .DataColumnStyle
                    .Interior.Color = Globals.ThisAddIn.NormConfig.DataColumnBkColor
                    .Interior.Pattern = Excel.XlPattern.xlPatternSolid
                    .Borders.Weight = Globals.ThisAddIn.NormConfig.InsideBorder_Weight
                    .Borders.Color = Globals.ThisAddIn.NormConfig.InsideBorder_Color
                    .IncludeAlignment = False
                    .IncludeFont = False
                    .IncludeNumber = False
                    .IncludeProtection = False
                End With


                .DataHeaderStyle = Globals.ThisAddIn.Application.ActiveWorkbook.Styles.Add("Normalizer Data Header")
                With .DataColumnStyle
                    .Interior.Color = Globals.ThisAddIn.NormConfig.DataHeaderBkColor
                    .Interior.Pattern = Excel.XlPattern.xlPatternSolid
                    .Borders.Weight = Globals.ThisAddIn.NormConfig.InsideBorder_Weight
                    .Borders.Color = Globals.ThisAddIn.NormConfig.InsideBorder_Color
                    .IncludeAlignment = False
                    .IncludeFont = False
                    .IncludeNumber = False
                    .IncludeProtection = False
                End With

                .CommonDataStyle = Globals.ThisAddIn.Application.ActiveWorkbook.Styles.Add("Normalizer Common Data")
                With .DataColumnStyle
                    .Interior.Color = Globals.ThisAddIn.NormConfig.CommonDataBkColor
                    .Interior.Pattern = Excel.XlPattern.xlPatternSolid
                    .Borders.Weight = Globals.ThisAddIn.NormConfig.InsideBorder_Weight
                    .Borders.Color = Globals.ThisAddIn.NormConfig.InsideBorder_Color
                    .IncludeAlignment = False
                    .IncludeFont = False
                    .IncludeNumber = False
                    .IncludeProtection = False
                End With

                .CommonHeaderStyle = Globals.ThisAddIn.Application.ActiveWorkbook.Styles.Add("Normalizer Common Header")
                With .DataColumnStyle
                    .Interior.Color = Globals.ThisAddIn.NormConfig.CommonHeaderBkColor
                    .Interior.Pattern = Excel.XlPattern.xlPatternSolid
                    .Borders.Weight = Globals.ThisAddIn.NormConfig.InsideBorder_Weight
                    .Borders.Color = Globals.ThisAddIn.NormConfig.InsideBorder_Color
                    .IncludeAlignment = False
                    .IncludeFont = False
                    .IncludeNumber = False
                    .IncludeProtection = False
                End With
                PreviewPreped = True
            End With
            HeaderColGroupObj.ApplyStyles()
            DataColGroupObj.ApplyStyles()

        End If
        If CurDataColGroupNum <> Col Then
            CurDataColGroupNum = Col
            DataColGroupObj.ChangeColGroup(CurDataColGroupNum)
        End If


    End Sub

    Public Sub CleanUpPreview()
        With Globals.ThisAddIn.NormConfig
            .BackupFormatSheet.Range(.AllContent.Address).Copy()
            .AllContent.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats)
            .BackupFormatSheet.Delete()
            .CommonHeaderStyle.Delete()
            .CommonDataStyle.Delete()
            .DataColumnStyle.Delete()
            .DataHeaderStyle.Delete()
        End With


        PreviewPreped = False
    End Sub

    Public ReadOnly Property Verify As List(Of String)
        Get
            Dim tmpRet As New List(Of String)
            Dim Maddr As String


            If DataColCount < 1 Then
                tmpRet.Add("Header size not setup correctly, (data columns < 0)")
                Return tmpRet
            End If

            If DataColCount Mod DataColGroupCount <> 0 Then
                tmpRet.Add("Data size does not evenly fit groups")
                Return tmpRet
            End If

            'check for merged cells and error out
            For Each c In Globals.ThisAddIn.NormConfig.AllContent
                If c.MergeCells Then
                    Maddr = c.Address
                    tmpRet.Add("At least one Merged Cell at: " & Maddr & " This might cause problems")
                End If
            Next c

            Return tmpRet

        End Get
    End Property


    Public ReadOnly Property Columns() As Long
        Get
            Return ColCount
        End Get

    End Property
    Public ReadOnly Property Rows() As Long
        Get
            Return RowCount
        End Get

    End Property

    Public ReadOnly Property DataColumnGroups() As Long
        Get
            Return DataGroupCount
        End Get

    End Property

    Public ReadOnly Property CommonDataColumns() As Long
        Get
            Return HeaderColCount
        End Get

    End Property

    Public ReadOnly Property NewHeaderColumns() As Long
        Get
            Return NewHeaderColumnCount
        End Get

    End Property

    Public ReadOnly Property HeaderRows() As Long
        Get
            Return HeaderRowCount
        End Get

    End Property

    Public ReadOnly Property DataRows() As Long
        Get
            Return DataRowCount
        End Get

    End Property

    Public ReadOnly Property CommonData() As Excel.Range
        Get
            Return HeaderColGroupObj
        End Get
    End Property

    Public Function DataColGroup(colGroup As Long) As Excel.Range
        colGroup = FBIndex(colGroup, DataColGroupCount)
        If CurDataColGroupNum <> colGroup Then
            CurDataColGroupNum = colGroup
            DataColGroupObj.ChangeColGroup(CurDataColGroupNum)
        End If
        Return DataColGroupObj
    End Function


End Class
