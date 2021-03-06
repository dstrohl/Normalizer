﻿Imports Normalizer.normalizerCommon
Imports System.Collections.Generic


Public Class CornerItem
    Dim Item As Excel.Range
    Public Sub New(ByRef range As Excel.Range, row As Integer, col As Integer)
        Me.Item = range.Item(row, col)
    End Sub

    Public ReadOnly Property Value() As String
        Get
            Return Item.Value2
        End Get
    End Property

    Public ReadOnly Property Column() As String
        Get
            Return Item.Column
        End Get
    End Property

    Public ReadOnly Property Row() As String
        Get
            Return Item.Row
        End Get
    End Property


End Class

' **************************************************************************************************
' **
' **  Column Base
' **
' **************************************************************************************************


Public MustInherit Class ColumnClassBase
    ' Class for individual columns
    'Public ColHeaderRange As Excel.Range
    'Public ColDataRange As Excel.Range
    'Public CurColumn As Long = 1
    'Public StartCol As Long
    'Public EndCol As Long
    'Public Columns As Long
    'Public Rows As Long
    'Public DataRows As Long
    'Public HeaderRows As Long
    'Public MyHeaderStyle As Excel.Style
    'Public MyDataStyle As Excel.Style
    'Public Styled As Boolean = False
    'Public CurrentDataColumnNumber
    Public _ColumnsParent As ColumnsClass

    'Private TopRow, BottomRow, LeftCol, RightCol As Long

    Public Sub New(ByRef Parent As ColumnsClass)
        _ColumnsParent = Parent
    End Sub


    'Public Sub Setup(Optional ByVal Col As Long = 1)
    '    'SetCols(Col)

    '    'Rows = Globals.ThisAddIn.NormConfig.AllContent.Rows.CountLarge
    '    'HeaderRows = Globals.ThisAddIn.NormConfig.HeaderRowCount
    '    'DataRows = Rows - HeaderRows

    '    'ColHeaderRange = Globals.ThisAddIn.NormConfig.AllContent.Range(getRCText(StartCol, 1, EndCol, Globals.ThisAddIn.NormConfig.HeaderRowCount))
    '    'ColDataRange = Globals.ThisAddIn.NormConfig.AllContent.Range(getRCText(StartCol, DataStartRow, EndCol, Rows))
    'End Sub

    Public ReadOnly Property ColDataRange() As Excel.Range
        Get
            WriteDebug("Getting ColDataRange : " & getRCText(StartCol, DataStartRow, EndCol, Rows))
            Return _ColumnsParent.AllContent.Range(getRCText(StartCol, DataStartRow, EndCol, Rows))
        End Get
    End Property


    Public ReadOnly Property ColHeaderRange() As Excel.Range
        Get
            WriteDebug("Getting ColHeaderRange : " & getRCText(StartCol, 1, EndCol, HeaderRows))
            Return _ColumnsParent.AllContent.Range(getRCText(StartCol, 1, EndCol, HeaderRows))
        End Get
    End Property

    Public MustOverride ReadOnly Property StartCol() As Long

    Public MustOverride ReadOnly Property EndCol() As Long

    Public ReadOnly Property Columns() As Long
        Get
            Return EndCol - StartCol + 1
        End Get

    End Property
    Public ReadOnly Property Rows() As Long
        Get
            Return _ColumnsParent.Rows
        End Get

    End Property

    Public ReadOnly Property DataStartRow() As Long
        Get
            Return HeaderRows + 1
        End Get
    End Property

    Public ReadOnly Property HeaderRows() As Long
        Get
            Return _ColumnsParent.HeaderRowCount
        End Get

    End Property

    Public ReadOnly Property DataRows() As Long
        Get
            Return Rows - HeaderRows
        End Get

    End Property



    Public Function Corner(direction As String, ByRef range As Excel.Range, Optional ByVal row_offset As Integer = 0, Optional ByVal col_offset As Integer = 0) As CornerItem
        Dim col, row As Long
        Dim row_max As Long = range.Rows.CountLarge
        Dim col_max As Long = range.Columns.CountLarge

        RangeDebug(range, direction & " Corner")
        Select Case direction
            Case TopLeft
                row = 1
                col = 1
            Case TopRight
                row = 1
                col = col_max
            Case LowerLeft
                row = row_max
                col = 1
            Case LowerRight
                row = row_max
                col = col_max
        End Select

        row = row + row_offset
        If row < 1 Then row = 1
        If row > row_max Then row = row_max


        col = col + col_offset
        If col < 1 Then col = 1
        If col > col_max Then col = col_max

        Dim tmpRet As New CornerItem(range, row, col)
        Return tmpRet

    End Function

    Public Function HeaderCorner(direction As String, Optional ByVal row_offset As Integer = 0, Optional ByVal col_offset As Integer = 0) As CornerItem
        RangeDebug(ColHeaderRange, "col header range")
        Return Corner(direction, ColHeaderRange, row_offset, col_offset)
    End Function

    Public Function DataCorner(direction As String, Optional ByVal row_offset As Integer = 0, Optional ByVal col_offset As Integer = 0) As CornerItem
        RangeDebug(ColDataRange, direction & " Data Corner")
        Return Corner(direction, ColDataRange, row_offset, col_offset)
    End Function

End Class

' **************************************************************************************************
' **
' **  Data Column
' **
' **************************************************************************************************


Public Class DataColumClass
    Inherits ColumnClassBase
    Public MaxColumnGroups As Long

    Public Sub New(ByRef Parent As ColumnsClass)
        MyBase.New(Parent)
    End Sub

    Public Overrides ReadOnly Property StartCol As Long
        Get
            Return _ColumnsParent.HeaderColCount + 1 + ((_ColumnsParent.CurDataColGroupNum - 1) * _ColumnsParent.DataColGroupSize)
        End Get
    End Property

    Public Overrides ReadOnly Property EndCol As Long
        Get
            Return StartCol + _ColumnsParent.DataColGroupSize - 1
        End Get
    End Property

    'Overrides Sub SetCols(col As Long)
    '    With Globals.ThisAddIn.NormConfig
    '        StartCol = .HeaderColCount + 1 + ((col - 1) * .DataColGroupCount)
    '        EndCol = StartCol + .DataColGroupCount - 1
    '        MaxColumnGroups = (.AllContent.Columns.CountLarge - .HeaderColCount) / .DataColGroupCount

    '    End With

    'End Sub

    'Public Sub ChangeColGroup(col As Long)
    '    With Globals.ThisAddIn.NormConfig
    '        MaxColumnGroups = (.AllContent.Columns.CountLarge - .HeaderColCount) / .DataColGroupCount
    '        WriteDebug("Max Column Groups: " & Val(MaxColumnGroups))
    '        Dim colGroup As Long = FBIndex(col, MaxColumnGroups)
    '        WriteDebug("Going to Column: : " & Val(MaxColumnGroups))
    '        If CurColumn <> colGroup Then
    '            CurColumn = colGroup

    '            SetCols(CurColumn)
    '            ColHeaderRange = .AllContent.Range(getRCText(StartCol, 1, EndCol, .HeaderRowCount))
    '            ColDataRange = .AllContent.Range(getRCText(StartCol, 2, EndCol, Rows))

    '        End If
    '    End With
    'End Sub



    'Public ReadOnly Property DataRange() As Excel.Range
    '    Get
    '        Return ColDataRange
    '    End Get
    'End Property

    Public Function OldHeader(col As Integer) As Excel.Range
        Dim row As Integer = 1
        If col > Columns Then
            col = col Mod Columns
            row = col / Columns
        End If

        Return ColHeaderRange.Item(row, col)
    End Function


    'Public Function NewDataHeader(col As Integer) As String
    '    Return Globals.ThisAddIn.NormConfig.OldDataColLabel(col)
    'End Function

    'Public Function HeaderColHeader(col As Integer) As String
    '    Return Globals.ThisAddIn.NormConfig.OldHeaderColLabel(col)
    'End Function



End Class

' **************************************************************************************************
' **
' **  Header Column
' **
' **************************************************************************************************


Public Class HeaderColumClass
    Inherits ColumnClassBase

    Public Sub New(ByRef Parent As ColumnsClass)
        MyBase.New(Parent)
    End Sub

    'Overrides Sub SetCols(col As Long)
    '    StartCol = col
    '    EndCol = Globals.ThisAddIn.NormConfig.HeaderColCount
    'End Sub

    Public Overrides ReadOnly Property StartCol As Long
        Get
            Return 1
        End Get
    End Property

    Public Overrides ReadOnly Property EndCol As Long
        Get
            Return _ColumnsParent.HeaderColCount
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


' **************************************************************************************************
' **
' **  All Working Columns
' **
' **************************************************************************************************


Public Class ColumnsClass

    Private _HeaderRowCount As Long = 1
    Private _HeaderColCount As Long = 1
    Private _DataColGroupSize As Integer = 1

    Private _AllContent As Excel.Range
    Public DataLoaded As Boolean = False

    Private _HeaderColGroupObj As HeaderColumClass
    Private _DataColGroupObj As DataColumClass

    Public Columns As Long
    Public Rows As Long
    Public DataRowCount As Long
    Public DataColCount As Long
    Public DataGroupCount As Long
    Public CurDataColGroupNum As Long = 1
    Public ColGroupSizeOptions As New List(Of Integer) From {1}

    Public Sub New()
        'Dim _HeaderColGroupObj As New HeaderColumClass(Parent:=Me)
        'Dim _DataColGroupObj As New DataColumClass(Parent:=Me)
    End Sub

    Public Property AllContent() As Excel.Range
        Get
            Return _AllContent
        End Get
        Set(value As Excel.Range)

            CurDataColGroupNum = 1


            If value.Columns.CountLarge < 2 Or value.Rows.CountLarge < 2 Then
                DataLoaded = False
                _AllContent = Nothing

                Columns = 0
                Rows = 0
                'DataRowCount = 0
                'DataGroupCount = 0

                _HeaderColGroupObj = Nothing
                _DataColGroupObj = Nothing


            Else
                _AllContent = value
                DataLoaded = True

                Columns = _AllContent.Columns.CountLarge
                Rows = _AllContent.Rows.CountLarge

                _HeaderColGroupObj = New HeaderColumClass(Me)
                _DataColGroupObj = New DataColumClass(Me)

            End If

            HeaderRowCount = 1
            HeaderColCount = 1
            DataColGroupSize = 1

        End Set
    End Property

    Public ReadOnly Property MaxHeaderColumns() As Long
        Get
            If DataLoaded Then
                Return Columns - 1
            Else
                Return 1
            End If

        End Get
    End Property

    Public ReadOnly Property MaxHeaderRows() As Long
        Get
            If DataLoaded Then
                Return Rows - 1
            Else
                Return 1
            End If
        End Get
    End Property


    Public Property HeaderColCount() As Long
        Get
            Return _HeaderColCount
        End Get
        Set(value As Long)
            _HeaderColCount = value
            If DataLoaded Then
                DataColCount = Columns - _HeaderColCount
                ColGroupSizeOptions.Clear()
                For i As Integer = 1 To DataColCount - 1
                    If DataColCount Mod i = 0 Then
                        ColGroupSizeOptions.Add(i)
                    End If
                Next

                If Not ColGroupSizeOptions.Contains(DataColGroupSize) Then
                    DataColGroupSize = 1
                End If

            Else
                DataColCount = 0
                DataColGroupSize = 1
                ColGroupSizeOptions = New List(Of Integer) From {1}
            End If

        End Set
    End Property

    Public Property HeaderRowCount() As Long
        Get
            Return _HeaderRowCount
        End Get
        Set(value As Long)
            _HeaderRowCount = value
            If DataLoaded Then
                DataRowCount = Rows - _HeaderRowCount
            Else
                DataRowCount = 0
            End If

        End Set
    End Property

    Public Property DataColGroupSize() As Long
        Get
            Return _DataColGroupSize
        End Get
        Set(value As Long)
            If ColGroupSizeOptions.Contains(value) Then
                _DataColGroupSize = value
            Else
                _DataColGroupSize = Nearest(value, ColGroupSizeOptions)
            End If

            If DataLoaded Then
                DataGroupCount = DataColCount / _DataColGroupSize
            Else
                DataGroupCount = 0
            End If
        End Set
    End Property

    Public ReadOnly Property NewHeaderColumnCount() As Long
        Get
            Return HeaderRowCount * _DataColGroupSize
        End Get
    End Property

    Public ReadOnly Property CommonData() As HeaderColumClass
        Get
            Return _HeaderColGroupObj
        End Get
    End Property

    Public Function DataColGroup(colGroup As Long) As DataColumClass
        colGroup = FBIndex(colGroup, DataGroupCount)
        If CurDataColGroupNum <> colGroup Then
            CurDataColGroupNum = colGroup
            '_DataColGroupObj.ChangeColGroup(CurDataColGroupNum)
        End If
        Return _DataColGroupObj
    End Function


End Class
