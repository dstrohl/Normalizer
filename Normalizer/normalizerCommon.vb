Module normalizerCommon

    Public First As String = "First"
    Public Second As String = "Second"
    Public SecondLast As String = "SecondLast"
    Public Last As String = "Last"
    Public Left As String = "Left"
    Public Right As String = "Right"
    Public TopLeft As String = "TopLeft"
    Public TopRight As String = "TopRight"
    Public LowerLeft As String = "LowerLeft"
    Public LowerRight As String = "LowerRight"


    Public Sub RangeDebug(ByRef range As Excel.Range, name As String)
        Dim rows As Long = range.Rows.CountLarge
        Dim cols As Long = range.Columns.CountLarge
        Dim address As String = range.Address
        'Dim ul_row As Long = range.Item(1, 1).row
        'Dim ul_col As Long = range.Item(1, 1).column
        'Dim ll_row As Long = range.Item(cols, rows).row
        'Dim ll_col As Long = range.Item(cols, rows).column

        WriteDebug("Range: [" & name & "] address : " & address & "(R=" & rows.ToString & "/C=" & cols.ToString & ")")
        'WriteDebug("Range: [" & name & "]    rows : " & rows.ToString)
        'WriteDebug("Range: [" & name & "]    cols : " & cols.ToString)
        'WriteDebug("Range: [" & name & "]  ul_row : " & ul_row.ToString)
        'WriteDebug("Range: [" & name & "]  ul_col : " & ul_col.ToString)
        'WriteDebug("Range: [" & name & "]  ll_row : " & ll_row.ToString)
        'WriteDebug("Range: [" & name & "]  ll_col : " & ll_col.ToString)


    End Sub

    Public Function Distance(a As Object, b As Object) As Object
        If a - b > 0 Then
            Return a - b
        Else
            Return (a - b) * -1
        End If
    End Function


    Public Function Nearest(a As Object, list As Object) As Object
        Dim tmp_ret As Object = 99999999
        For Each b As Object In list
            If Distance(a, b) < tmp_ret Then tmp_ret = Distance(a, b)
        Next
        Return tmp_ret
    End Function


    Public Sub WriteDebug(debugstr As String)
        Diagnostics.Debug.Write("Normalizer : " & debugstr & " |" & vbCrLf)
    End Sub


    Function ConvertToLetter(iCol As Long) As String
        Dim iAlpha As Integer
        Dim iRemainder As Integer
        Dim tmpRet As String = ""

        iAlpha = Int(iCol / 27)
        iRemainder = iCol - (iAlpha * 26)
        If iAlpha > 0 Then
            tmpRet = Chr(iAlpha + 64)
        End If
        If iRemainder > 0 Then
            tmpRet = tmpRet & Chr(iRemainder + 64)
        End If

        Return tmpRet
    End Function

    Public Function getRCText(col As Long, row As Long, Optional toCol As Long = Nothing, Optional toRow As Long = Nothing) As String
        Dim tmpRet As String = ""

        tmpRet = ConvertToLetter(col) & row
        If toCol <> Nothing Then
            tmpRet = tmpRet & ":" & ConvertToLetter(toCol) & toRow
        End If

        Return tmpRet

    End Function


    Public Function FBIndex(index As Long, max As Long) As Long
        If index < 0 Then
            Return max + index + 1
        ElseIf index > max
            Return max
        End If
        Return index

    End Function

    Public Sub UpdateArray(ByRef oldArray As Object, newArray As Object)
        ReDim oldArray(newArray.length)
        oldArray = newArray

    End Sub
End Module
