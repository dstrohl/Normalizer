Module normalizerCommon
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

    Public Function getRCText(col As Long, row As Long, Optional toCol As Long = vbNull, Optional toRow As Long = vbNull) As String
        Dim tmpRet As String = ""

        tmpRet = ConvertToLetter(col) & row
        If toCol <> vbNull Then
            tmpRet = tmpRet & ":" & ConvertToLetter(toCol) & toRow
        End If

        Return tmpRet

    End Function


    Public Function FBIndex(index As Long, max As Long) As Long
        If index < 0 Then
            Return max - index + 1
        End If
        Return index

    End Function

    Public Sub UpdateArray(ByRef oldArray As Object, newArray As Object)
        ReDim oldArray(newArray.length)
        oldArray = newArray

    End Sub
End Module
