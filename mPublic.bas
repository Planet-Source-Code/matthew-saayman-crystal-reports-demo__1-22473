Attribute VB_Name = "mPublic"
Public Function Q(ByVal sField As String, Optional Char As String = "'") As String
    sField = Replace(sField, Char, Char + Char)
    sField = Char + sField + Char
    Q = sField
End Function

Public Function QC(ByVal sField As String, Optional Char As String = "'") As String
    sField = Replace(sField, Char, Char + Char)
    sField = Char + sField + Char
    QC = sField & ","
End Function
Public Function QCLF(ByVal sField As String, Optional Char As String = "'") As String
    sField = Replace(sField, Char, Char + Char)
    sField = Char + sField + Char
    QCLF = sField & "," & vbCrLf
End Function

Public Sub ShowError(Optional CalledFrom As String = "")
        If CalledFrom <> "" Then
           CalledFrom = CalledFrom & " - "
        End If
        MsgBox "Err:" & Err.Number & vbCrLf & Err.Description, _
            vbCritical, CalledFrom & "Error"
End Sub

Public Function Filename(ByVal path As String, Optional IncExt As Boolean = True) As String
Dim l, i, P, lp As Integer
Dim rString As String
    rString = ""
    P = InStr(1, path, "\")
    lp = P
    While P <> 0
        lp = P + 1
        P = InStr(lp, path, "\")
    Wend
    If lp = 0 Then
        rString = path
    Else
        rString = Mid(path, lp, Len(path) - lp + 1)
    End If
    
    If IncExt = False And InStr(1, rString, ".") > 0 Then
        rString = Mid(rString, 1, InStr(1, rString, ".") - 1)
    End If
    
    Filename = rString
End Function

Public Function FilePath(ByVal path As String) As String
Dim l, i, P, lp As Integer
Dim rString As String
    rString = ""
    P = InStr(1, path, "\")
    lp = P
    While P <> 0
        lp = P + 1
        P = InStr(lp, path, "\")
    Wend
    If lp = 0 Then
        rString = ""
    Else
        rString = Left(path, lp - 1)
    End If
    
    
    FilePath = rString
End Function

Public Function FileExt(ByVal path As String, _
                        Optional IncludePoint As Boolean = True) As String
Dim l, i, P, lp As Integer
Dim rString As String
    rString = ""
    P = InStr(1, path, ".")
    lp = P
    While P <> 0
        lp = P + 1
        P = InStr(lp, path, ".")
    Wend
    If lp = 0 Then
        rString = ""
    Else
        If IncludePoint = True Then
            rString = "." & Mid(path, lp, Len(path) - lp + 1)
        Else
            rString = Mid(path, lp, Len(path) - lp + 1)
        End If
    End If
    FileExt = rString
End Function


