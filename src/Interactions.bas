Attribute VB_Name = "Interactions"
Sub Refresh_Click()
    ActiveSheet.Range("B2:C255").ClearContents
    ActiveSheet.Calculate
    ActiveSheet.Range("B2").Select
    Call LDQueryExpand(ActiveSheet.Range("token").value, ActiveSheet.Range("query").value)
    ActiveSheet.Calculate
End Sub

Sub Refresh_Click2()
    ActiveSheet.Range("F3:K255").ClearContents
    ActiveSheet.Calculate
    ActiveSheet.Range("F3").Select
    Call LDQueryExpand(Sheets("Simple Query").Range("token").value, ActiveSheet.Range("query2").value)
    ActiveSheet.Calculate
End Sub

Public Sub LDQueryExpand(token As String, query As String)
    Dim v As Variant
    Dim R As Range
    Dim Size(1) As Long
    
    v = LDQuery(token, query)
    
    Set R = ActiveCell
    
    If VarType(v) = vbString Then
        R.value = v
        GoTo ret:
    End If
    
    Size(0) = UBound(v, 1)
    Size(1) = UBound(v, 2)
    
    ' expand the range to fit our array
    Set R = R.Resize(Size(0) + 1, Size(1) + 1)
    
    ' with this string we check if we would overwrite other cells by expanding our array
    MsgBoxResult = vbOK
    
    If WorksheetFunction.CountBlank(R) <> R.Count Then
        
        MsgBoxResult = MsgBox("WARNING: Content will be overwritten in " & R.Address, vbOKCancel, "OVERWRITE WARNING")
    End If
    
    If MsgBoxResult = vbOK Then
        R.value = v
    End If

ret:
End Sub
