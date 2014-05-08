Attribute VB_Name = "LDashHelpers"

'The MIT License (MIT)

'Copyright (c) 2014 Loganis - iWebMa Limited

'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:

'The above copyright notice and this permission notice shall be included in
'all copies or substantial portions of the Software.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
'THE SOFTWARE.

Public Function LDTrf(keys As Range, valueTypes As Range, values As Range) As String
    
    If keys.Columns.Count > 1 Or valueTypes.Columns.Count > 1 Or values.Columns.Count > 1 Then GoTo Err
    
    Dim n As Integer
    n = keys.Rows.Count
    
    If n <> valueTypes.Rows.Count Or n <> values.Rows.Count Then
        GoTo Err
    End If
    
    Dim i As Integer
    Dim ret As String
    ret = "{"
    
    Dim key As String
    Dim valueType As String
    Dim value As String
    Dim channel As String
    
    For i = 1 To n
    
        key = keys(i, 1)
        valueType = valueTypes(i, 1)
        value = values(i, 1)
    
        If value <> "" Then
            If UCase(key) = ":CHA" Then
                channel = value
                GoTo Continue
            End If
            
            If UCase(key) = ":MET" And channel <> "" Then value = channel + ":" + value
            
            If UCase(valueType) = "STRING" Then
                ret = ret + " " + key + " """ + value + """"
            ElseIf UCase(valueType) = "DATE" Then
                ret = ret + " " + key + " """ + Format(value, "yyyy-mm-dd") + """"
            Else
                ret = ret + " " + key + " " + value
            End If
        End If
        
Continue:
    Next
    
    ret = ret + " }"
    
    LDTrf = ret
    Exit Function
    
Err:
    
    LDTrf = "#LD Error: Keys, ValueTypes and Values need to be 1-dimensional arrays of the same size"
    
End Function

