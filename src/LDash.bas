Attribute VB_Name = "LDash"

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

Public Function LDQuery(token As String, query As String) As Variant
    Dim objHTTP As New MSXML2.ServerXMLHTTP
    URL = "https://ldash.loganis.com/LQS.aspx/api/v3"
    objHTTP.Open "POST", URL, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    Call objHTTP.setOption(2, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS)

    Set sc = CreateObject("ScriptControl")
    sc.Language = "JScript"
    sc.AddCode "function encode(str) {return encodeURIComponent(str);}"
    Dim encoded As String
    encoded = sc.Run("encode", query)
    sendString = "token=" & token & "&query=" & encoded
    objHTTP.send (sendString)
    
    If objHTTP.statusText = "OK" Then
        response = objHTTP.responseText
        sc.Eval "var obj=(" & response & ")" 'evaluate the json response
        'add some accessor functions
        sc.AddCode "function getNumberOfColumns(){return obj.colnames.length;}"
        sc.AddCode "function getNumberOfRows(){return obj.rows.length;}"
        sc.AddCode "function getColNames(i){return obj.colnames[i];}"
        sc.AddCode "function getElement(i,j){return obj.rows[i][j];}"
    
        NumberOfColumns = sc.Run("getNumberOfColumns")
        NumberOfRows = sc.Run("getNumberOfRows")
        
        Dim ret() As Variant
        ReDim ret(NumberOfRows, NumberOfColumns - 1)
        
        For i = 0 To NumberOfRows
        For j = 0 To NumberOfColumns - 1
            If i <> 0 Then
                ret(i, j) = sc.Run("getElement", i - 1, j)
            Else
                ret(0, j) = sc.Run("getColNames", j)
            End If
        Next j
        Next i
        
        LDQuery = ret
        
    Else
        LDQuery = "#LD Error: " & objHTTP.statusText & objHTTP.responseText
        
    End If
    
End Function
