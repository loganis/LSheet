Attribute VB_Name = "LdashMetadata"

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

Public Function LDMetadata() As Variant
    Dim objHTTP As New MSXML2.ServerXMLHTTP
    URL = "https://www.googleapis.com/analytics/v3/metadata/ga/columns?pp=1"
    objHTTP.Open "GET", URL, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    Call objHTTP.setOption(2, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS)

    Set sc = CreateObject("ScriptControl")
    sc.Language = "JScript"
    'sc.AddCode "function encode(str) {return encodeURIComponent(str);}"
    'Dim encoded As String
    'encoded = sc.Run("encode", query)
    sendString = ""
    objHTTP.send (sendString)
    
    If objHTTP.statusText = "OK" Then
        response = objHTTP.responseText
        sc.Eval "var obj=(" & response & ")" 'evaluate the json response
    
        NumberOfItems = sc.Eval("obj.totalResults")
        
        Dim ret() As Variant
        ReDim ret(1 To NumberOfItems, 1 To 4)
        Dim numDim As Integer
        Dim numMet As Integer
        
        numDim = 0
        numMet = 0
        
        Dim stype As String
        
        sc.AddCode "function getItemType(i){return obj.items[i-1].attributes.type;}"
        sc.AddCode "function getItemId(i){return obj.items[i-1].id;}"
        sc.AddCode "function getItemDesc(i){return obj.items[i-1].attributes.description;}"
        
        For i = 1 To NumberOfItems
            stype = sc.Run("getItemType", i)
            If stype = "DIMENSION" Then
                numDim = numDim + 1
                ret(numDim, 1) = sc.Run("getItemId", i)
                ret(numDim, 2) = Left(sc.Run("getItemDesc", i), 255)
            ElseIf stype = "METRIC" Then
                numMet = numMet + 1
                ret(numMet, 3) = sc.Run("getItemId", i)
                ret(numMet, 4) = Left(sc.Run("getItemDesc", i), 255)
            Else: GoTo Err
            End If
        Next i
        
        LDMetadata = ret
        
    Else
        LDMetadata = "#LD Error: " & objHTTP.statusText & " " & objHTTP.responseText
        
    End If
    
    Exit Function
Err:
    LDMetadata = "#LD Error: Unknown type in GA Metadata"
    
End Function

