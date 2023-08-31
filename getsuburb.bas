Attribute VB_Name = "ģ��2"
Option Explicit

Function GetSuburbName(address As String) As String
    ' API Key - replace with your own key
    Dim apiKey As String
    apiKey = "your api" ' Consider storing this securely
    
    ' Encode the address for use in the URL
    Dim encodedAddress As String
    encodedAddress = URLEncode(address)
    
    ' Google Maps Geocoding API url
    Dim apiUrl As String
    apiUrl = "https://maps.googleapis.com/maps/api/geocode/xml?address=" & encodedAddress & "&key=" & apiKey

    Dim myRequest As Object
    Set myRequest = CreateObject("MSXML2.ServerXMLHTTP")
    
    On Error GoTo ErrorHandler
    myRequest.Open "GET", apiUrl, False
    myRequest.send
    
    Debug.Print myRequest.responseText

    
    Dim myDomDoc As Object
    Set myDomDoc = CreateObject("MSXML2.DOMDocument")
    myDomDoc.LoadXML myRequest.responseText
    
    Dim suburbNode As Object
    Set suburbNode = myDomDoc.SelectSingleNode("//address_component[type='locality']/long_name")

    
        
    If Not suburbNode Is Nothing Then
        GetSuburbName = suburbNode.Text
    Else
        GetSuburbName = "Error"
    End If

    Exit Function
ErrorHandler:
    MsgBox "An error occurred. " & vbCrLf & "Error number: " & Err.Number & vbCrLf & "Error description: " & Err.Description
    GetSuburbName = "Error"
    
End Function

Function URLEncode(s As String) As String
    Dim i As Integer
    Dim CharCode As Integer
    Dim result As String

    For i = 1 To Len(s)
        CharCode = Asc(Mid(s, i, 1))
        If (CharCode >= 48 And CharCode <= 57) Or (CharCode >= 65 And CharCode <= 90) Or (CharCode >= 97 And CharCode <= 122) Then
            result = result & Chr(CharCode)
        Else
            result = result & "%" & Right("0" & Hex(CharCode), 2)
        End If
    Next i

    URLEncode = result
End Function

