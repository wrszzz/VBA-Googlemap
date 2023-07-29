Attribute VB_Name = "ģ��1"
Function TravelTime(origin As String, destination As String) As String
    'API Key - replace with your own key
    apiKey = "YOUR API KEY"
    
    'Google Maps API url
    apiUrl = "https://maps.googleapis.com/maps/api/distancematrix/xml?origins=" & origin & "&destinations=" & destination & "&mode=transit&key=" & apiKey

    Dim myRequest As Object
    Set myRequest = CreateObject("MSXML2.ServerXMLHTTP")
    
    myRequest.Open "GET", apiUrl, False
    myRequest.send
    
    Dim myDomDoc As Object
    Set myDomDoc = CreateObject("MSXML2.DOMDocument")
    myDomDoc.LoadXML myRequest.responseText
    
    Dim myNode As Object
    
    ' Extracting the travel time from the XML response
    Set myNode = myDomDoc.SelectSingleNode("//duration/text")
    '' Extracting the distance from the XML response
    ' Set myNode = myDomDoc.SelectSingleNode("//distance/text")
    

    If Not myNode Is Nothing Then
        TravelTime = myNode.Text
    ' If Not myNode Is Nothing Then
    '     TravelDistance = myNode.Text    
    Else
        TravelTime = "Error"
        ' TravelDistance = "Error"
    End If
End Function

