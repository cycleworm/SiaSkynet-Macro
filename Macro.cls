VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Function PrepareSelection() As String
    PrepareSelection = ""
    Dim myRange As Range
    Dim rowCount
    Dim columnCount
    
    Set myRange = Selection
    If TypeOf Selection Is Excel.Range Then
        Debug.Print Selection.Address(0, 0)
        For Each area In Selection.Areas
            Debug.Print area.Address(0, 0)
            rowCount = UBound(Selection.Value, 1)
            columnCount = UBound(Selection.Value, 2)
            PrepareSelection = RangeToString(myRange, rowCount, columnCount)
        Next
    
    Else
        Debug.Print TypeName(Selection)
    End If
End Function

Function RangeToString(ByVal myRange As Range, ByVal rowCount As Integer, ByVal columnCount As Integer) As String
    RangeToString = ""
    Dim count
    If Not myRange Is Nothing Then
        Dim myCell As Range
        count = 1
        For Each myCell In myRange
            If (count Mod columnCount) = 0 Then
                RangeToString = RangeToString & myCell.Value & vbNewLine
            Else
                RangeToString = RangeToString & myCell.Value & vbTab
            End If
            count = count + 1
        Next myCell
    End If
End Function

Function UploadToSkynet(ByVal uploadData As String) As String
    UploadToSkynet = ""

    sUrl = "https://siasky.net/skynet/skyfile"
    sFileName = "tempfile.txt"
    bAsync = False
    Const STR_BOUNDARY  As String = "3fbd04f5-b1ed-4060-99b9-fca7ff59c113"
    Dim nFile           As Integer
    Dim baBuffer()      As Byte
   
    '--- prepare body
    uploadData = "--" & STR_BOUNDARY & vbCrLf & _
        "Content-Disposition: form-data; name=""file""; filename=""" & Mid$(sFileName, InStrRev(sFileName, "\") + 1) & """" & vbCrLf & _
        "Content-Type: text/html" & vbCrLf & vbCrLf & _
        uploadData & vbCrLf & _
        "--" & STR_BOUNDARY & "--"
    '--- post

    With CreateObject("Microsoft.XMLHTTP")
        .Open "POST", sUrl, bAsync
        .SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & STR_BOUNDARY
        .Send uploadData
     
        If Not bAsync Then
            pvPostFile = .ResponseText
            Dim Json As Object
            Set Json = JsonConverter.ParseJson(.ResponseText)
            Cells(6, 5).Value = "https://siasky.net/" & Json("skylink")
        End If
    End With
End Function
 

Private Sub Load_Click()
    Dim hReq As Object
    Dim i As Long
    Dim var As Variant
    Dim ws As Worksheet
    Dim Skylink As String
      
    Set ws = Tabelle1
    
    Skylink = Cells(3, 5).Value
    
    
    Set hReq = CreateObject("MSXML2.XMLHTTP")
        With hReq
            .Open "GET", Skylink, False
            .Send
        End With
    
    textArray = Split(hReq.ResponseText, vbNewLine)
    
    For y = 0 To UBound(textArray)
        rowArray = Split(textArray(y), vbTab)
        For x = 0 To UBound(rowArray)
            Cells(y + 8, x + 1).Value = rowArray(x)
        Next
    Next
End Sub

Private Sub Upload_Click()
    Dim uploadData
    uploadData = PrepareSelection
    UploadToSkynet (uploadData)
End Sub
