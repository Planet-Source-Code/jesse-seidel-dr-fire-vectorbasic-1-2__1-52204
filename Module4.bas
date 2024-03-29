Attribute VB_Name = "Module1"
Public Function ReadFile(strPath As String) As Variant
    On Error GoTo eHandler
    Dim iFileNumber As Integer
    Dim blnOpen As Boolean
    iFileNumber = FreeFile
    Open strPath For Input As #iFileNumber
    blnOpen = True
    ReadFile = Input(LOF(iFileNumber), iFileNumber)
eHandler:
    If blnOpen Then Close #iFileNumber
    If Err Then MsgBox Err.Description, vbOKOnly + vbExclamation, Err.Number & " - " & Err.Source
End Function


Public Function WriteFile(strPath As String, strValue As String) As Boolean
    On Error GoTo eHandler
    Dim iFileNumber As Integer
    Dim blnOpen As Boolean
    iFileNumber = FreeFile
    Open strPath For Output As #iFileNumber
    blnOpen = True
    Print #iFileNumber, strValue
eHandler:
    If blnOpen Then Close #iFileNumber


    If Err Then
        MsgBox Err.Description, vbOKOnly + vbExclamation, Err.Number & " - " & Err.Source
    Else
        WriteFile = True
    End If
End Function


