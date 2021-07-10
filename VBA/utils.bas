Attribute VB_Name = "utils"
Option Explicit

' File sytem object
Public Function initFSO() As Object
 Static oFSO As Object
 If oFSO Is Nothing Then Set oFSO = CreateObject("Scripting.FileSystemObject")
 Set initFSO = oFSO
End Function

' Shell object
Public Function initAppShell() As Object
 Static oAppShell As Object
 If oAppShell Is Nothing Then Set oAppShell = CreateObject("Shell.Application")
 Set initAppShell = oAppShell
End Function

Function readFile(iv_file As String) As String
    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile (iv_file)
    
    readFile = objStream.ReadText()
    
    objStream.Close
End Function

Public Sub writeFile(iv_file As String, iv_text As String)
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    fsT.Type = 2 'Specify stream type - we want To save text/string data.
    fsT.Charset = "utf-8" 'Specify charset For the source text data.
    fsT.Open 'Open the stream And write binary data To the object
    fsT.WriteText iv_text
    fsT.SaveToFile iv_file, 2 'Save binary data To disk
    fsT.Close
End Sub

Public Sub writeJSON(iv_file As String, sJSON As String)
 sJSON = Left(sJSON, Len(sJSON) - 1) & "]"
 
 ' No need in saving
 If Len(sJSON) = 1 Then Exit Sub
 
 Call writeFile(iv_file, sJSON)
End Sub
