VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FolderInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public path As String
Public files As Collection

Public Function getBaseName() As String
 getBaseName = g_FSO.getBaseName(Me.path)
End Function

Public Function getVisibleFilesCount() As Long
 Dim oFileInfo As FileInfo
 
 For Each oFileInfo In Me.files
  If oFileInfo.IsHidden() Then GoTo lbContinue
  
  getVisibleFilesCount = getVisibleFilesCount + 1
lbContinue:
 Next
End Function

Public Function findIcon(sPrefix As String) As String
 Dim oFileInfo As FileInfo
 
 For Each oFileInfo In Me.files
  If oFileInfo.f_name Like sPrefix & ".*" Then
  
   ' Ok 1-st item
   findIcon = oFileInfo.get_json
   Exit Function
  End If
 Next
End Function

Public Sub make_json(fileName As String)
 Dim sJSON As String, oFileInfo As FileInfo
 
 sJSON = "["
 For Each oFileInfo In Me.files
  If oFileInfo.IsHidden() Then GoTo lbContinue
  
  sJSON = sJSON & oFileInfo.get_json() & ","
lbContinue:
 Next
 
 Call writeJSON(Me.path & fileName, sJSON)
End Sub
