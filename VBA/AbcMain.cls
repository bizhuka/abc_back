VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AbcMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private db As String
Private name As String
Private desc As String
Private auth As String
Private cnt As Long
Private icon As String

Public Function init_abc(vAbcMain As Variant, i As Long, folderInfos As Collection) As Boolean
 Dim sPrefix As String, sDbKey As String
 sPrefix = vAbcMain(i, 1)
 sDbKey = LCase(sPrefix)
 
 name = vAbcMain(i, 2)
 desc = vAbcMain(i, 3)
 auth = vAbcMain(i, 4)
 
 Dim sLog As String
 Dim oCardsFolder As FolderInfo, oIconsFolder As FolderInfo
 ' Try block
 On Error Resume Next
  Set oCardsFolder = folderInfos("cards/" & sDbKey & "/")
  Set oIconsFolder = folderInfos("icons/")
  If Err.Number <> 0 Then sLog = "ERROR: No folder"
 On Error GoTo 0
 
 If sLog = "" Then
  ' Case sansetive name !
  db = oCardsFolder.getBaseName()
 
  cnt = oCardsFolder.getVisibleFilesCount()
  
  ' _icon with ext jpg png gif
  icon = oIconsFolder.findIcon(sPrefix)
  If icon = "" Then sLog = "ERROR: No icon"
 End If
 
 ' Ok
 If sLog = "" Then
  sLog = g_now & " cnt=" & cnt
  init_abc = True
 End If
 
 ' Set info
 vAbcMain(i, 5) = sLog
End Function

Public Function get_json() As String
 ' 3 types of files!
 get_json = "{" & _
  """db"":""" & db & """," & _
  """name"":""" & name & """," & _
  """desc"":""" & desc & """," & _
  """auth"":""" & auth & """," & _
  """cnt"":" & cnt / 3 & "," & _
  """icon"":" & icon & _
 "}"
End Function
