Attribute VB_Name = "main"
Option Explicit

Public g_FSO As Object
Public g_AppShell As Object
Public g_now As String
Private g_forWeb As Boolean
Private g_root As String


Private Sub init(forWeb As Boolean, root As String)
 Set g_FSO = initFSO()
 Set g_AppShell = initAppShell()
 g_now = Format(Now(), "yyyy-mm-dd hh:mm:ss")
 
 g_forWeb = forWeb
 g_root = root
End Sub

Private Sub sub_make_json()
 ' As constructor
 Call init(True, shMain.Range("ROOT_FOLDER").Value)
 
 Dim folderInfos As Collection, oFolderInfo As FolderInfo
 Call processFolder(g_root, folderInfos)
 
 For Each oFolderInfo In folderInfos
  Call oFolderInfo.make_json("_all.json")
 Next
 
 Call makeAbcMain(folderInfos)
 
 MsgBox "Done"
End Sub

Private Sub makeAbcMain(folderInfos As Collection)
 Dim rAbcMain As Range, vAbcMain As Variant, i As Long
 
 ' Main table
 Set rAbcMain = shMain.ListObjects("abcMain.json").DataBodyRange
 vAbcMain = rAbcMain.Value
 
 Dim oAbcMain As AbcMain, sJSON As String
 
 sJSON = "["
 For i = LBound(vAbcMain) To UBound(vAbcMain)
  Set oAbcMain = New AbcMain
  If oAbcMain.init_abc(vAbcMain, i, folderInfos) Then
    sJSON = sJSON & oAbcMain.get_json() & ","
  End If
 Next
 
 Call writeJSON(g_root & "cards/abcMain.json", sJSON)
 
 ' Show info
 rAbcMain.Value = vAbcMain
End Sub

''' Private Sub find_icons()
'''
'''  ' As constructor
'''  Call init(False, shMain.Range("ROOT_FOLDER").Value)
'''
'''  Dim folderInfos As Collection
'''  Call processFolder(g_root, folderInfos)
'''
'''  ' Rename icons
'''  Dim oIconFolder As FolderInfo
'''  Set oIconFolder = folderInfos("_icon\")
'''
'''   ' Copy icons object
'''  Dim oFileInfo As FileInfo, vFrom, vTo
'''  For Each oFileInfo In oIconFolder.files
'''   ' Current icon file
'''   Debug.Print "File " & oFileInfo.f_name
'''
'''   vFrom = g_root & oFileInfo.f_relative_path
'''   vTo = g_root & "cards\" & LCase(oFileInfo.getBaseName() & "\_icon" & oFileInfo.getExt())
'''
'''   ' And copy
'''   Call g_FSO.CopyFile(vFrom, vTo, True)
'''  Next
'''
''' MsgBox "Done"
'''End Sub

' Обработка папки
Private Sub processFolder(ByVal sPath As String, ByRef coll As Collection)
 Dim oFolderInfo As FolderInfo, oShell As Object, sFile As String, i As Long, sFullPath As String
 Dim collSubFolders As Collection
 
 ' Init 1 time
 If coll Is Nothing Then Set coll = New Collection
 
 ' current folder
 Set oFolderInfo = New FolderInfo
 oFolderInfo.path = sPath
 Set oFolderInfo.files = New Collection
 
 ' Addby key
 coll.Add Key:=LCase(get_rel_path(sPath)), Item:=oFolderInfo
 
 ' process sub folders
 Set collSubFolders = New Collection

 sFile = Dir(sPath & "*.*", vbNormal + vbDirectory)
 Do While sFile <> ""
  If sFile = "." Or sFile = ".." Then GoTo lbContinue
  
  sFullPath = sPath & sFile
  
  ' Check also subfolder
  If g_FSO.FolderExists(sFullPath) Then
   collSubFolders.Add sFullPath & "\"
  Else
   Dim oFile As Object, oFileInfo As FileInfo
   Set oFile = g_FSO.GetFile(sFullPath)
   
   Set oFileInfo = New FileInfo
   oFileInfo.f_relative_path = get_rel_path(sFullPath)
   oFileInfo.f_name = oFile.name
   oFileInfo.f_size = oFile.size
   oFileInfo.f_date = Format(oFile.DateLastModified, "yyyymmddhhmmss")
   oFolderInfo.files.Add oFileInfo
  End If
 
lbContinue:
   sFile = Dir()
 Loop
 
 ' Check also subfolders
 Dim vFolder As Variant
 For Each vFolder In collSubFolders
  processFolder vFolder, coll
 Next
End Sub

Private Function get_rel_path(sPath As String) As String
 get_rel_path = Replace(sPath, g_root, "")
 If g_forWeb Then get_rel_path = Replace(get_rel_path, "\", "/")
End Function
