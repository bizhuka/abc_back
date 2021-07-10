VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public f_name As String
Public f_relative_path As String
Public f_size As Long
Public f_date As String

Public Function getBaseName() As String
 getBaseName = g_FSO.getBaseName(Me.f_name)
End Function

Public Function getExt() As String
 getExt = "." & g_FSO.GetExtensionName(Me.f_name)
End Function

Public Function IsHidden() As Boolean
 IsHidden = Left(Me.f_name, 1) = "_"
End Function

Public Function get_json() As String
 get_json = "{" & _
  """name"":""" & Me.f_relative_path & """," & _
  """size"":" & Me.f_size & "," & _
  """date"":""" & Me.f_date & """" & _
 "}" & vbCrLf
End Function
