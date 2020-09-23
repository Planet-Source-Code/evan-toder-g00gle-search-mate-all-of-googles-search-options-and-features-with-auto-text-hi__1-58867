VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniReadWrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


'the string path to the ini file where dealing with
Public strIniFilePath As String
'Section name of the ini file well be reading or writing to
Public strSection As String


'¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨
'write a value to the ini file (strIniFile)
'in the specified section (strSection)
'¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨
Public Sub WriteToINI(strKeyName As String, strKeyVal As String)
On Error GoTo ErrHandler

 Call WritePrivateProfileString(strSection, strKeyName, strKeyVal, strIniFilePath)
 
Exit Sub
ErrHandler:
    ERR.Source = ERR.Source & "." & VarType(Me) & ".ProcName"
    MsgBox ERR.Number & vbTab & ERR.Source & ERR.Description
    ERR.Clear
    Resume Next
End Sub

'¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨
'read values from and INI file an input to variables or controls settings
'syntax: backcolor = obj.ReadFromINI(strKeyName)
'¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨
Public Function ReadFromINI(strKeyName As String, DefaultStringVal As String) As String
On Error GoTo ErrHandler
  
  ReadFromINI = String(255, " ")
  Call GetPrivateProfileString(strSection, strKeyName, DefaultStringVal, ReadFromINI, 255, strIniFilePath)
  
Exit Function
ErrHandler:
    ERR.Source = ERR.Source & "." & VarType(Me) & ".ProcName"
    MsgBox ERR.Number & vbTab & ERR.Source & ERR.Description
    ERR.Clear
    Resume Next
End Function

