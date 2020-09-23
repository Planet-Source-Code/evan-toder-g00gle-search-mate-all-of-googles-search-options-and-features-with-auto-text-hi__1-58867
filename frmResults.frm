VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmResults 
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form2"
   ScaleHeight     =   5775
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5730
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   10680
      ExtentX         =   18838
      ExtentY         =   10107
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private cHilight    As cHilighter


Private Sub Form_Resize()

 WebBrowser1.Move 0, 0, (Width - 60), Height - (Height - ScaleHeight)

End Sub

 
'--------------------------------------------------
' highlight the words on this results page
'--------------------------------------------------
Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
  
  Dim nonsense_str   As String
  
  nonsense_str = Trim(WebBrowser1.LocationURL)
  If nonsense_str = "http:///" Or _
     nonsense_str = "about:blank" Or _
     nonsense_str = "" Then Exit Sub
     
  Set cHilight = New cHilighter
  cHilight.hilite_marker_text hwnd, WebBrowser1.Document _
                     , Form1.txtSearch, _
                     Form1.lblSelHiliteColor.ForeColor
  Set cHilight = Nothing
 
End Sub

'--------------------------------------------------
' we dont want any popups here, or do we?
'--------------------------------------------------
Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
  
Dim iYesNo      As Integer
Dim sText       As String
Dim sTitle      As String
    
    sText = Chr(34) & WebBrowser1.LocationName & Chr(34) & _
                      " is attempting to " & vbCrLf & _
                      "open a new browser window." & vbCrLf & _
                      "Do you wish to allow it ?"
    iYesNo = MsgBox(sText, vbYesNo)
    'If user selects yes then
    If iYesNo = vbYes Then
       Cancel = False
    Else
       Cancel = True
    End If
    
End Sub

 
