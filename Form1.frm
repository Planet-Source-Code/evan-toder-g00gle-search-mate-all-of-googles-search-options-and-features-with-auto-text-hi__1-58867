VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   " "
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOnTop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   9225
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   21
      ToolTipText     =   "place this window on top of others"
      Top             =   6795
      Width           =   510
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   8280
      Top             =   4455
   End
   Begin VB.Frame Frame1 
      Height          =   3570
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   4245
      Begin VB.Frame Frame2 
         Caption         =   "search..."
         ForeColor       =   &H00808080&
         Height          =   510
         Left            =   90
         TabIndex        =   16
         Top             =   945
         Width           =   4110
         Begin VB.TextBox txtInSite 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   2385
            TabIndex        =   19
            Text            =   "www..com"
            ToolTipText     =   "double click to reset"
            Top             =   180
            Width           =   1680
         End
         Begin VB.OptionButton optSiteRestrict 
            Caption         =   "in site..."
            Height          =   240
            Index           =   1
            Left            =   1440
            TabIndex        =   18
            Top             =   225
            Width           =   915
         End
         Begin VB.OptionButton optSiteRestrict 
            Caption         =   "all the web"
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   17
            Top             =   225
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.ComboBox cboWindow 
         Height          =   315
         Left            =   1305
         TabIndex        =   14
         Text            =   "new window"
         Top             =   3015
         Width           =   1500
      End
      Begin VB.ComboBox cboFileType 
         Height          =   315
         Left            =   1305
         TabIndex        =   12
         Text            =   "any"
         Top             =   2655
         Width           =   1500
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         Left            =   1305
         TabIndex        =   10
         Text            =   "none"
         Top             =   2295
         Width           =   1140
      End
      Begin VB.ComboBox cboResults 
         Height          =   315
         Left            =   1305
         TabIndex        =   8
         Text            =   "10"
         Top             =   1935
         Width           =   780
      End
      Begin VB.ComboBox cboCategory 
         Height          =   315
         Left            =   1305
         TabIndex        =   6
         Text            =   "web search"
         Top             =   1575
         Width           =   1770
      End
      Begin VB.TextBox txtExclude 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   585
         Width           =   2220
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   285
         Left            =   2925
         TabIndex        =   3
         Top             =   180
         Width           =   1140
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   45
         TabIndex        =   2
         Top             =   180
         Width           =   2850
      End
      Begin VB.Label lblSelHiliteColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "hilight color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   3555
         TabIndex        =   20
         ToolTipText     =   "select highlighter color"
         Top             =   1845
         Width           =   555
      End
      Begin VB.Image imgHiliter 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   3645
         Picture         =   "Form1.frx":030A
         ToolTipText     =   "search words hilighter"
         Top             =   1575
         Width           =   270
      End
      Begin VB.Image imgTack 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   3825
         Picture         =   "Form1.frx":0894
         ToolTipText     =   "lock panel"
         Top             =   3150
         Width           =   270
      End
      Begin VB.Label Label6 
         Caption         =   "show results in"
         Height          =   240
         Left            =   225
         TabIndex        =   15
         Top             =   3060
         Width           =   1410
      End
      Begin VB.Label Label2 
         Caption         =   "filetype"
         Height          =   240
         Left            =   765
         TabIndex        =   13
         Top             =   2700
         Width           =   1410
      End
      Begin VB.Label Label5 
         Caption         =   "adult filter"
         Height          =   240
         Left            =   540
         TabIndex        =   11
         Top             =   2340
         Width           =   1410
      End
      Begin VB.Label Label3 
         Caption         =   "results per page"
         Height          =   240
         Left            =   135
         TabIndex        =   9
         Top             =   1980
         Width           =   1410
      End
      Begin VB.Label Label4 
         Caption         =   "search category"
         Height          =   240
         Left            =   135
         TabIndex        =   7
         Top             =   1620
         Width           =   1410
      End
      Begin VB.Label Label1 
         Caption         =   "results must not include"
         Height          =   240
         Left            =   90
         TabIndex        =   5
         Top             =   630
         Width           =   1770
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   9645
      ExtentX         =   17013
      ExtentY         =   12726
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, ByRef lpPoint As Pointapi) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As Pointapi) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal Flags As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
 
Private Type Pointapi
   x As Long
   y As Long
End Type

Private Type variables
     mouse_over_frame   As Boolean
     hilite_color       As Long
End Type
Dim v               As variables
 
 
Private cINI        As clsIniReadWrite
Private cGoogle     As cGoogle
Private cClrDialog  As clsColorDialog
Private cHilite     As cHilighter

'--------------------------------------------------
' fill the combo boxes and do class initialization
'--------------------------------------------------
Private Sub Form_Load()

With cboCategory
   .AddItem "web search"
   .AddItem "image search"
   .AddItem "groups search"
   .AddItem "news search"
   .AddItem "froogle search (shopping)"
   'web search default
   .ListIndex = 0
End With

With cboFileType
   .AddItem "any"
   .AddItem "Adobe Acrobat"
   .AddItem "Adobe Postscript"
   .AddItem "Microsoft Word"
   .AddItem "Microsoft Excel"
   .AddItem "Microsoft Powerpoint"
   .AddItem "Rich Text"
   'default any/all filetypes
   .ListIndex = 0
End With

With cboResults
    .AddItem "10"
    .AddItem "20"
    .AddItem "50"
    .AddItem "100"
    'default 50 results per page
    .ListIndex = 2
End With

With cboFilter
    .AddItem "none"
    .AddItem "moderate"
    .AddItem "strict"
    'default no filtering
    .ListIndex = 0
End With

With cboWindow
   .AddItem "same window"
   .AddItem "new window"
   'default show results in new window
   .ListIndex = 1
End With


Set cGoogle = New cGoogle
Set cINI = New clsIniReadWrite
 
With cINI
    .strIniFilePath = App.Path & "\settings"
    .strSection = "settings"
    'retrieve hilighter color, yellow default
    v.hilite_color = CLng(.ReadFromINI("hilite", RGB(255, 255, 170)))
    lblSelHiliteColor.ForeColor = v.hilite_color
End With

Set cINI = Nothing

End Sub

'--------------------------------------------------
'  resize browser to fill form
'--------------------------------------------------
Private Sub Form_Resize()
 
 Dim titleheight  As Long
 Dim titlewidth   As Long
 
 titleheight = (Height - ScaleHeight) + 20
 titlewidth = (Width - ScaleWidth)
 WebBrowser1.Move 0, 0, (Width - 80), (Height - titleheight)
 
 With picOnTop
    .Move (Width - (.Width + titlewidth)), _
           (Height - (titleheight + .Height))
 End With
 
End Sub
'--------------------------------------------------
'  save settings and kill object reference
'--------------------------------------------------
Private Sub Form_Terminate()

  Set cGoogle = Nothing
  Set cClrDialog = Nothing
  Set cINI = New clsIniReadWrite
 
  With cINI
     .strIniFilePath = App.Path & "\settings"
     .strSection = "settings"
     'retrieve hilighter color, yellow default
     .WriteToINI "hilite", CStr(v.hilite_color)
  End With

  Set cINI = Nothing

End Sub
'--------------------------------------------------
'   the class returns the string properly formated
'   for the google search engin
'--------------------------------------------------
Private Sub cmdSearch_Click()
 
  If Len(Trim$(txtSearch)) > 0 Then
  
     Dim insite As String, endString As String
  
     If optSiteRestrict(1).Value = True Then
        insite = txtInSite
     Else
        insite = ""
     End If
      
     endString = cGoogle.GoogleSearchString(txtSearch _
                                     , txtExclude, insite _
                                     , cboCategory.ListIndex _
                                     , cboResults.ListIndex _
                                     , cboWindow.ListIndex _
                                     , cboFilter.ListIndex _
                                     , cboFileType.ListIndex)
      Debug.Print endString
      WebBrowser1.Navigate endString
   End If
  
End Sub
 
'--------------------------------------------------
' starts timer1 which keeps track of if
' frame1 should be opened or closed
'--------------------------------------------------
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  'if we want the panel locked then stop timer1
  'and lock the panel (frame1) .left=0: exit sub
  If imgTack.Appearance = 1 Then
     If Timer1 = True Then
        Timer1 = False
        Frame1.Left = 0
        v.mouse_over_frame = False
     End If
     Exit Sub
  End If
  
  If v.mouse_over_frame = False Then
     Timer1.Interval = 20
     Timer1.Enabled = True
     v.mouse_over_frame = True
  End If
  
End Sub
'--------------------------------------------------
'  user pressing hilighter button(image) so create
'  visual effect of it being pushed
'--------------------------------------------------
Private Sub imgHiliter_Click()

  imgHiliter.Appearance = IIf(imgHiliter.Appearance = 0, 1, 0)
  Call WebBrowser1_DocumentComplete(Nothing, "")
  
End Sub
'--------------------------------------------------
'  user pressing tac button(image) so create
'  visual effect of it being pushed
'--------------------------------------------------
Private Sub Imgtack_Click()
 
  imgTack.Appearance = IIf(imgTack.Appearance = 0, 1, 0)

End Sub
 
'--------------------------------------------------
' user selecting a new hilighter color
'--------------------------------------------------
Private Sub lblSelHiliteColor_Click()

  Dim tacAppear    As Long
  
  'store the up or down appearance of the img(imgTack)
  tacAppear = imgTack.Appearance
  '"press" the tack down so panel stays open
  'while we are selecting new hilighter color
  imgTack.Appearance = 1
  
  Set cClrDialog = New clsColorDialog
  With cClrDialog
      .hWndParent = hwnd
      .Flags = colPreventFullOpen
      v.hilite_color = .ShowColor(v.hilite_color)
      lblSelHiliteColor.ForeColor = v.hilite_color
  End With
  Set cClrDialog = Nothing
  
  'reset imgTack back
  imgTack.Appearance = tacAppear
  
End Sub
'--------------------------------------------------
' user selecting whether to restrict search to the
' web or a specific site.  enable or disable the
' associated textbox as needed
'--------------------------------------------------
Private Sub optSiteRestrict_Click(Index As Integer)

  txtInSite.Enabled = IIf(optSiteRestrict(1).Value = True, True, False)
  
End Sub
    
'--------------------------------------------------
' place or remove this form from top of zorder
'--------------------------------------------------
Private Sub picOnTop_Click()
 
 Dim topval As Long
 
 With picOnTop
    If .Appearance = 0 Then
       .Appearance = 1
       .ToolTipText = "remove window from top"
       topval = -1
    Else
       .Appearance = 0
       .ToolTipText = "place this window on top"
       topval = 1
    End If
 End With
 
 SetWindowPos hwnd, topval, 0, 0, 0, 0, &H1 Or &H2
 BringWindowToTop hwnd
 
End Sub

'--------------------------------------------------
'  a method for determining that the mouse is
'  within a rect area (frame1 in this case)
'  without having to worry about window under
'  mouse and parent of the window under mouse
'--------------------------------------------------
Private Sub Timer1_Timer()
 
  Dim pt          As Pointapi
  Dim hwndFromPt  As Long
  Dim framePixWid As Long
  Dim framePixHei As Long
  
  If imgTack.Appearance = 1 Then Exit Sub
  
  framePixWid = (Frame1.Width / Screen.TwipsPerPixelX)
  framePixHei = (Frame1.Height / Screen.TwipsPerPixelY)
  GetCursorPos pt
  ScreenToClient Frame1.hwnd, pt
  
  'if the cursor lies outside of the frames rect
  'then scroll closed the frame
  If (pt.x < 0) Or _
     (pt.x > framePixWid) Or _
     (pt.y < 0) Or _
     (pt.y > framePixHei) Then
 
     If v.mouse_over_frame = True Then
        v.mouse_over_frame = False
     End If
     
     If Frame1.Left > -(Frame1.Width - 250) Then
        Frame1.Left = (Frame1.Left - 120)
     Else
        Timer1.Interval = 0
     End If
  Else
     If Frame1.Left < 0 Then
        Frame1.Left = (Frame1.Left + 120)
     End If
  End If
 
End Sub
'--------------------------------------------------
' reset the textbox on dble click to have www..com
' and place the cursor between the 2 dots
'--------------------------------------------------
Private Sub txtInSite_DblClick()
  
   txtInSite = "www..com"
   txtInSite.SelStart = 4
  
End Sub
'--------------------------------------------------
' the search page has finished loading so hilight
' search text if imghiliter is down
'--------------------------------------------------
Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
  
  Dim nonsense_str   As String
  
  nonsense_str = Trim(WebBrowser1.LocationURL)
  If nonsense_str = "http:///" Or _
     nonsense_str = "about:blank" Or _
     nonsense_str = "" Then Exit Sub
 
  If imgHiliter.Appearance = 1 Then
     Set cHilite = New cHilighter
     Caption = cHilite.hilite_marker_text(hwnd, WebBrowser1.Document, _
                                       Trim(txtSearch), v.hilite_color) & _
                                " occurrences of " & Chr(34) & _
                                txtSearch & Chr(34) & " found"
     Set cHilite = Nothing
  End If
  
End Sub
'--------------------------------------------------
' new window opening for search results, hopefully
'--------------------------------------------------
Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)

 Dim frmRes  As New frmResults
 Set ppDisp = frmRes.WebBrowser1.Application
 frmRes.Show vbModeless, Me

End Sub
