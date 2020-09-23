VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
'##############################
'make this is the private  class
'##############################

Private Type CHOOSECOLOR
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As Long
  Flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Public Enum EnumColFlags
  colEnableHook = &H10
  colEnableTemplate = &H20
  colEnableTemplateHandle = &H40
  colFullOpen = &H2
  colPreventFullOpen = &H4
  colRgbinit = &H1
  colShowHelp = &H8
End Enum

Private Declare Function ChooseColor_API Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long

Private m_eFlags As EnumColFlags
Private m_lnghWndParent As Long

Private Sub Class_Initialize()
  m_eFlags = colFullOpen Or colRgbinit
End Sub

Public Property Get Flags() As EnumColFlags
  Flags = m_eFlags
End Property
Public Property Let Flags(ByVal eValue As EnumColFlags)
  m_eFlags = eValue
End Property

Public Property Get hWndParent() As Long
  hWndParent = m_lnghWndParent
End Property
Public Property Let hWndParent(ByVal lngValue As Long)
  m_lnghWndParent = lngValue
End Property
'------------------------
'function return = long
'(color selected)
'orig color is the color of the
'object before selecting new color
'so if cancel is selected or there
'is an error then the color
'isnt altered
'------------------------
Public Function ShowColor(OrigColor As Long) As Long
 On Error GoTo ERR_HANDLER:
'-----VARIABLES----------
Dim cc As CHOOSECOLOR
Dim alngColorRef(15) As Long
'-----START CODE--------
  With cc ' Initialize the CHOOSECOLOR type
     .lStructSize = LenB(cc)
     .hwndOwner = m_lnghWndParent
     .lpCustColors = VarPtr(alngColorRef(0))
     .rgbResult = ShowColor
     .Flags = m_eFlags
  End With
  
  If ChooseColor_API(cc) <> 0 Then ' Display the dialog
     ShowColor = cc.rgbResult
  Else
     ShowColor = OrigColor
  End If
'-----END CODE---------
Exit Function
ERR_HANDLER:
  Debug.Print "sub Show " & ERR.Number; ERR.Description
End Function

