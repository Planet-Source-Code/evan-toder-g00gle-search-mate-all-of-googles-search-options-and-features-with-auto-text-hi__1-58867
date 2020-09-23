VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHilighter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long



 

'--------------------------------------------------
' returns the number of instances the highlighted
' word was found
'--------------------------------------------------
Function hilite_marker_text(yourFormsHwnd As Long, _
                   yourDocument As HTMLDocument, _
                   str_text As String, _
                   HighlightColor) As Long
                   
   'dont process if text isnt supplied
   If Len(Trim$(str_text)) = 0 Then Exit Function
   
   Dim rngCnt         As Long
   Dim oBODY          As HTMLBody
   Dim Range          As IHTMLTxtRange
   Dim bfound         As Boolean
   
   If yourDocument Is Nothing Then Exit Function
   
   'make sure reference is set to body
   Set oBODY = yourDocument.body
   
  'start the range to encompass ALL the pages text
   Set Range = oBODY.createTextRange
   
   'speed up the process
   LockWindowUpdate yourFormsHwnd
   
   Do 'tell the  range object to find the text (str_to_find (txtFind))
     bfound = Range.findText(str_text)
     
     'If its found, select it, change its backcolor
     If bfound Then
        On Error Resume Next
        Range.Select
        yourDocument.execCommand "backcolor", False, HighlightColor
       'this tells the orange object to resume the search with
       'the start point being the end of the word just found
        Range.collapse False
        rngCnt = (rngCnt + 1)
     End If
 
     DoEvents
      'keep going til we dont find the word(s) anymore
    Loop Until Not (bfound)
    
    'return number of matches found
    hilite_marker_text = rngCnt
    
    'reenable update painting
    LockWindowUpdate 0&
    
    'scroll the page back to the top
    yourDocument.parentWindow.Scroll 0, 0
    Set Range = Nothing
    
End Function
