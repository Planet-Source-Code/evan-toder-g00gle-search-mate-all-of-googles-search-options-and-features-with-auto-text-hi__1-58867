VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGoogle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum enFilter
    [filter_none=0] = 0
    [moderate=1] = 1
    [strict=2] = 2
End Enum

Enum enCategory
   [web=0] = 0
   [images=1] = 1
   [groups=2] = 2
   [news=3] = 3
   [froogle=4] = 4
End Enum

Enum enResults
   [Ten=0] = 0
   [Twenty=1] = 1
   [Fifty=2] = 2
   [OneHundred=3] = 3
End Enum

Enum enWindow
    [sameWindow=0] = 0
    [newWindow=1] = 1
End Enum

Enum enFileType
    [filetype_all=0] = 0
    [AdobeAcrobat=1] = 1
    [AdobePostscript=2] = 2
    [microsoftWord=3] = 3
    [microsoftExcel=4] = 4
    [microsoftPowerpoint=5] = 5
    [richText=6] = 6
End Enum
 
'THE FINAL SEARCH STRING PUT TOGETHER AND READY FOR GOOGLE (WE HOPE)
Public Function GoogleSearchString(searchWords As String, _
                                   Optional wordsToExclude As String, _
                                   Optional searchInSite As String, _
                                   Optional searchCategory As enCategory, _
                                   Optional numResultsPerPage As enResults, _
                                   Optional windowToUse As enWindow, _
                                   Optional adultFilter As enFilter, _
                                   Optional fileRestriction As enFileType)
 
 'make sure a search word(s) were supplied
 If Len(Trim$(searchWords)) = 0 Then Exit Function
 
 Dim m_SearchWords      As String
 Dim m_wordsToExclude   As String
 Dim m_searchInSite     As String
 Dim m_searchCategory   As String
 Dim m_resultsPerPage   As String
 Dim m_windowToUse      As String
 Dim m_adultFilter      As String
 Dim m_fileRestriction  As String
 
 
 m_SearchWords = searchWords
 
 'words to explicitly exclude ?
 If Len(Trim$(wordsToExclude)) > 0 Then
     m_wordsToExclude = "&as_eq=" & wordsToExclude
 Else
     m_wordsToExclude = ""
 End If
 
 'search within specific site ?
 If Len(Trim$(searchInSite)) > 0 Then
    m_searchInSite = "&as_sitesearch=" & searchInSite
 Else
    m_searchInSite = ""
 End If
 
 'search category
 If searchCategory = [web=0] Then
      m_searchCategory = "search?q="
 ElseIf searchCategory = [images=1] Then
      m_searchCategory = "images?q="
 ElseIf searchCategory = [groups=2] Then
      m_searchCategory = "groups?q="
 ElseIf searchCategory = [news=3] Then
      m_searchCategory = "news?q="
 ElseIf searchCategory = [froogle=4] Then
      m_searchCategory = "froogle?q="
 End If
 
 'num results per page
 If numResultsPerPage = [Ten=0] Then
     m_resultsPerPage = "&num=10"
 ElseIf numResultsPerPage = [Twenty=1] Then
     m_resultsPerPage = "&num=20"
 ElseIf numResultsPerPage = [Fifty=2] Then
     m_resultsPerPage = "&num=50"
 ElseIf numResultsPerPage = [OneHundred=3] Then
     m_resultsPerPage = "&num=100"
 End If
 
 'window to use
 If windowToUse = [sameWindow=0] Then
     m_windowToUse = "&newwindow=0"
 ElseIf windowToUse = [newWindow=1] Then
     m_windowToUse = "&newwindow=1"
 End If
 
 'adult filter
 If adultFilter = [filter_none=0] Then
     m_adultFilter = "&safe=off"
 ElseIf adultFilter = [moderate=1] Then
     m_adultFilter = ""
 ElseIf adultFilter = [strict=2] Then
     m_adultFilter = "&safe=active"
 End If
 
 'file restriction
 If fileRestriction = [filetype_all=0] Then
     m_fileRestriction = ""
 ElseIf fileRestriction = [AdobeAcrobat=1] Then
     m_fileRestriction = "&as_filetype=pdf"
 ElseIf fileRestriction = [AdobePostscript=2] Then
     m_fileRestriction = "&as_filetype=ps"
 ElseIf fileRestriction = [microsoftWord=3] Then
     m_fileRestriction = "&as_filetype=doc"
 ElseIf fileRestriction = [microsoftExcel=4] Then
     m_fileRestriction = "&as_filetype=xls"
 ElseIf fileRestriction = [microsoftPowerpoint=5] Then
     m_fileRestriction = "&as_filetype=ppt"
 ElseIf fileRestriction = [richText=6] Then
     m_fileRestriction = "&as_filetype=rtf"
 End If
 
 
 'we can now put it all together and return the properly
 'formatted string which will be the url for the search
 GoogleSearchString = "http://www.google.com/" & _
                        m_searchCategory & _
                        m_SearchWords & _
                        m_resultsPerPage & _
                        "&hl=en" & _
                        m_windowToUse & _
                        m_wordsToExclude & _
                        m_fileRestriction & _
                        m_searchInSite & _
                        m_adultFilter

End Function
 

 
