Option Explicit

#Const Testing = False
Private m_acc As AccessDb

Public Sub TestEnglandSearchByLocation()
  Dim h As HttpHelper
  Set h = New HttpHelper
  Dim html As String
  Dim file As String
  
  html = h.PostForm("https://find-postgraduate-teacher-training.education.gov.uk/?l=2&qualifications=QtsOnly,PgdePgceWithQts,Other&fulltime=False&parttime=False&hasvacancies=True&senCourses=False", Nothing)
  file = ThisWorkbook.path & "\SearchResponse.html"
  WriteFile file, html
  MsgBox "Saved response to " & file
End Sub

Public Sub UCASSearch()
  
  Dim u As New UCASQUeryParameters
  u.SetValue Age, "P"
  u.SetValue Regions, "02"
  
  Dim uSearch As UCASSearch
  Set uSearch = New UCASSearch
  
  uSearch.Initialise "http://search.gttr.ac.uk/cgi-bin/hsrun.hse/General/2017_gttr_search"
  Dim res As String
  Call uSearch.Search(u)
  Dim f As frmBrowser
  Set f = New frmBrowser
  f.Show vbModeless
  f.ShowHtml uSearch.HtmlDoc.ToString
  
End Sub

Public Sub TestFileParsing()
  Dim uSrch As UCASSearch
  Dim u As New UCASQUeryParameters
  
  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
  
  Set uSrch = New UCASSearch
  
  'Dim f As frmBrowser
  'Set f = New frmBrowser
  'f.Show vbModeless

  'u.SetValue Age, "P"
  u.SetValue Regions, "02"
  
  uSrch.LogFile = "UCAS Search.log"
  
  uSrch.Initialise "http://search.gttr.ac.uk/cgi-bin/hsrun.hse/General/2018_gttr_search" ' , f
  
  uSrch.SearchParameters = u
  #If Testing Then
    uSrch.Load ThisWorkbook.path & "\Site\South East 2017.html"
  #Else
    uSrch.Search u
  #End If
   
  uSrch.LoadUCASTeacherTraining
  
  Application.ScreenUpdating = True
End Sub

Private Sub ShowProgress()
  On Error GoTo ERR_CATCH
  Dim f As frmProgress
  Set f = New frmProgress
  f.Total = 100
  
  Application.ScreenUpdating = True
  
  f.Show vbModeless
  Dim n As Integer
  f.Caption = "Performing Search"
  Sleep 1000
  DoEvents
  f.Caption = "Loading South East"
  For n = 1 To 100
    f.Header = "Subject " & n
    f.Info = "School " & n
    f.value = n
    Sleep 1000
    DoEvents
  Next n
ERR_CATCH:
  On Error Resume Next
  If Not f Is Nothing Then f.Hide
  Set f = Nothing
End Sub

Private Sub TestDataTableRefresh()
  UpsertDataTable
End Sub

Private Sub TestCsvImport()
  Dim file  As String
  Dim csv   As New CSVReader
  Dim dbg   As New FDebug
  Dim URN   As String, pc As String
  Dim sql   As String, sqlTemplate As String
  
  Dim oDb As AccessDb
  Set oDb = New AccessDb
  oDb.Initialise
  
  oDb.ExecuteSQLNoResults "DELETE * FROM Urn;"
  
  sqlTemplate = "INSERT INTO URN( URN, URNPostalCode ) VALUES('{0}','{1}');"
  
  file = ThisWorkbook.path & "\edubasealldata20151115.csv"
  csv.OpenFile file, True
  dbg.TimerStart
  While csv.ReadLine()
    sql = StrFormat(sqlTemplate, csv.Field("URN"), csv.Field("PostCode"))
    If csv.row Mod 1000 = 0 Then
      Debug.Print csv.row
    End If
    oDb.ExecuteSQLNoResults sql
  Wend
  dbg.TimerEnd
  
End Sub

Private Sub TestAddressMatching()
  Dim tl As TrainingLocation
  
  Set m_acc = New AccessDb
  
  tl.Initialise m_acc
  tl.MatchURN
  
End Sub



Private Sub LTest()
  Debug.Print Levenshtein("Oxford University", "of Oxford University")
  Debug.Print Levenshtein("George Abbot SCITT", "George Abbot School")
End Sub

Function Levenshtein(ByVal string1 As String, ByVal string2 As String) As Long

  Dim i As Long, j As Long, bs1() As Byte, bs2() As Byte
  Dim string1_length As Long
  Dim string2_length As Long
  Dim distance() As Long
  Dim min1 As Long, min2 As Long, min3 As Long

  string1_length = Len(string1)
  string2_length = Len(string2)
  ReDim distance(string1_length, string2_length)
  bs1 = string1
  bs2 = string2

  For i = 0 To string1_length
    distance(i, 0) = i
  Next

  For j = 0 To string2_length
    distance(0, j) = j
  Next

  For i = 1 To string1_length
    For j = 1 To string2_length
      If bs1((i - 1) * 2) = bs2((j - 1) * 2) Then ' *2 because Unicode every 2nd byte is 0
        distance(i, j) = distance(i - 1, j - 1)
      Else
        min1 = distance(i - 1, j) + 1
        min2 = distance(i, j - 1) + 1
        min3 = distance(i - 1, j - 1) + 1
        If min1 <= min2 And min1 <= min3 Then
          distance(i, j) = min1
        ElseIf min2 <= min1 And min2 <= min3 Then
          distance(i, j) = min2
        Else
          distance(i, j) = min3
        End If

      End If
    Next
  Next

  Levenshtein = distance(string1_length, string2_length)

End Function

