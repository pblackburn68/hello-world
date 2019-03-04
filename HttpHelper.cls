Option Explicit

' Requires reference to Microsostf Win Http Services
Private WithEvents m_http As WinHttp.WinHttpRequest

Private m_doc      As HtmlDoc
Private m_Proxy    As String
Private m_Busy     As Boolean

Public Event OnGetStringCompleted(html As String)
Public Event OnGetHtmlCompleted(doc As HtmlDoc)
Public Event OnError(Err As ErrObject)

Public Property Get Proxy() As String
  Proxy = m_Proxy
End Property
Public Property Let Proxy(value As String)
  m_Proxy = value
End Property

Public Property Get HtmlDoc() As HtmlDoc
  Set HtmlDoc = m_doc
End Property

Public Property Get Response() As String
  Response = m_http.responseText
End Property

Public Property Get Busy() As Boolean
  Busy = m_Busy
End Property

Public Function GetHtml(Url As String, Optional async As Boolean = False) As HtmlDoc
  Set m_doc = New HtmlDoc
  Dim html As String
  
  html = GetAsString(Url, async)
  
  If Not async Then
    m_doc.WriteHtml html
    Set GetHtml = m_doc
  End If
End Function


Public Function GetAsString(Url As String, Optional async As Boolean = False) As String
On Error GoTo ERR_CATCH
  If m_Busy Then
    Err.Raise vbObjectError + 1000, "HttpHelper", "Currently waiting for a http get request to respond"
    Exit Function
  End If
  
  m_Busy = True
  
  With m_http
      .Open "GET", Url, async
      .SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
      If Len(m_Proxy) > 0 Then .SetProxy 2, m_Proxy
      .send
      If Not async Then
        GetAsString = .responseText
        m_Busy = False
      End If
  End With
  Exit Function
ERR_CATCH:
  If MsgBox("Failed loading URL " & Url & vbCrLf & "Error : " & Err.message, vbRetryCancel, "Http Error") = vbRetry Then
    Resume Next
  Else
    RaiseEvent OnError(Err)
  End If
End Function


Public Function PostForm(Url As String, FieldCollection As Dictionary, Optional async As Boolean = False) As String
On Error GoTo ERR_CATCH
  
  If m_Busy Then
    Err.Raise vbObjectError + 1000, "HttpHelper", "Currently waiting for a http get request to respond"
    Exit Function
  End If
  Dim mpf As MultiPartFormData
  Set mpf = New MultiPartFormData
  
  mpf.AddFormField "l", 1
  mpf.AddFormField "lq", "London"
  mpf.AddFormField "rad", 5
  mpf.AddFormField "q", ""
  
  Debug.Print mpf.ContentType
  Debug.Print mpf.ContentDisposition
  
  m_Busy = True
  
  With m_http
      .Open "POST", Url, async
      .SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
      .SetRequestHeader "Content-type", mpf.ContentType
      
      If Len(m_Proxy) > 0 Then .SetProxy 2, m_Proxy

      .send mpf.ContentDisposition
      
      If Not async Then
        PostForm = .responseText
        m_Busy = False
      End If
  End With
  Exit Function
ERR_CATCH:
  If MsgBox("Failed loading URL " & Url & vbCrLf & "Error : " & Err.message, vbRetryCancel, "Http Error") = vbRetry Then
    Resume Next
  Else
    RaiseEvent OnError(Err)
  End If
End Function



Public Function LoadAsString(fileName As String) As String
  Dim fso As New FileSystemObject
  Dim ts As TextStream
  Set ts = fso.OpenTextFile(fileName)
  If Not ts Is Nothing Then
    LoadAsString = ts.ReadAll
    ts.Close
    Set ts = Nothing
  End If
  Set fso = Nothing
End Function

Public Function LoadHtml(fileName As String) As HtmlDoc
  Dim html As String
  html = Me.LoadAsString(fileName)
  Set LoadHtml = New HtmlDoc
  LoadHtml.WriteHtml html
End Function


Private Sub Class_Initialize()
  Set m_http = New WinHttp.WinHttpRequest
End Sub

Private Sub m_http_OnResponseFinished()
  m_Busy = False
  Dim html As String
  html = m_http.responseText
  If Not m_doc Is Nothing Then
    m_doc.WriteHtml html
  End If
  RaiseEvent OnGetStringCompleted(html)
  RaiseEvent OnGetHtmlCompleted(m_doc)
End Sub

