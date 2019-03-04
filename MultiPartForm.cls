Option Explicit

Private m_boundary  As String
Private m_fields    As Dictionary

Public Function AddFormField(name As String, value As String)
  Call m_fields.Add(name, value)
  
End Function

Public Property Get ContentDisposition() As String
  Dim vName As Variant
  Dim ff As String
  For Each vName In m_fields
    ff = ff & m_boundary & vbCrLf
    ff = ff & "Content-Disposition: form-data; name=""" & vName & """" & vbCrLf & vbCrLf
    ff = ff & m_fields(vName) & vbCrLf
  Next
  ContentDisposition = ff
End Property

Public Property Get ContentType() As String
  ContentType = "multipart/form-data; boundary=" & m_boundary
End Property

Public Property Get Boundary() As String
  Boundary = m_boundary
End Property

Private Sub Class_Initialize()
    m_boundary = String(10, "-") & RandonString(6)
    Set m_fields = New Dictionary
End Sub


Private Function RandonString(length As Integer) As String
    Dim c As Integer, idx As Integer
    Dim s As String
    Randomize
    For idx = 1 To length
        s = s & Chr(Random(65, 90))
    Next
    RandonString = s
End Function


Private Function Random(lower As Integer, upper As Integer) As Integer
    Random = Int((upper - lower + 1) * Rnd + lower)
End Function

