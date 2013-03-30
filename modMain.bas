Attribute VB_Name = "modMain"
Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset

Public strSMSNum As String
Public strURL As String
Public colSenderGSM As Collection
Public colSenderCDMA As Collection
Private m_SafeChar(0 To 255) As Boolean
Private Const FLAG_ICC_FORCE_CONNECTION = &H1
Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long

Public Function CountChar(ByVal txt As String) As String

Dim cntMsg As Integer
Dim cntChar As Integer
cntMsg = Int((Len(txt) / 160)) + 1
cntChar = (160 * cntMsg) - Len(txt)
CountChar = cntMsg & "/" & cntChar & " " & "characters available"
End Function


Public Function URLEncode(ByVal txt As String) As String
Dim i As Integer
Dim ch As String
Dim ch_asc As Integer
Dim result As String

    SetSafeChars

    result = ""
    For i = 1 To Len(txt)
        ' Translate the next character.
        ch = Mid$(txt, i, 1)
        ch_asc = Asc(ch)
        If ch_asc = vbKeySpace Then
            ' Use a plus.
            result = result & "+"
        ElseIf m_SafeChar(ch_asc) Then
            ' Use the character.
            result = result & ch
        Else
            ' Convert the character to hex.
            result = result & "%" & Right$("0" & _
                Hex$(ch_asc), 2)
        End If
    Next i

    URLEncode = result
End Function


' Set m_SafeChar(i) = True for characters that
' do not need protection.
Private Sub SetSafeChars()
Static done_before As Boolean
Dim i As Integer

    If done_before Then Exit Sub
    done_before = True

    For i = 0 To 47
        m_SafeChar(i) = False
    Next i
    For i = 48 To 57
        m_SafeChar(i) = True
    Next i
    For i = 58 To 64
        m_SafeChar(i) = False
    Next i
    For i = 65 To 90
        m_SafeChar(i) = True
    Next i
    For i = 91 To 96
        m_SafeChar(i) = False
    Next i
    For i = 97 To 122
        m_SafeChar(i) = True
    Next i
    For i = 123 To 255
        m_SafeChar(i) = False
    Next i
End Sub

Public Function IsAllNumbers(ByVal txt As String) As Boolean
Dim intLen, i As Integer
Dim intAChar As Integer
intLen = Len(txt)
IsAllNumbers = True
If intLen < 10 Then IsAllNumbers = False: Exit Function
For i = 1 To intLen
    intAChar = Asc(Mid(txt, i, 1))
        
    If intAChar < 48 Or intAChar > 57 Then
        IsAllNumbers = False
        Exit For
    End If
Next
End Function





Public Function IsConnectedToInternet() As Boolean
DoEvents
If InternetCheckConnection("http://www.google.com", FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
    IsConnectedToInternet = False
Else
    IsConnectedToInternet = True
End If

'' This DOESN'T work
'If InternetCheckConnection("192.169.2.40", FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
'MsgBox "Connection to 192.169.2.40 failed!", vbInformation
'Else
'MsgBox "Connection to 192.169.2.40 succeeded!", vbInformation
'End If
End Function
