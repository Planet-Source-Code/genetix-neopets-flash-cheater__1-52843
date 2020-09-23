VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl http 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   855
   ScaleHeight     =   375
   ScaleWidth      =   855
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1320
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  http://"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "http"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Score Cheater Created by Genetix
'Copyright 2004
'Removal in this copyright notice will result in prosecution
'Distribution of this source code is illegal
Option Explicit

Private Const PORT_DEFAULT_HTTP = 80


Public Enum Action
    ACTION_CONNECTED
    ACTION_REQUEST_SENT
    ACTION_FILE_TRANSFER_BEGIN
    ACTION_HEADERS_RECEIVED
    ACTION_FILE_TRANSFER_COMPLETE
    ACTION_DISCONNECTED
    ACTION_REDIRECT
    ACTION_USER_CANCEL
End Enum

'Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" _
 '  (ByRef lpdwFlags As Long, _
  ' ByVal lpszConnectionName As String, _
  ' ByVal dwNameLen As Long, _
  ' ByVal dwReserved As Long _
  ' ) As Long

Private Enum InternetConnectionState
   INTERNET_CONNECTION_MODEM = &H1&
   INTERNET_CONNECTION_LAN = &H2&
   INTERNET_CONNECTION_PROXY = &H4&
   INTERNET_RAS_INSTALLED = &H10&
   INTERNET_CONNECTION_OFFLINE = &H20&
   INTERNET_CONNECTION_CONFIGURED = &H40&
End Enum


''return error codes
Private Const ERROR_BASE = 30000
Private Const ERROR_INVALID_URL = ERROR_BASE + 1

''local data
Dim sHostName As String
Dim sFileName As String
Dim sRequestHeader As String
Dim sRequestTemplate As String
Dim iRemotePort As Integer

''winsock http info
Dim bHeadersReceived As Boolean
Dim lContentLength As Long
Dim lBytesReceived As Long

''property vars
Dim bUseProxy As Boolean
Dim sDataReceived As String
Dim sProxyHost As String
Dim iProxyPort As Integer
Dim iTimeout As Integer
Dim sHtmlData As String
Dim sDataHeader As String


''events
Event FileLoaded(FileContent As String, FileSize As Long)
Event Progress(Percent As Long, Total As Long)
Event Action(ActionNumber As Integer, Description As String)
Event Error(ErrorNumber As Integer, Description As String)
Public Function CountCharacters(Source$) As Integer
    Dim counter%, t%
    Const Characters$ = "abcdefghijklmnopqrstuvwxyz=&"


    For t% = 1 To Len(Source$)


        If InStr(Characters, LCase$(Mid$(Source$, t%, 1))) <> 0 Then
            counter% = counter% + 1
        End If
    Next t%
    CountCharacters = counter%
End Function


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ''load properties from propbag ie. property browser
    
    bUseProxy = PropBag.ReadProperty("UseProxy", False)
    sProxyHost = PropBag.ReadProperty("ProxyHost", "")
    iProxyPort = PropBag.ReadProperty("ProxyPort", 0)
    iTimeout = PropBag.ReadProperty("Timeout", 20)
    
End Sub

Private Sub UserControl_Resize()
    ''resize the control to the size of the label
    
    UserControl.Height = Label1.Height
    UserControl.Width = Label1.Width
End Sub

Public Function OpenUrl(sUrl As String, POSTorGET As String, Optional sReferer As String = "", Optional PostData As String = "", Optional Cookies As String) As Boolean
    Call Reset
    
    ''if it isn't a valid url, then exit sub
    If Not IsValidUrl(sUrl) Then
        RaiseEvent Error(ERROR_INVALID_URL, "Invalid Url")
        Exit Function
    End If
    
    ''evaluate host name
    sHostName = DetermineRemoteHostName(sUrl)
    ''evaluate remote port
    iRemotePort = DetermineRemotePort
    ''evaluate remote file name
    sFileName = DetermineRemoteFileName(sUrl)
    
    ''get request header template
    'sRequestHeader = sRequestTemplate
    
    
    'sRequestHeader = sReferer
    
    sRequestHeader = POSTorGET & " " & "_$-$_$-" & " HTTP/1.1" & Chr(13) & Chr(10) & "Host: @$@@$@" & Chr(13) & Chr(10) & "User-Agent: Mozilla/4.0 (compatible; MSIE 5.01; Windows NT 5.0)" & Chr(13) & Chr(10) & "Content-Type: application/x-www-form-urlencoded" & Chr(13) & Chr(10)
    If sReferer <> "" Then sRequestHeader = sRequestHeader & "Referer: " & sReferer & Chr(13) & Chr(10)
    If Cookies <> "" Then sRequestHeader = sRequestHeader & "Cookie: " & Cookies & Chr(13) & Chr(10)
    If PostData <> "" Then sRequestHeader = sRequestHeader & "Content-Length: " & Len(PostData) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & PostData & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    If POSTorGET = "GET" Then sRequestHeader = sRequestHeader & Chr(13) & Chr(10)
    
    
    ''and replace the unknowns with correct stuff
    sRequestHeader = Replace(sRequestHeader, "_$-$_$-", sFileName)
    sRequestHeader = Replace(sRequestHeader, "@$@@$@", sHostName)
    'MsgBox (sRequestHeader)
    ''if the referer is not "" then add the referer field to the header
    
    
    ''add a final carriage return new line for http compliance
    'sRequestHeader = sRequestHeader & Chr(13) & Chr(10)
     
    ''if using a proxy host then connect to the proxy host
    If bUseProxy Then
        Winsock1.Connect sProxyHost, iProxyPort
    Else
    ''else connect to the remote server ... ie. yahoo.com
        Winsock1.Connect sHostName, PORT_DEFAULT_HTTP
    End If
    
    ''return a true value
    OpenUrl = True
End Function











Private Sub Winsock1_Connect()
    ''upon connecting to the remote host, send the request header
    Winsock1.SendData sRequestHeader
    
End Sub

Private Sub Winsock1_Close()
    ''the close event signifies the session has ended and the file has been retrieved
    RaiseEvent Action(ACTION_FILE_TRANSFER_COMPLETE, "Download complete")
    ''close the socket since it doesn't always happen automatically
    Winsock1.Close
    ''raise another event, disconnect
    RaiseEvent Action(ACTION_DISCONNECTED, "Disconnected from server")
    ''process the response code
    Call ProcessHttpResponseCode
    
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim sBuffer As String
    Dim vHeader As Variant
    Dim iPos As Integer
    Dim iPercent As Integer
    Dim sHeaders As String
    
    ''get the data
    Call Winsock1.GetData(sBuffer, vbString)
    
    ''add the new data to the storage buffer
    sDataReceived = sDataReceived & sBuffer
    ''increment the bytes received amount
    lBytesReceived = lBytesReceived + bytesTotal
    
    ''if the headers haven't been received yet...
    If bHeadersReceived = False Then
        ''check to see if two vbcrlf's are in the message (signifies
        '' the end of the response headers and the beginning of the
        '' response data)
        iPos = InStr(1, sDataReceived, vbCrLf & vbCrLf)
        If iPos > 0 Then
            ''headers have been receieved
            bHeadersReceived = True
            ''raise event (headers received)
            RaiseEvent Action(ACTION_HEADERS_RECEIVED, "Data headers received")
            ''subtract the size of the headers from the total data size
            ''since they don't really count as part of the data
            lBytesReceived = lBytesReceived - iPos - 3
            ''the headers are left of the 2 vbcrlfs
            sHeaders = Left(sDataReceived, iPos - 1)
            
            'set local property var
            sDataHeader = sHeaders
            
            ''check to see if there was an error
            If IsHttpError(sHeaders) Then
                Winsock1.Close
                RaiseEvent Action(ACTION_DISCONNECTED, "Disconnected from server")
                Call ProcessHttpResponseCode
                Exit Sub
            End If
            
            ''retrieve content length from header
            lContentLength = CLng(Val(GetHttpHeaderValue(sHeaders, "Content-Length")))
            
        End If
    Else
        ''progress update if lContentLength was set
        If lContentLength > 0 Then
            iPercent = (lBytesReceived / lContentLength) * 100
            RaiseEvent Progress(CLng(iPercent), lContentLength)
        End If
        
        ''file transfer complete
        If lBytesReceived = lContentLength Then
            RaiseEvent Action(ACTION_FILE_TRANSFER_COMPLETE, "Download complete")
            Winsock1.Close
            RaiseEvent Action(ACTION_DISCONNECTED, "Disconnected from server")
            Call ProcessHttpResponseCode
            Exit Sub
        End If
    End If
    
End Sub


Private Function Reset()
    ''this function resets everything
    Winsock1.Close
    sDataReceived = ""
    sRequestTemplate = "GET _$-$_$- HTTP/1.0" & Chr(13) & Chr(10) & _
        "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, application/x-comet, */*" & Chr(13) & Chr(10) & _
        "Accept-Language: en" & Chr(13) & Chr(10) & _
        "Accept-Encoding: gzip , deflate" & Chr(13) & Chr(10) & _
        "Cache-Control: no-cache" & Chr(13) & Chr(10) & _
        "Proxy-Connection: Close" & Chr(13) & Chr(10) & _
        "User-Agent: SSM Agent 1.0" & Chr(13) & Chr(10) & _
        "Host: @$@@$@" & Chr(13) & Chr(10)


    bHeadersReceived = False
    lContentLength = 0
    lBytesReceived = 0
    
End Function
Public Function Cancel()
    ''reset and raise event cancel
    Call Reset
    RaiseEvent Action(ACTION_USER_CANCEL, "Action Cancelled")
    
End Function

Private Function DetermineRemoteHostName(sUrl As String)
    Dim iPos As Integer
    Dim sString As String
    
    sString = sUrl
    ''strip the http:// from the sUrl param
    sString = Right$(sString, Len(sString) - 7)
    
    iPos = InStr(sString, "/")
    If iPos > 0 Then
        DetermineRemoteHostName = Left$(sString, iPos - 1)
    Else
        DetermineRemoteHostName = sString
    End If

End Function

Private Function DetermineRemoteFileName(sUrl As String)
    Dim iPos As Integer
    Dim sHost As String
    Dim sText As String
    Dim sFile As String
    
    sText = Right$(sUrl, Len(sUrl) - 7)
    iPos = InStr(sText, "/")
    If iPos > 0 Then
        sHost = Left$(sText, iPos - 1)
        sFile = Right$(sText, Len(sText) - iPos + 1)
    Else
        sHost = sText
        sFile = "/"
    End If

    If bUseProxy = True Then
        ''if using a proxy servert than the remote file has to be absolute, not relative
        ''so it must have the http:// and the host name and the file name
        DetermineRemoteFileName = "http://" & sHost & sFile
    Else
        ''if not using a proxy server than it can be a relative path
        DetermineRemoteFileName = sFile
    End If
    
End Function

Private Function DetermineRemotePort() As Integer
    If bUseProxy = True Then
        DetermineRemotePort = iProxyPort
    Else
        DetermineRemotePort = PORT_DEFAULT_HTTP
    End If
    
End Function

Private Function IsHttpError(sHttpHeader As String) As Boolean
    ''this function determines there was an http error
    Select Case GetHttpResponseCode(sHttpHeader)
        Case 500, 501, 502, 503
            IsHttpError = True
        Case 400, 401, 403, 404
            IsHttpError = True
        Case Else
            IsHttpError = False
    End Select
    
End Function

Private Function IsValidUrl(sUrl As String) As Boolean
    If Left$(sUrl, 7) <> "http://" Then Exit Function
    IsValidUrl = True
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ''write all the properties to the propbag
    Call PropBag.WriteProperty("UseProxy", bUseProxy, False)
    Call PropBag.WriteProperty("ProxyHost", sProxyHost, "")
    Call PropBag.WriteProperty("ProxyPort", iProxyPort, "")
    Call PropBag.WriteProperty("Timeout", iTimeout, 20)
    
    
End Sub

Private Function ProcessHttpResponseCode()
    Dim sHeader As String
    Dim iPos As Integer
    
    iPos = InStr(sDataReceived, vbCrLf & vbCrLf)
    If iPos = 0 Then Exit Function
    
    sHeader = Left(sDataReceived, iPos - 1)
    sHtmlData = Right$(sDataReceived, Len(sDataReceived) - iPos - 3)
    
    Select Case GetHttpResponseCode(sHeader)
        ''500s and 400s are errors, raiseevent error
        Case 500, 501, 502, 503
            RaiseEvent Error(GetHttpResponseCode(sHeader), GetHttpResponseDescription(sHeader))
        Case 400, 401, 403, 404
            RaiseEvent Error(GetHttpResponseCode(sHeader), GetHttpResponseDescription(sHeader))
        ''300s are all redirects, so redirect the request
        Case 300, 301, 302, 303, 307
            'Call RedirectRequest(sHeader)
            RaiseEvent FileLoaded(sHtmlData, lContentLength)
        ''200s are good, they mean success
        Case 200, 202
            RaiseEvent FileLoaded(sHtmlData, lContentLength)
        
    End Select
    
End Function

Private Function GetHttpResponseCode(sHttpHeader As String) As Long
    Dim sHeaders() As String
    
    ''get the http response number from the headers, it is in the first header
    sHeaders = Split(sHttpHeader, vbCrLf)
    GetHttpResponseCode = Val(Mid$(sHeaders(0), InStr(sHeaders(0), Chr(32))))
    
End Function

Private Function GetHttpResponseDescription(sHttpHeader As String) As String
    Dim sHeaders() As String
    
    ''get the description after the code, ie. 200 OK
    ''will return OK
    sHeaders = Split(sHttpHeader, vbCrLf)
    GetHttpResponseDescription = Mid$(sHeaders(0), InStr(sHeaders(0), Chr(32)))
    
    
End Function
Private Function GetHttpHeaderValue(sHttpHeader As String, sValueName As String) As String
    Dim sHeaders() As String
    Dim vHeaders As Variant
    
    ''gets the value for corresponding header ie.
    ''File Content: pkzip file
    ''if sValueName is File Content than returns pkzip file
    
    sHeaders = Split(sHttpHeader, vbCrLf)
    For Each vHeaders In sHeaders
        vHeaders = LCase(vHeaders)
        If InStr(vHeaders, LCase(sValueName)) > 0 Then
            GetHttpHeaderValue = Trim$(Mid$(vHeaders, InStr(vHeaders, Chr(32))))
            Exit Function
        End If
    Next
    
End Function
Private Sub RedirectRequest(sHttpHeader As String)
    'Dim sHeaders() As String
    'Dim vHeader As Variant
    'Dim sNewUrl As String
    
    
    ''processes the redirection
    ''get the new url from the headers
    'sNewUrl = GetHttpHeaderValue(sHttpHeader, "Location")
    'If InStr(sNewUrl, "http://") Then
        'RaiseEvent Action(ACTION_REDIRECT, "Redirecting to " & sNewUrl)
        ''and open it
'        Call OpenUrl(sNewUrl)
        'Exit Sub
 '   Else
  '      RaiseEvent Action(ACTION_REDIRECT, "Redirecting to http://" & sHostName & sNewUrl)
   '     Call OpenUrl("http://" & sHostName & sNewUrl)
    '    Exit Sub
    'End If
        
End Sub


''everything below is just property handling code
''no need for comments i hope =)



Public Property Get UseProxy() As Boolean
    UseProxy = bUseProxy
End Property

Public Property Let UseProxy(ByVal bNewValue As Boolean)
    bUseProxy = bNewValue
End Property

Public Property Get DataHeader() As String
    DataHeader = sDataHeader
End Property
Public Property Get TotalData() As String
    TotalData = sDataReceived
End Property

Public Property Get htmldata() As String
    htmldata = sHtmlData
End Property

Public Property Get ProxyHost() As String
    ProxyHost = sProxyHost
End Property

Public Property Let ProxyHost(sNewValue As String)
    sProxyHost = sNewValue
End Property

Public Property Get ProxyPort() As Integer
    ProxyPort = iProxyPort
End Property

Public Property Let ProxyPort(iNewValue As Integer)
    iProxyPort = iNewValue
End Property

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent Error(Number, Description)
End Sub

Private Sub Winsock1_SendComplete()
    RaiseEvent Action(ACTION_REQUEST_SENT, "Request sent")
End Sub

Public Property Let Timeout(iNewValue As Integer)
    iTimeout = iNewValue
End Property

Public Property Get Timeout() As Integer
    Timeout = iTimeout
End Property






