Attribute VB_Name = "Module1"
Option Explicit


' Initializes an application's use of the Win32 Internet functions
Public Declare Function InternetOpen Lib "wininet.dll" _
    Alias "InternetOpenA" _
        (ByVal sAgent As String _
        , ByVal lAccessType As Long _
        , ByVal sProxyName As String _
        , ByVal sProxyBypass As String _
        , ByVal lFlags As Long) As Long


' Use registry access settings.
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3
Public Const INTERNET_INVALID_PORT_NUMBER = 0

Public Const FTP_TRANSFER_TYPE_ASCII = &H1
Public Const FTP_TRANSFER_TYPE_BINARY = &H2


' Opens a HTTP session for a given site.
Public Declare Function InternetConnect Lib "wininet.dll" _
    Alias "InternetConnectA" _
        (ByVal hInternetSession As Long _
        , ByVal sServerName As String _
        , ByVal nServerPort As Integer _
        , ByVal sUsername As String _
        , ByVal sPassword As String _
        , ByVal lService As Long _
        , ByVal lFlags As Long _
        , ByVal lContext As Long) As Long

' Closes a single Internet handle or a subtree of Internet handles.
Public Declare Function InternetCloseHandle Lib "wininet.dll" _
        (ByVal hInet As Long) As Integer


' Adds one or more HTTP request headers to the HTTP request handle.
Public Declare Function HttpAddRequestHeaders Lib "wininet.dll" _
    Alias "HttpAddRequestHeadersA" _
        (ByVal hHttpRequest As Long _
        , ByVal sHeaders As String _
        , ByVal lHeadersLength As Long _
        , ByVal lModifiers As Long) As Integer

' Flags to modify the semantics of this function. Can be a combination of these values:

' Adds the header only if it does not already exist; otherwise, an error is returned.
Public Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000

' Adds the header if it does not exist. Used with REPLACE.
Public Const HTTP_ADDREQ_FLAG_ADD = &H20000000

' Replaces or removes a header. If the header value is empty and the header is found,
' it is removed. If not empty, the header value is replaced
Public Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000


' Number of the TCP/IP port on the server to connect to.
Public Const INTERNET_DEFAULT_FTP_PORT = 21
Public Const INTERNET_DEFAULT_GOPHER_PORT = 70
Public Const INTERNET_DEFAULT_HTTP_PORT = 80
Public Const INTERNET_DEFAULT_HTTPS_PORT = 443
Public Const INTERNET_DEFAULT_SOCKS_PORT = 1080

Public Const INTERNET_OPTION_CONNECT_TIMEOUT = 2
Public Const INTERNET_OPTION_RECEIVE_TIMEOUT = 6
Public Const INTERNET_OPTION_SEND_TIMEOUT = 5

Public Const INTERNET_OPTION_USERNAME = 28
Public Const INTERNET_OPTION_PASSWORD = 29
Public Const INTERNET_OPTION_PROXY_USERNAME = 43
Public Const INTERNET_OPTION_PROXY_PASSWORD = 44

' Type of service to access.
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3


' Opens an HTTP request handle.
Private Declare Function HttpOpenRequest Lib "wininet.dll" _
    Alias "HttpOpenRequestA" _
        (ByVal hConnect As Long _
        , ByVal lpszVerb As String _
        , ByVal lpszObjectName As String _
        , ByVal lpszVersion As String _
        , ByVal lpszReferer As String _
        , ByVal lpszAcceptTypes As Long _
        , ByVal dwFlags As Long _
        , ByVal dwContext As Long) As Long


' Sends the specified request to the HTTP server.
Private Declare Function HttpSendRequest Lib "wininet.dll" _
    Alias "HttpSendRequestA" _
(ByVal hRequest As Long _
, ByVal lpszHeaders As String _
, ByVal dwHeadersLength As Long _
, ByVal lpOptional As String _
, ByVal dwOptionalLength As Long) As Integer


' Queries for information about an HTTP request.
Public Declare Function HttpQueryInfo _
            Lib "wininet.dll" _
        Alias "HttpQueryInfoA" _
        (ByVal hHttpRequest As Long _
        , ByVal lInfoLevel As Long _
        , ByRef sBuffer As Any _
        , ByRef lBufferLength As Long _
        , ByRef lIndex As Long) As Integer


' The possible values for the lInfoLevel parameter include:
Public Const HTTP_QUERY_CONTENT_TYPE = 1
Public Const HTTP_QUERY_CONTENT_LENGTH = 5
Public Const HTTP_QUERY_EXPIRES = 10
Public Const HTTP_QUERY_LAST_MODIFIED = 11
Public Const HTTP_QUERY_PRAGMA = 17
Public Const HTTP_QUERY_VERSION = 18
Public Const HTTP_QUERY_STATUS_CODE = 19
Public Const HTTP_QUERY_STATUS_TEXT = 20
Public Const HTTP_QUERY_RAW_HEADERS = 21
Public Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Public Const HTTP_QUERY_FORWARDED = 30
Public Const HTTP_QUERY_SERVER = 37
Public Const HTTP_QUERY_USER_AGENT = 39
Public Const HTTP_QUERY_SET_COOKIE = 43
Public Const HTTP_QUERY_REQUEST_METHOD = 45
Public Const HTTP_STATUS_DENIED = 401
Public Const HTTP_STATUS_PROXY_AUTH_REQ = 407


'
' flags common to open functions (not InternetOpen()):
'

Public Const INTERNET_FLAG_RELOAD = &H80000000             ' retrieve the original item


'
' additional cache flags
'

Public Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000      ' don't write this item to the cache
Public Const INTERNET_FLAG_DONT_CACHE = INTERNET_FLAG_NO_CACHE_WRITE
Public Const INTERNET_FLAG_MAKE_PERSISTENT = &H2000000     ' make this item persistent in cache
Public Const INTERNET_FLAG_FROM_CACHE = &H1000000          ' use offline semantics
Public Const INTERNET_FLAG_OFFLINE = INTERNET_FLAG_FROM_CACHE


'##################################################
'Set the following Constants.
'##################################################

Private Const YOUR_HOST = "www.hogehoge.com"
Private Const YOUR_FILE_PATH = "/index.html"
Private Const USER_AGENT = "User-Agent: Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.90 Safari/537.36"

Sub Sub_HTTPSendRequest()

    
    Dim lngWinINet  As Long 'Internet Handler
    Dim lngHttpHnd  As Long 'HTTP Handler
    Dim lngReqHnd   As Long 'HTTP Request Handler
    
    Dim strTmpURL      As String * 255
    
On Error GoTo ErrorHandler
    

    'Open Internet Service and Get Internet Handler
    lngWinINet = InternetOpen(vbNullString _
                            , INTERNET_OPEN_TYPE_PRECONFIG _
                            , vbNullString _
                            , vbNullString _
                            , 0)


    'Connect to HTTP Server and Get HTTP Handler
    lngHttpHnd = InternetConnect(lngWinINet _
                              , YOUR_HOST _
                              , INTERNET_DEFAULT_HTTP_PORT _
                              , vbNullString _
                              , vbNullString _
                              , INTERNET_SERVICE_HTTP _
                              , 0 _
                              , 0)


    'URL is 255 Bytes fixed length
    strTmpURL = YOUR_FILE_PATH

    'Initialization of request and Get HTTP Request Handler
    lngReqHnd = HttpOpenRequest(lngHttpHnd _
                              , "GET" _
                              , strTmpURL _
                              , "HTTP/1.1" _
                              , vbNullString _
                              , 0 _
                              , INTERNET_FLAG_RELOAD _
                              , 0)

    'Send request
    Call HttpSendRequest(lngReqHnd _
                       , vbNullString _
                       , 0 _
                       , vbNullString _
                       , 0)

    'Close HTTP request
    Call InternetCloseHandle(lngReqHnd)
    

    'Disconnect
    Call InternetCloseHandle(lngHttpHnd)


    'Close Internet Service
    Call InternetCloseHandle(lngWinINet)
    


    Exit Sub

ErrorHandler:

    MsgBox Err.Number & ":" & Err.Description, vbCritical & vbOKOnly, "Error!"

End Sub


Sub Sub_GetHTTPStatus()

    Dim lngRC As Long
    
    Dim lngWinINet  As Long     'Internet Handler
    Dim lngHttpHnd  As Long     'HTTP Handler
    Dim lngReqHnd   As Long     'HTTP Request Handler
    
    Dim strHeader   As String   'HTTP Header
    
    Dim strTmpURL   As String * 255
    Dim lngTmpIndex As Long
    
    Dim strBuffer   As String * 1024 'Receive buffer
    Dim lngLength   As Long     'Response Data length
    
    
On Error GoTo ErrorHandler
    
    'Open Internet Service and Get Internet Handler
    lngWinINet = InternetOpen(vbNullString _
                            , INTERNET_OPEN_TYPE_PRECONFIG _
                            , vbNullString _
                            , vbNullString _
                            , 0)
    
    
    'Connect to HTTP Server and Get HTTP Handler
    lngHttpHnd = InternetConnect(lngWinINet _
                              , YOUR_HOST _
                              , INTERNET_DEFAULT_HTTP_PORT _
                              , vbNullString _
                              , vbNullString _
                              , INTERNET_SERVICE_HTTP _
                              , 0 _
                              , 0)
                              
    
    'URL is 255 Bytes fixed length
    strTmpURL = YOUR_FILE_PATH

    'Initialization of request and Get HTTP Request Handler
    lngReqHnd = HttpOpenRequest(lngHttpHnd _
                              , "GET" _
                              , strTmpURL _
                              , "HTTP/1.1" _
                              , vbNullString _
                              , 0 _
                              , INTERNET_FLAG_RELOAD _
                              , 0)
                              
    'Set User Agent
    strHeader = USER_AGENT

    'Set HTTP Header
    Call HttpAddRequestHeaders(lngReqHnd _
                                , strHeader _
                                , Len(strHeader) _
                                , HTTP_ADDREQ_FLAG_REPLACE _
                               Or HTTP_ADDREQ_FLAG_ADD)


    'Send request
    Call HttpSendRequest(lngReqHnd _
                       , vbNullString _
                       , 0 _
                       , vbNullString _
                       , 0)



    'Initialization
    lngLength = Len(strBuffer)
    strBuffer = vbNullString
    lngTmpIndex = 0


    'Get response from HTTP server
    Call HttpQueryInfo(lngReqHnd _
                     , HTTP_QUERY_STATUS_CODE _
                     , ByVal strBuffer _
                     , lngLength _
                     , lngTmpIndex)

    'Show HTTP status
    MsgBox Left(strBuffer, lngLength)
    
    
    'Close HTTP request
    Call InternetCloseHandle(lngReqHnd)
    

    'Disconnect
    Call InternetCloseHandle(lngHttpHnd)


    'Close Internet Service
    Call InternetCloseHandle(lngWinINet)


    Exit Sub

ErrorHandler:

    MsgBox Err.Number & ":" & Err.Description, vbCritical & vbOKOnly, "Error!"

End Sub


