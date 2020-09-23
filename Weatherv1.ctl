VERSION 5.00
Begin VB.UserControl Weather 
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1440
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1080
   ScaleWidth      =   1440
   Begin VB.Image Image1 
      Height          =   690
      Left            =   0
      Picture         =   "Weatherv1.ctx":0000
      Top             =   -15
      Width           =   885
   End
End
Attribute VB_Name = "Weather"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

Private Const IF_FROM_CACHE = &H1000000
Private Const IF_MAKE_PERSISTENT = &H2000000
Private Const IF_NO_CACHE_WRITE = &H4000000
       
Private Const BUFFER_LEN = 256

Dim Text

Private m_ZIP As Integer
Private m_Temp As String
Private m_Cond As String
Private m_Humid As String
Private m_Feel

' property defaults
Private Const m_def_ZIP = "11756"
Private Const m_def_Temp = "N/A"
Private Const m_def_Cond = "N/A"
Private Const m_def_Humid = "N/A"
Private Const m_def_Feel = "N/A"

Private Function GetUrlSource(sURL As String) As String
    Dim sBuffer As String * BUFFER_LEN, iResult As Integer, sData As String
    Dim hInternet As Long, hSession As Long, lReturn As Long

    'get the handle of the current internet connection
    hSession = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)
    'get the handle of the url
    If hSession Then hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, IF_NO_CACHE_WRITE, 0)
    'if we have the handle, then start reading the web page
    If hInternet Then
        'get the first chunk & buffer it.
        iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
        sData = sBuffer
        'if there's more data then keep reading it into the buffer
        Do While lReturn <> 0
            iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
            sData = sData + Mid(sBuffer, 1, lReturn)
        Loop
    End If
   
    'close the URL
    iResult = InternetCloseHandle(hInternet)

    GetUrlSource = sData
End Function
Private Function pReplace(strExpression As String, strFind As String, strReplace As String)
    Dim intX As Integer


    If (Len(strExpression) - Len(strFind)) >= 0 Then


        For intX = 1 To Len(strExpression)


            If Mid(strExpression, intX, Len(strFind)) = strFind Then
                strExpression = Left(strExpression, (intX - 1)) + strReplace + Mid(strExpression, intX + Len(strFind), Len(strExpression))
            End If
        Next
    End If
    pReplace = strExpression
End Function
Public Sub GetTemp()
On Error Resume Next
Dim Search
Dim Spot As Integer
Dim Spot2 As Integer
Dim done
Search = "insert current temp"
Spot = InStr(Text, Search) + 23
done = Mid(Text, Spot, 20)
Spot2 = InStr(done, "</B>")
Spot2 = Spot2 - 1
done = Mid(done, 1, Spot2)
done = pReplace(Trim(done), "&deg;", "° ")

m_Temp = done
End Sub
Public Sub GetHumidity()
On Error Resume Next
Dim Search
Dim Spot As Integer
Dim Spot2 As Integer
Dim done
Search = "insert humidity"
Spot = InStr(Text, "insert humidity") + 21
done = Mid(Text, Spot - 2, 100)
Spot2 = InStr(done, "</td>")
done = Mid(done, 1, Spot2 - 1)
Dim ZIP
done = Trim(done)

m_Humid = done
End Sub
Public Sub GetCondition()
On Error Resume Next
Dim Search
Dim Spot As Integer
Dim Spot2 As Integer
Dim done
Search = "insert forecast text"
Spot = InStr(Text, Search) + 24
done = Mid(Text, Spot, 100)
Spot2 = InStr(done, "</td>")
done = Mid(done, 1, Spot2 - 1)
done = pReplace(Trim(done), "&deg;", "° ")

m_Cond = done
End Sub
Public Sub GetFeelsLike()
On Error Resume Next
Dim Search
Dim Spot As Integer
Dim Spot2 As Integer
Dim done
Search = "insert feels like temp"
Spot = InStr(Text, Search) + 26
done = Mid(Text, Spot, 20)
Spot2 = InStr(done, "</font>")
done = Mid(done, 1, Spot2 - 1)
done = pReplace(Trim(done), "&deg;", "° ")

done = pReplace(Trim(done), "Feels Like: ", "")

m_Feel = done
End Sub

Private Sub UserControl_InitProperties()
m_ZIP = m_def_ZIP
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
m_ZIP = PropBag.ReadProperty("ZipCode", m_def_ZIP)
m_Temp = PropBag.ReadProperty("Temp", m_def_ZIP)
m_Cond = PropBag.ReadProperty("Condition", m_def_Cond)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ZipCode", m_ZIP, m_def_ZIP
        PropBag.WriteProperty "Temp", m_Temp, m_def_Temp
        PropBag.WriteProperty "Condition", m_Cond, m_def_Cond
End Sub
Public Sub UpdateWeather()
Text = GetUrlSource("http://www.weather.com/weather/local/" & m_ZIP)
GetTemp
GetCondition
GetHumidity
GetFeelsLike
End Sub

Public Property Get ZIP() As Integer
    ZIP = m_ZIP
End Property
Public Property Get Temp() As String
    Temp = m_Temp
End Property

Public Property Let ZIP(ByVal vNewValue As Integer)
    m_ZIP = vNewValue
    PropertyChanged "Zip Code"
End Property
Public Property Get Condition() As String
Condition = m_Cond
End Property
Public Property Get Humidity() As String
Humidity = m_Humid
End Property
Public Property Get FeelsLike() As String
FeelsLike = m_Feel
End Property
