        ��  ��                  u  ,   C U S T O M   ��          	        VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Multi part HTTP File Download"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin Project1.LSInetDownload LSInetDownload 
      Height          =   315
      Index           =   0
      Left            =   6300
      TabIndex        =   9
      Top             =   1080
      Width           =   2895
      _extentx        =   5106
      _extenty        =   556
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5940
      TabIndex        =   6
      Text            =   "80"
      Top             =   60
      Width           =   615
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3360
      Left            =   60
      TabIndex        =   5
      Top             =   1080
      Width           =   6195
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Text            =   "http://127.0.0.1/3prof.reg"
      Top             =   60
      Width           =   3075
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DOWNLOAD NOW!"
      Height          =   315
      Left            =   6960
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   1
      Top             =   0
      Width           =   0
   End
   Begin VB.PictureBox Picture2 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   2
      Top             =   0
      Width           =   0
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please vote! Do you understand that I make this hard work for you?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   540
      Width           =   5055
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Width           =   6675
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "port"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   8
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label2 
      Caption         =   "File will be downloaded in App.Path\downloaded.file."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   7
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Url (with HTTP:\\)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fd$()

Private Sub Command1_Click()
 'Shell "start http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=25577&lngWId=1", vbHide
 
 If Dir$(App.Path + "\downloaded.file") <> "" Then Kill App.Path + "\downloaded.file"
 ur$ = Text1.Text
 pr = Text2.Text
 LSInetDownload(0).url = ur$
 LSInetDownload(0).port = pr
 LSInetDownload(0).GetInformation rs, tl, er, 30
 Label4.Caption = ""
 If rs = True Then
  Label4.Caption = "RESUME SUPPORTED."
 Else
  Label4.Caption = "Resume not supported."
 End If
 Label4.Caption = Label4.Caption + "   File length:" & tl
 If er <> 0 Then
  'MsgBox "Can't connect to server. Error: " & er
  Exit Sub
 End If
 If rs = True Then 'If resume supported...
  DownloadWithThreads tl
 Else 'If resume not supported :(
  ReDim fd$(0)
  LSInetDownload(0).DownloadFile
 End If
End Sub

Sub DownloadWithThreads(tl)
'Total threads
  dl = 3
  
  
ur$ = Text1.Text
 pr = Text2.Text
 'Evaluating size of thread
  sz = CLng(tl / dl)
On Error Resume Next
f = 0
ReDim fd$(dl - 1)
For q = 0 To dl - 1
 Load LSInetDownload(q)
 LSInetDownload(q).url = ur$
 LSInetDownload(q).port = 80
 If q > 0 Then
  LSInetDownload(q).Left = LSInetDownload(q - 1)
  LSInetDownload(q).Top = LSInetDownload(q - 1).Top + LSInetDownload(q - 1).Height
  LSInetDownload(q).Visible = True
 End If
 LSInetDownload(q).DownloadFile 30, f, sz, Int(Rnd * 512) + 512 'Random Chunk size
 f = (f + 1) + sz
 'f = f + sz
 List1.AddItem ""
Next q
End Sub



Private Sub LSInetDownload_DownloadComplete(Index As Integer, SeekPosition As Long)
 Open App.Path + "\downloaded.file" For Binary As 1
  Seek #1, SeekPosition + 1
  Put #1, , fd$(Index)
 Close 1
End Sub

Private Sub LSInetDownload_DownloadNeedRestart(Index As Integer, BeFrom As Variant)
 LSInetDownload(Index).DownloadFile , BeFrom
End Sub

Private Sub LSInetDownload_DownloadProgress(Index As Integer, Procent As Variant, CurByte As Variant, TotBytes As Variant, data As String, cps As Long)
 List1.List(Index) = Procent & "%. " & CurByte & " bytes of " & TotBytes & ". CPS: " & cps
 fd$(Index) = fd$(Index) + data
End Sub

'
Private Sub LSInetDownload_StateChanged(Index As Integer, state As String)
 List1.List(Index) = state
End Sub

   6.  ,   C U S T O M   ��                   VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.UserControl LSInetDownload 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   1740
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2040
      Top             =   1620
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   600
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1260
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FF0000&
      Height          =   435
      Left            =   180
      Top             =   60
      Width           =   1395
   End
End
Attribute VB_Name = "LSInetDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------------------------
'                  Multi Part HTTP file download.
'                             BETA 1.0
'      Copyright (c) 2001 LeSaMiro Software.  All rights reserved.
'              http://lesamirosoftware.hypermart.net
' Autor of UserControl: Lev Zabudkin.
'
' You have a royalty-free right to use, modify, reproduce and distribute
' this Sample (and/or any modified version) in any way
' you find useful, provided that you agree that LeSaMiro Software has
' no warranty, obligations or liability for this Sample.
' ------------------------------------------------------------------------
'-------------------------------------------------------------------------
'This UserControl provides methods needed for downloading files
'from InterNet/IntraNet/etc...
'-------------------------------------------------------------------------
'
' Please vote !
'
'>>>You should understand that I make this hard work
'   for you.<<<
'>>>Do you understand?

'Hmm... variables :)
Dim ls, DownloadTimeOut As Double, TDownloadTimeOut As Double
Dim CurByte
Dim SekElapsed As Double
Dim TotBytes, DownLen
Dim BeFrom As Long
Dim DownChunkSize As Long
Dim lastUrl$, lastres As Boolean
Dim downstart As Boolean
Dim downdata As String
'Events...
Event DownloadProgress(Procent, CurByte, TotBytes, data As String, cps As Long)
Event DownloadComplete(SeekPosition As Long)
Event StateChanged(state$)
Event DownloadNeedRestart(BeFrom)
'Default Property Values:
Const m_def_Url = 0
Const m_def_Port = 0
'Property Variables:
Dim m_Url As Variant
Dim m_Port As Variant

'This procedure will return:
'ResumeSupport=true - if resume supported by server
'TotalFileLength    - if server return Length of file
'IfError            - return 0 if no errors,
'                   - or -1 if can't connect to server,
'                   - or
'                     server answer (for example: 404)
Public Sub GetInformation(ResumeSupport, TotalFileLength, IfError, Optional TimeOut = 10)
 RaiseEvent StateChanged("Connecting...")
 ih$ = "Connection: close" + vbCrLf
 'Checking for Resume Support
 ih$ = ih$ & "Range: bytes=2-" & vbCrLf
 Inet1.RemotePort = port
 'Get Url HTTP/1.x... etc..
 Inet1.Execute url, "GET", , ih$
 'For checking timeout
 nt# = Timer + TimeOut
 Do
  DoEvents
  'TimeOut ?
  If nt# < Timer Then
   IfError = -1
   RaiseEvent StateChanged("Can't connect to server.")
   RaiseEvent DownloadNeedRestart(BeFrom)
   Exit Sub
  End If
 Loop While ls <> 12
 RaiseEvent StateChanged("Getting information...")
 'Getting header from server
 h$ = Inet1.GetHeader
 'Getting response code from header
 c = GetHttpResponseCode(h$)
 'Default - no errors
 IfError = 0
 'Server can resume
 If c = 206 Then
  ResumeSupport = True
  'Getting file length from header
  TotalFileLength = GetHttpHeaderFieldValue(h$, "content-length")
  'adding 2 bytes... because we checking for resume
  'from 2 byte of file
  TotalFileLength = TotalFileLength + 2
  a = 1
 ElseIf c = 200 Then 'Server can't resume :(
  ResumeSupport = False
  'Getting file length from header
  TotalFileLength = GetHttpHeaderFieldValue(h$, "content-length")
 Else
  'Not 200 or 206.
  'Returning response code
  IfError = c
 End If
 RaiseEvent StateChanged("Getting information... Succesfully")
End Sub

Sub StopDownload()
 downstart = False
 Timer1.Enabled = False
 Timer2.Enabled = False
End Sub

Function DownloadFile(Optional TimeOut = 10, Optional from = -1, Optional bytes = -1, Optional ChunkSize = 2048, Optional Retries = 0) As Integer
 Timer1.Enabled = False
 downstart = False
 Timer2.Enabled = True
 CurByte = 0
 SekElapsed = 0
 RaiseEvent StateChanged("Start downloading...")
 DownChunkSize = ChunkSize
 ih$ = "Connection: close" + vbCrLf
 If from >= 0 Then
  fr$ = Trim$(Str$(from))
 Else
  fr$ = ""
 End If
 If bytes = -1 Then
  b$ = ""
 ElseIf fr$ <> "" Then
   If from = 0 Then
    b$ = Trim$(Str$(from + bytes + 1))
   Else
    b$ = Trim$(Str$(from + bytes))
   End If
 End If
 'if start position specified...
 If fr$ <> "" Then
  'then make the additional header
  ih$ = ih$ & "Range: bytes=" + fr$ + "-" + b$ & vbCrLf
  BeFrom = Val(fr$)
 Else
  BeFrom = 0
 End If
 cret = 0
Retry:
 'Setting remote port
 Inet1.RemotePort = port
 DoEvents
 'We send GET command to server
 Inet1.Execute url, "GET", , ih$
 DoEvents
  nt# = Timer + TimeOut
 Do
  DoEvents
  'TimeOut ?
  If nt# < Timer Then
   If Retries > 0 And cret < Retries Then
    cret = cret + 1
    GoTo Retry
   End If
   RaiseEvent StateChanged("Error while starting downloading.")
   RaiseEvent DownloadNeedRestart(BeFrom)
   Exit Function
  End If
  'Waiting for connect
 Loop While ls <> 12
 'We get header
 h$ = Inet1.GetHeader
 'We get response code from header
 c = GetHttpResponseCode(h$)
 'Oups, response code different from 200 or 206...
 If c <> 206 And c <> 200 Then
 '... we not support this, exitting...
  Download = c
  Exit Function
 End If
' Debug.Print h$

'What length of file?
 cl$ = GetHttpHeaderFieldValue(h$, "content-length")
 TotBytes = Val(cl$)
 DownloadTimeOut = TimeOut
 TDownloadTimeOut = DownloadTimeOut + Timer
 downstart = True
 Timer1.Enabled = True
 RaiseEvent StateChanged("Downloading...")
 'Start download
End Function

Private Sub Inet1_StateChanged(ByVal state As Integer)
 ls = state
End Sub

Private Function GetHttpResponseCode(strHttpHeader) As Integer
    '
    Dim varCode As Variant
    '
    varCode = Mid(strHttpHeader, InStr(1, strHttpHeader, " ") + 1, 3)
    '
    If IsNumeric(varCode) Then
        '
        GetHttpResponseCode = CInt(varCode)
        '
    End If
    '
End Function

Private Function GetHttpHeaderFieldValue(strHttpHeader As String, strHttpHeaderField As String) As String
    '
    Dim strBuffer As String
    Dim intStart As Integer
    Dim strSearchString As String
    '
    If strHttpHeader = "" Then Exit Function
    strSearchString = vbCrLf & strHttpHeaderField & ": "
    intStart = InStr(1, LCase$(strHttpHeader), LCase$(strSearchString))
    If intStart > 0 Then
     intStart = intStart + Len(strSearchString)
     strBuffer = Mid$(strHttpHeader, intStart, InStr(intStart, strHttpHeader, vbCrLf) - intStart)
    '
     If Len(strBuffer) > 0 Then
        GetHttpHeaderFieldValue = strBuffer
     End If
    End If
    '
End Function
Private Sub Timer1_Timer()
 Dim ic As String
 Dim temp As String
 Dim b() As Byte
 Dim q As Integer
 Dim cps As Long
 If downstart = False Then Exit Sub
 Do
  DoEvents
 Loop While Inet1.StillExecuting = True
  'We get the chunk
   b() = Inet1.GetChunk(DownChunkSize, 1)
  ic = ""
  'Transforming to string...
  For q = LBound(b) To UBound(b)
   ic = ic + Chr$(b(q))
  Next q
  CurByte = CurByte + Len(ic)
  'TimeOut?
  If TDownloadTimeOut < Timer Then
   Timer1.Enabled = False
   downstart = False
   TDownloadTimeOut = Timer + DownloadTimeOut
   RaiseEvent DownloadNeedRestart(BeFrom)
   Exit Sub
  End If
  TDownloadTimeOut = DownloadTimeOut + Timer
  If SekElapsed > 0 Then cps = CLng(CurByte / SekElapsed)
  If (CurByte >= TotBytes And TotBytes > 0) Or ic = "" Then
   downstart = False
   Timer1.Enabled = False
   Timer2.Enabled = False
    If CurByte >= TotBytes Then
     RaiseEvent DownloadProgress(100, CurByte, TotBytes, ic, cps)
     ShowProc 100
     RaiseEvent StateChanged("Download complete.")
     RaiseEvent DownloadComplete(BeFrom)
    Else
     RaiseEvent DownloadNeedRestart(BeFrom + (CurByte - 1))
    End If
    Exit Sub
  End If
    If TotBytes > 0 Then
      Procent = Int((CurByte / TotBytes) * 100)
       ShowProc Procent
       temp = ic
       ic = ""
        RaiseEvent DownloadProgress(Procent, CurByte, TotBytes, temp, cps)
     Else
        temp = ic
        ic = ""
        ShowProc Str$(CurByte) + " bytes"
        RaiseEvent DownloadProgress(0, CurByte, TotBytes, temp, cps)
     End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get url() As Variant
    url = m_Url
End Property

Public Property Let url(ByVal New_Url As Variant)
    m_Url = New_Url
    PropertyChanged "Url"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get port() As Variant
    port = m_Port
End Property

Public Property Let port(ByVal New_Port As Variant)
    m_Port = New_Port
    PropertyChanged "Port"
End Property

Private Sub Timer2_Timer()
 SekElapsed = SekElapsed + 1
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Url = m_def_Url
    m_Port = m_def_Port
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Url = PropBag.ReadProperty("Url", m_def_Url)
    m_Port = PropBag.ReadProperty("Port", m_def_Port)
End Sub

Private Sub UserControl_Resize()
 ShowProc 0
End Sub

Sub ShowProc(pr)
 Shape1.Left = 0
 Shape1.Top = 0
 Shape1.Height = UserControl.Height
 If IsNumeric(pr) Then
  Shape1.Width = (UserControl.Width / 100) * pr
  Label1.Caption = pr & "%"
 Else
  Label1.Caption = pr
 End If
 Label1.Left = (UserControl.Width / 2) - (Label1.Width / 2)
 Label1.Top = (UserControl.Height / 2) - (Label1.Height / 2)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Url", m_Url, m_def_Url)
    Call PropBag.WriteProperty("Port", m_Port, m_def_Port)
End Sub

  �  ,   C U S T O M   ��                   Type=Exe
Form=Form1.frm
Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\..\..\WINDOWS\SYSTEM\StdOle2.Tlb#OLE Automation
Object={48E59290-9880-11CF-9754-00AA00C00908}#1.0#0; MSINET.OCX
UserControl=inetdown.ctl
IconForm="Form1"
Startup="Form1"
ExeName32="Project1.exe"
Command32=""
Name="Project1"
HelpContextID="0"
CompatibleMode="0"
MajorVer=1
MinorVer=0
RevisionVer=0
AutoIncrementVer=0
ServerSupportFiles=0
VersionCompanyName="LeSaMiro Company"
CompilationType=0
OptimizationType=0
FavorPentiumPro(tm)=0
CodeViewDebugInfo=0
NoAliasing=0
BoundsCheck=0
OverflowCheck=0
FlPointCheck=0
FDIVCheck=0
UnroundedFP=0
StartMode=0
Unattended=0
Retained=0
ThreadPerObject=0
MaxNumberOfThreads=1
 p   ,   C U S T O M   ��                   Form1 = 44, 44, 371, 365, Z, 22, 22, 349, 343, C
LSInetDownload = 132, 132, 459, 453, , 110, 110, 437, 431, C
