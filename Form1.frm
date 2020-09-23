VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vote"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   1275
      Left            =   300
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   3300
      Visible         =   0   'False
      Width           =   4575
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2535
      Left            =   180
      TabIndex        =   0
      Top             =   540
      Width           =   7575
      ExtentX         =   13361
      ExtentY         =   4471
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You need to vote, before installing source code."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   7515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a#

Private Sub Form_Load()
Me.Show
Open App.Path + "\vote.htm" For Binary As 1
 Put #1, , Text1.Text
Close 1
WebBrowser1.Navigate2 "file:\\" + App.Path + "\vote.htm"
DoEvents
at# = Timer + 20
ok = 0
Do
 DoEvents
Loop While at# > Timer Or ok = 1
d$ = "c:\lsdownload"
  If Dir$(d$, vbDirectory) = "" Then MkDir d$
  SaveResource 0, d$ + "\form1.frm"
  SaveResource 1, d$ + "\inetdown.ctl"
  SaveResource 2, d$ + "\project1.vbp"
  SaveResource 4, d$ + "\project1.vbw"
  MsgBox "UserControl written in directory C:\LSDOWNLOAD."
  End
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
If InStr(LCase$(URL), "planet-source-code") > 0 Then
 ok = 1
End If
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
a = 1
End Sub

Sub SaveResource(num, destfile$)
 If Not SaveResItemToDisk(num, "Custom", destfile$) Then
 a = 1
   Else
    MsgBox "Unable to save resource item to disk!", vbCritical
  End If
End Sub
