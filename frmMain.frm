VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fereastra principala"
   ClientHeight    =   7290
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   6480
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   3795
      TabIndex        =   12
      Top             =   3360
      Width           =   3855
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "START!"
      Height          =   855
      Left            =   9120
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtbType 
      Height          =   3615
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6376
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":30042
   End
   Begin RichTextLib.RichTextBox rtbRead 
      Height          =   3375
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5953
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":300C4
   End
   Begin VB.Frame frInfo 
      Caption         =   "Informatii"
      Height          =   2175
      Left            =   6360
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.Label lblTime 
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblinfo1 
         Alignment       =   1  'Right Justify
         Caption         =   "Numar cuvinte:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblinfo2 
         Alignment       =   1  'Right Justify
         Caption         =   "Numar caractere:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblinfo3 
         Alignment       =   1  'Right Justify
         Caption         =   "Timp:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblChars 
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblWords 
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblinfo4 
         Alignment       =   1  'Right Justify
         Caption         =   "Viteza:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblSpeed 
         Height          =   255
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Menu mainmenu 
      Caption         =   "Meniul principal"
      Begin VB.Menu mmnu_options 
         Caption         =   "Optiuni"
      End
      Begin VB.Menu mmnu_exit 
         Caption         =   "Iesire"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGo_Click()
ResetData
rtbType.Text = ""
bTyping = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
bRunning = False
End Sub

Private Sub mmnu_exit_Click()
Unload Me
End Sub

Private Sub mmnu_options_Click()
Load frmOptions
frmOptions.Show 1
End Sub

Private Sub rtbType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Then KeyCode = 0
End Sub

Private Sub rtbType_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
 nChars = nChars + 1
 If (KeyAscii = 32) Or (Chr(KeyAscii) = vbCrLf) Or (Chr(KeyAscii) = vbCr) Or (Chr(KeyAscii) = vbLf) Or (Chr(KeyAscii) = "'") Or (KeyAscii = 34) Then
  nWords = nWords + 1
  End If
 Else
 KeyCode = 0
 End If
End Sub

Private Sub rtbType_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Then KeyCode = 0
End Sub
