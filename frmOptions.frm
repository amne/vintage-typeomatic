VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Optiuni"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Anulare"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame frSound 
      Caption         =   "Alte optiuni"
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   3855
      Begin VB.TextBox txtTypeTime 
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton opSpeed 
         Caption         =   "Caractere pe minut"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   10
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton opSpeed 
         Caption         =   "Cuvinte pe minut"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "Metronom"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblInfo1 
         Caption         =   "Timp de tastare:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame frDifficulty 
      Caption         =   "Nivel de dificultate"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.OptionButton opDifficulty 
         Caption         =   "Expert"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   2775
      End
      Begin VB.OptionButton opDifficulty 
         Caption         =   "Avansat"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   2775
      End
      Begin VB.OptionButton opDifficulty 
         Caption         =   "Incepator"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Label lblAbout 
      Caption         =   "Despre: TypeOmatic v1.0 - soft creat de Cruceru Cornel Cristian, elev XII G, Colegiul National ""Elena Cuza"", Craiova, Dolj"
      Height          =   1215
      Left            =   4320
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim i&
If Not IsNumeric(txtTypeTime.Text) Then
 MsgBox "Timpul introdus nu este valid!", vbInformation
 Exit Sub
 End If
For i = 0 To 2
 If opDifficulty(i).Value Then Setari.Dificultate = i + 1
 Next i
If opSpeed(0).Value Then Setari.TipViteza = 1
If opSpeed(1).Value Then Setari.TipViteza = 2
If chkSound.Value = 1 Then
 Setari.bSunet = True
 Else
 Setari.bSunet = False
 End If
If txtTypeTime.Text <> "" Then Setari.TimpTastare = txtTypeTime.Text
SalveazaSetari
CitesteText
Unload Me
End Sub

Private Sub Form_Load()
Dim i&
For i = 0 To 2
 If i = Setari.Dificultate - 1 Then opDifficulty(i).Value = True Else opDifficulty(i).Value = False
 Next i
opSpeed(0) = (Setari.TipViteza = swpm)
opSpeed(1) = (Setari.TipViteza = scpm)
txtTypeTime.Text = Setari.TimpTastare
If Setari.bSunet Then chkSound.Value = 1 Else chkSound.Value = 0
End Sub


