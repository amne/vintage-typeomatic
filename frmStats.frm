VERSION 5.00
Begin VB.Form frmStats 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Statistici"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   5520
      Width           =   1095
   End
   Begin VB.ListBox lstType 
      Height          =   2985
      Left            =   2640
      MultiSelect     =   1  'Simple
      TabIndex        =   1
      Top             =   2160
      Width           =   1815
   End
   Begin VB.ListBox lstRead 
      Height          =   2985
      Left            =   360
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblStats 
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblStats 
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblStats 
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Acuratete:"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Viteza de tastare:"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Timp:"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblType 
      Caption         =   "Cuvinte tastate:"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblRead 
      Caption         =   "Cuvinte citite:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tWL1 As TWordList
Dim tWL2 As TWordList

Private Sub cmdOK_Click()
Unload Me
End Sub


Private Function ParseText(sTxt$) As TWordList
Dim sTemp1$, sT$(), tempWL As TWordList, i&, j&, k As String
j = 1
sTxt = Trim(sTxt)
For i = 1 To Len(sTxt)
 k = Mid(sTxt, i, 1)
 If (k = Chr(32)) Or (k = vbCrLf) Or (k = vbCr) Or (k = vbLf) Or (k = "'") Or (k = Chr(34)) Then
  tempWL.numWords = tempWL.numWords + 1
  ReDim Preserve tempWL.aWords(tempWL.numWords) As String
  tempWL.aWords(tempWL.numWords) = Mid(sTxt, j, i - j)
  j = i + 1
  End If
 Next i
tempWL.numWords = tempWL.numWords + 1
ReDim Preserve tempWL.aWords(tempWL.numWords) As String
tempWL.aWords(tempWL.numWords) = Mid(sTxt, j, Len(sTxt))
ParseText = tempWL
End Function


Private Sub Form_Load()
Dim i&, iMax&, nWordsGood&
tWL1 = ParseText(frmMain.rtbRead.Text)
tWL2 = ParseText(frmMain.rtbType.Text)
iMax = tWL1.numWords
If tWL1.numWords > tWL2.numWords Then iMax = tWL2.numWords
For i = 1 To iMax
 lstRead.AddItem tWL1.aWords(i)
 lstType.AddItem tWL2.aWords(i)
 If tWL1.aWords(i) = tWL2.aWords(i) Then
  nWordsGood = nWordsGood + 1
  Else
  lstRead.Selected(i - 1) = True
  lstType.Selected(i - 1) = True
  End If
 Next i
lblStats(0).Caption = Statistici.nTime
If Setari.TipViteza = scpm Then lblStats(1).Caption = nSpeed & " CarPM"
If Setari.TipViteza = swpm Then lblStats(1).Caption = nSpeed & " CuvPM"
lblStats(2).Caption = nWordsGood & "/" & iMax
Statistici.nAccuracy = 100 * nWordsGood / iMax
End Sub
