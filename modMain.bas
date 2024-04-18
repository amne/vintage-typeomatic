Attribute VB_Name = "modMain"
Option Explicit

Public Enum ESpeedType
 swpm = 1
 scpm = 2
 End Enum

Public Enum EDifficulty
 dincepator = 1
 davansat = 2
 dexpert = 3
 End Enum

Public Type TWordList
 numWords As Long
 aWords() As String
 End Type

Public Type TSetari
 Dificultate As EDifficulty
 TipViteza As ESpeedType
 TimpTastare As Long
 bSunet As Boolean
 End Type

Public Type TStats
 nSpeed As Long 'in cuvinte pe minute sau caractere pe minute dupa caz
 nWordsChars As Long 'numarul de cuvinte/caractere la fel dupa caz
 nAccuracy As Long 'procentul de acuratete
 nTime As Long
 End Type

Public Setari As TSetari
Public Statistici As TStats

Public bRunning As Boolean
Public nChars As Long
Public nWords As Long
Public nTime As Long
Public nSpeed As Long
Public bTyping As Boolean

Public Sub main()
Dim nSec1&, nSec2&
frmMain.Show
bRunning = True
CitesteSetari
CitesteText
ResetData
nSec1 = Second(Now)
While bRunning
 If bTyping Then
  If frmMain.rtbType.Text <> "" Then
   If Second(Now) <> nSec1 Then
    nSec2 = nSec2 + 1
    nSec1 = Second(Now)
    If Setari.bSunet Then Beep
    End If
   End If
  End If
 frmMain.lblTime.Caption = nTime - nSec2
 Select Case Setari.TipViteza
  Case scpm:
   If nSec2 <> 0 Then nSpeed = 60 * nChars / nSec2
   frmMain.lblSpeed = nSpeed & " CarPM"
   Statistici.nWordsChars = nChars
  Case swpm:
   If nSec2 <> 0 Then nSpeed = 60 * nWords / nSec2
   frmMain.lblSpeed = nSpeed & " CuvPM"
   Statistici.nWordsChars = nWords
  End Select
 Statistici.nTime = nSec2
 Statistici.nSpeed = nSpeed
 frmMain.lblWords = nWords
 frmMain.lblChars = nChars
 If Len(frmMain.rtbType.Text) = Len(frmMain.rtbRead.Text) Then nSec2 = nTime
 If bTyping Then
  If nSec2 = nTime Then
   nSec2 = 0
   bTyping = False
   Load frmStats
   frmStats.Show 1
   End If
  End If
 DoEvents
 Wend
End
End Sub

Public Sub CitesteSetari()
Setari.Dificultate = GetSetting(App.Title, "Setari", "Dificultate", 1)
Setari.TipViteza = GetSetting(App.Title, "Setari", "TipViteza", 1)
Setari.bSunet = (GetSetting(App.Title, "Setari", "Sunet", 1) = 2)
Setari.TimpTastare = GetSetting(App.Title, "Setari", "TimpTastare", 60)
End Sub
Public Sub SalveazaSetari()
SaveSetting App.Title, "Setari", "Dificultate", Setari.Dificultate
SaveSetting App.Title, "Setari", "TipViteza", Setari.TipViteza
If Setari.bSunet Then
 SaveSetting App.Title, "Setari", "Sunet", 2
 Else
 SaveSetting App.Title, "Setari", "Sunet", 1
 End If
SaveSetting App.Title, "Setari", "TimpTastare", Setari.TimpTastare
End Sub

Public Sub ResetData()
nChars = 0
nWords = 0
nTime = Setari.TimpTastare
End Sub


Public Sub CitesteText()
Dim tStr$
frmMain.rtbRead.Text = ""
Select Case Setari.Dificultate
  Case dincepator: Open App.Path & "\incepator.txt" For Input As #1
  Case davansat: Open App.Path & "\avansat.txt" For Input As #1
  Case dexpert: Open App.Path & "\expert.txt" For Input As #1
  End Select
While Not EOF(1)
 frmMain.rtbRead.SelStart = Len(frmMain.rtbRead.Text)
 Line Input #1, tStr
 frmMain.rtbRead.SelText = tStr
 frmMain.rtbRead.SelStart = Len(frmMain.rtbRead.Text)
 Wend
Close #1
End Sub


