VERSION 5.00
Begin VB.Form KlachtenVenster 
   Caption         =   "Klachten:"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   ClipControls    =   0   'False
   Icon            =   "Klachten.frx":0000
   ScaleHeight     =   14.375
   ScaleMode       =   4  'Character
   ScaleWidth      =   40.125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton SluitenKnop 
      Cancel          =   -1  'True
      Caption         =   "&Sluiten"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox KlachtVeld 
      Height          =   2895
      Left            =   0
      MaxLength       =   65535
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "KlachtenVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze procedure bevat het klachten venster.
Option Explicit


'Deze procedure stelt dit venster in.
Private Sub Form_Load()
On Error GoTo Fout
   KlachtenVenster.Width = Screen.Width \ 2
   KlachtenVenster.Height = Screen.Height \ 2
   
   KlachtVeld.Text = Klacht(GeselecteerdLidnummer)
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure past dit venster aan de nieuwe afmetingen aan.
Private Sub Form_Resize()
On Error Resume Next
   KlachtVeld.Width = KlachtenVenster.ScaleWidth
   KlachtVeld.Height = KlachtenVenster.ScaleHeight - 2
   SluitenKnop.Left = (KlachtenVenster.ScaleWidth - 2) - SluitenKnop.Width
   SluitenKnop.Top = KlachtenVenster.ScaleHeight - 2
End Sub











'Deze procedure sluit dit venster.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Fout
   Klacht(GeselecteerdLidnummer) = KlachtVeld.Text
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht dit venster te sluiten.
Private Sub SluitenKnop_Click()
On Error GoTo Fout
   Unload KlachtenVenster
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


