VERSION 5.00
Begin VB.Form LedenVenster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leden"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   2550
   ClipControls    =   0   'False
   Icon            =   "LidBox.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   7.063
   ScaleMode       =   4  'Character
   ScaleWidth      =   21.25
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton KlachtKnop 
      Caption         =   "&Klacht"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox LidnaamVeld 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.ComboBox LidnummerVeld 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton SluitenKnop 
      Cancel          =   -1  'True
      Caption         =   "&Sluiten"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label LidnaamLabel 
      Caption         =   "Lidnaam:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.Label LidnummerLabel 
      Caption         =   "Lidnummer:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin VB.Menu HulpMenu 
      Caption         =   "&Hulp"
   End
   Begin VB.Menu InformatieMenu 
      Caption         =   "&Informatie"
   End
End
Attribute VB_Name = "LedenVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze procedure bevat het ledenvenster.
Option Explicit
'Deze procedure werkt de ledenlijst bij.
Private Sub WerkLidnummerlijstBij()
On Error GoTo Fout
Dim Lid As Long
   Screen.MousePointer = vbHourglass
   
   LidnummerVeld.Clear
   For Lid = LBound(Klacht()) To UBound(Klacht())
      If Lidnaam(Lid) = vbNullString Then
         Klacht(Lid) = vbNullString
         If GeselecteerdLidnummer = Lid Then GeselecteerdLidnummer = GEEN_LID
      Else
         LidnummerVeld.AddItem Lid + 1
      End If
   Next Lid
   
   Screen.MousePointer = vbDefault
   
   If GeselecteerdLidnummer = GEEN_LID Then
      If LidnummerVeld.ListCount > 0 Then
         GeselecteerdLidnummer = Val(LidnummerVeld.List(0) - 1)
      Else
         GeselecteerdLidnummer = GEEN_LID
      End If
   End If
   
   If GeselecteerdLidnummer = GEEN_LID Then
      LidnummerVeld.Text = vbNullString
      LidnaamVeld.Text = vbNullString
   Else
      LidnummerVeld.Text = CStr(GeselecteerdLidnummer + 1)
      LidnaamVeld.Text = Lidnaam(GeselecteerdLidnummer)
   End If
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure stelt dit venster in.
Private Sub Form_Load()
On Error GoTo Fout
   Screen.MousePointer = vbHourglass
      
   GeselecteerdLidnummer = GEEN_LID
   
   LaadLeden
   LaadKlachten
   WerkLidnummerlijstBij
   
EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure sluit dit venster.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Fout
   Screen.MousePointer = vbHourglass
   
   BewaarLeden
   BewaarKlachten

EindeProcedure:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure toont de hulp voor dit programma.
Private Sub HulpMenu_Click()
On Error GoTo Fout
Dim HulpTekst As String
   
   HulpTekst = "Leden:" & vbCr
   HulpTekst = HulpTekst & String$(70, "-") & vbCr
   HulpTekst = HulpTekst & "Om een lid in te voeren:" & vbCr
   HulpTekst = HulpTekst & "Voer eerst het lid nummer in en" & vbCr
   HulpTekst = HulpTekst & "voer dan de naam van het lid in." & vbCr
   HulpTekst = HulpTekst & vbCr
   HulpTekst = HulpTekst & "Om een lid te verwijderen:" & vbCr
   HulpTekst = HulpTekst & "Verwijder eerst de naam van het lid en" & vbCr
   HulpTekst = HulpTekst & "verwijder dan de naam van het lid." & vbCr
   HulpTekst = HulpTekst & vbCr
   HulpTekst = HulpTekst & "Klachten:" & vbCr
   HulpTekst = HulpTekst & String$(70, "-") & vbCr
   HulpTekst = HulpTekst & "Selecteer eerst een lid en klik dan" & vbCr
   HulpTekst = HulpTekst & "op de knop Klacht om een klacht" & vbCr
   HulpTekst = HulpTekst & "in te voeren."
   
   MsgBox HulpTekst, vbInformation, App.Title & " - Hulp"

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Dit programma geeft de opdracht om de programmainformatie te tonen.
Private Sub InformatieMenu_Click()
On Error GoTo Fout
   ToonProgrammainformatie
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure toont het klachten venster.
Private Sub KlachtKnop_Click()
On Error GoTo Fout
   KlachtenVenster.Show
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde lidnaam vast.
Private Sub LidnaamVeld_LostFocus()
On Error GoTo Fout
   If Not GeselecteerdLidnummer = GEEN_LID Then
      Lidnaam(GeselecteerdLidnummer) = LidnaamVeld.Text
      WerkLidnummerlijstBij
   End If
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub



'Deze procedure vraagt de ingevoerde lidnaam op.
Private Sub LidnummerVeld_Click()
On Error GoTo Fout
   GeselecteerdLidnummer = Val(LidnummerVeld.List(LidnummerVeld.ListIndex)) - 1
   LidnaamVeld.Text = Lidnaam(GeselecteerdLidnummer)
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde gegevens van een lid vast.
Private Sub LidnummerVeld_LostFocus()
On Error GoTo Fout
   If Not LidnummerVeld.Text = vbNullString Then
      If IsGeldigLidnummer(LidnummerVeld.Text) Then
         GeselecteerdLidnummer = Val(LidnummerVeld.Text) - 1
         LidnaamVeld.Text = Lidnaam(GeselecteerdLidnummer)
      Else
         LidnummerVeld.SetFocus
      End If
   End If
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure sluit dit venster.
Private Sub SluitenKnop_Click()
On Error GoTo Fout
   Unload KlachtenVenster
   Unload LedenVenster
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub



