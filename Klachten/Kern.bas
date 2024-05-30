Attribute VB_Name = "KernModule"
'Deze module bevat de kern procedures voor dit programma.
Option Explicit
Public Const GEEN_LID As Long = -1   'Definieert "geen lid".

Public GeselecteerdLidnummer As Long   'Bevat het geselecteerde lidnummer.
Public Klacht(0 To 9999) As String     'Bevat de lijst van klachten.
Public Lidnaam(0 To 9999) As String    'Bevat de lijst van lidnamen.
'Deze procedure bewaart de klachten.
Public Sub BewaarKlachten()
On Error GoTo Fout
Dim BestandH As Integer
Dim Lid As Long

   BestandH = FreeFile()
   Open "Klachten.lst" For Output Lock Read Write As BestandH
      For Lid = LBound(Klacht()) To UBound(Klacht())
         Print #BestandH, Converteer16BitsNaarBytes(Len(Klacht(Lid))); Klacht(Lid);
      Next Lid
   Close BestandH
   
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure bewaart de ledenlijst.
Public Sub BewaarLeden()
On Error GoTo Fout
Dim BestandH As Integer
Dim Lid As Long

   BestandH = FreeFile()
   Open "Leden.lst" For Output Lock Read Write As BestandH
      For Lid = LBound(Klacht()) To UBound(Klacht())
         Print #BestandH, Chr$(Len(Lidnaam(Lid))); Lidnaam(Lid);
      Next Lid
   Close BestandH

EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub
'Deze procedure converteert de opgegeven waarde naar een 16-bit byte waarde.
Private Function Converteer16BitsNaarBytes(Waarde As Long) As String
On Error GoTo Fout
Dim LinkerByte As Long
Dim RechterByte As Long

   LinkerByte = Waarde \ &H100
   RechterByte = Waarde And &HFF
EindeProcedure:
   Converteer16BitsNaarBytes = Chr$(LinkerByte) & Chr$(RechterByte)
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function
'Deze procedure converteert de opgegeven bytes naar een 16-bit waarde.
Private Function ConverteerBytesNaar16Bits(Bytes As String) As Long
On Error GoTo Fout
Dim Waarde As Long
   Waarde = (Asc(Left$(Bytes, 1)) * &H100) Or Asc(Right$(Bytes, 1))
EindeProcedure:
   ConverteerBytesNaar16Bits = Waarde
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function


'Deze procedure handelt eventuele fouten af.
Public Function HandelFoutAf(Optional VraagVorigeKeuzeOp As Boolean = False) As Long
Dim Bericht As String
Dim Foutcode As Long
Static Keuze As Long

   Screen.MousePointer = vbDefault

   Bericht = Err.Description
   Foutcode = Err.Number
   Err.Clear
      
   If Not VraagVorigeKeuzeOp Then
      Bericht = Bericht & vbCr & "Foutcode: " & CStr(Foutcode)
   
      Keuze = MsgBox(Bericht, vbAbortRetryIgnore Or vbExclamation)
   End If
   
   HandelFoutAf = Keuze
   
   If Keuze = vbAbort Then End
End Function

'Deze procedure controleert of het opgegeven lidnummer geldig is en stuurt het resultaat terug.
Public Function IsGeldigLidnummer(Lidnummer As String) As Boolean
On Error GoTo Fout
Dim IsGeldig As Boolean

   IsGeldig = False

   If CStr(CLng(Val(Lidnummer))) = Lidnummer Then
      If Val(Lidnummer) < 1 Or Val(Lidnummer) > 10000 Then
         MsgBox "Het lidnummer moet tussen de 1 en 10000 zijn.", vbExclamation
      ElseIf Val(Lidnummer) = 0 Then
         MsgBox "Het lidnummer kan geen nul zijn.", vbExclamation
      Else
         IsGeldig = True
      End If
   End If
   
EindeProcedure:
   IsGeldigLidnummer = IsGeldig
   Exit Function

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Function

'Deze procedure laadt de klachten.
Public Sub LaadKlachten()
On Error GoTo Fout
Dim BestandH As Integer
Dim Lengte As Long
Dim Lid As Long

   BestandH = FreeFile()
   Open "Klachten.lst" For Binary Lock Read Write As BestandH
      If LOF(BestandH) = 0 Then
         Close BestandH
         Kill "Klachten.lst"
      Else
         For Lid = LBound(Klacht()) To UBound(Klacht())
            Lengte = ConverteerBytesNaar16Bits(Input$(2, BestandH)): Klacht(Lid) = Input$(Lengte, BestandH)
         Next Lid
      End If
   Close BestandH
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub


'Deze procedure laadt de ledenlijst.
Public Sub LaadLeden()
On Error GoTo Fout
Dim BestandH As Integer
Dim Lengte As Long
Dim Lid As Long

   BestandH = FreeFile()
   Open "Leden.lst" For Binary Lock Read Write As BestandH
      If LOF(BestandH) = 0 Then
         Close BestandH
         Kill "Leden.lst"
      Else
         For Lid = LBound(Klacht()) To UBound(Klacht())
            Lengte = Asc(Input$(1, BestandH)): Lidnaam(Lid) = Input$(Lengte, BestandH)
         Next Lid
      End If
   Close BestandH
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

'Deze procedure toont de informatie over dit programma.
Public Sub ToonProgrammainformatie()
On Error GoTo Fout
   MsgBox App.Comments, vbInformation, App.Title & " v" & App.Major & "." & App.Minor & App.Revision & ", door: " & App.CompanyName
EindeProcedure:
   Exit Sub

Fout:
   If HandelFoutAf(VraagVorigeKeuzeOp:=False) = vbIgnore Then Resume EindeProcedure
   If HandelFoutAf() = vbRetry Then Resume
End Sub

