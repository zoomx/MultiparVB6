Attribute VB_Name = "Funzioni"
Option Explicit

Public Function InputComTimeOut(TimeOut As Integer) As String
'Attende un input il cui terminatore e' vbLF
'Con TIMEOUT
    Dim TimeStop As Long
    Dim Linea As String
    Dim Dummy As String
    Dim iShelll As New CShell
        
        'set ishell=
        TimeOut = TimeOut * 1000 'passiamo ai millisecondi
        
        TimeStop = iShelll.GetTimeMS + TimeOut
        'Debug.Print iShelll.GetTimeMS
        fMain.AFSerial1.InputLen = 1
        Do
            DoEvents

        Loop Until (fMain.AFSerial1.InBufferCount >= 1) Or (iShelll.GetTimeMS > TimeStop)
        If fMain.AFSerial1.InBufferCount >= 1 Then
            Linea = ""
            Dummy = ""
            TimeStop = iShelll.GetTimeMS + TimeOut ' Imposta l'ora di fine
            Do Until Dummy = vbLf Or (iShelll.GetTimeMS > TimeStop)
                DoEvents
                Dummy = fMain.AFSerial1.Input
                Linea = Linea + Dummy
            Loop
        Else
            Linea = "TimeOut"
        End If
        InputComTimeOut = Linea

End Function

Public Sub OpenCom()
    'Apre la porta com
    'Se e' andata bene ComOk e' True altrimenti e' False
    Dim Msg As String

    On Error GoTo ErroreCom
    ComOk = False
    'Apre la porta seriale se non è già aperta
    If fMain.AFSerial1.PortOpen = False Then fMain.AFSerial1.PortOpen = True
    ComOk = True
    Exit Sub
ErroreCom:
    Select Case Err.Number
        Case 8005  'La Com è già aperta
            Msg = "Errore la porta Com" + Str$(ComPort) + " è già in uso"
            #If APPFORGE Then
                MsgBox Msg, vbOKOnly
            #Else
                MsgBox Msg, vbOKOnly, "Errore"
            #End If

            
            Err.Clear   ' Cancella i campi dell'oggetto
            ComOk = False
            Exit Sub
        Case 8002
            Msg = "Errore la porta Com" + Str$(ComPort) + " non esiste!"
            #If APPFORGE Then
                MsgBox Msg, vbOKOnly
            #Else
                MsgBox Msg, vbOKOnly, "Errore"
            #End If

            Err.Clear   ' Cancella i campi dell'oggetto
            ComOk = False
            Exit Sub
        Case Else
            Msg = Err.Description
            #If APPFORGE Then
                MsgBox Msg, vbOKOnly
            #Else
                MsgBox Msg, vbOKOnly, "Errore"
            #End If


            Exit Sub
    End Select

End Sub

Public Sub CloseCom()
    'Chiude la porta seriale se non è già chiusa
    fMain.AFSerial1.InBufferCount = 0
    If fMain.AFSerial1.PortOpen = True Then fMain.AFSerial1.PortOpen = False
End Sub
Public Function Val2(Valore As String) As Single
'Simile alla val ma per separatore decimale usa sia il
'punto che la virgola
    Dim ip As Integer
    Dim iv As Integer
    Dim lStringa As Integer
    Dim Temp As Single
    Dim Stringa As String
    
    Stringa = CStr(Valore)
    'C'è il punto?
    ip = InStr(Stringa, ".")
    'C'è la virgola?
    iv = InStr(Stringa, ",")
    lStringa = Len(Stringa)
    If iv <> 0 Then 'Se c'è la virgola la sostituisce col punto
        Stringa = Left(Stringa, iv - 1) + "." + Right(Stringa, lStringa - iv)
        ip = iv
    End If
    Temp = CSng(Stringa)
    'If ip <> 0 And iv <> 0 Then
    'Se ci sono tutte e due?
    Val2 = Temp
End Function

