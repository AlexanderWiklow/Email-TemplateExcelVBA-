Attribute VB_Name = "Module1"
Option Explicit

Sub SkickaMail()
    Dim appOutlook As Outlook.Application
    Dim MailOutlook As Outlook.MailItem
    Dim CellV, Adresser As String
    
    'Samla ihop epostadresser som �r markerade, b�rja i H2 och stega ner�t.
    'Bygg upp en textstr�ng med alla markerade e-postadresser
    For Each CellV In Range("H2:H50000")
        If CellV.Value <> "" Then
            If CellV.Offset(0, 1) <> "" Then
                Adresser = Adresser & ";" & CellV.Value
            End If
        Else
            Exit For 'n�r man tr�ffar p� tomt
        End If
    Next CellV
    
    'MsgBox Adresser Anv�nds inte nur
    'Ta bort inledande semikolon fr�n adresstr�ngen --
    
    Adresser = Right(Adresser, Len(Adresser) - 1)
    
    'Starta Outlook, skapar en ny instans --
    Set appOutlook = CreateObject("Outlook.Application")
    
    'Skapa ett nytt mail i den aktuella instansen(kopia av en mall) av outlook --
    Set MailOutlook = appOutlook.CreateItem(olMailItem)
    
    'Bygg upp epostmeddelandet --
    MailOutlook.To = Adresser
    MailOutlook.CC = "chefen@jobbet.se"
    MailOutlook.Subject = "Personalfest!"
    MailOutlook.Body = "Hej" & Chr(10) & _
                        "Personalfest p� g�ng." & Chr(10) & Chr(10) & _
                        "MVH" & Chr(10) & Chr(10) & _
                        "Alex"
    MailOutlook.Attachments.Add ("C:\Users\awikl\Downloads\florian-olivo-4hbJ-eymZ1o-unsplash.jpg")
    MailOutlook.Display
    MailOutlook.Send
    
    'OBJ! St�ng objektvariabler --
    Set MailOutlook = Nothing
    Set appOutlook = Nothing
End Sub

