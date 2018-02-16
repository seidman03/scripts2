Sub estymacja()

'
' Makro do automatycznej wysyłki raportu
'

' Deklaracja zmiennych
Dim i As Long
Dim OutApp, OutMail As Object
Dim strto, strcc, strbcc, strsub, strbody As String

' Zapobieganie wyskakiwaniu niepotrzebnych okienek 
Application.DisplayAlerts = False

' Wstrzymanie odświeżania ekranu
Application.ScreenUpdating = False

' Ustalenie ścieżki załącznika
Filename = "C:\Users\User\Desktop\Załącznik.xlsx"


' Aktywacja arkusza, w którym znajdują się adresy mailowe osób, do których wysyłany będzie raport 
ActiveWorkbook.Worksheets("Kontakty").Select

' Stworzenie obiektu aplikacji Outlook
Set OutApp = CreateObject("Outlook.Application")
OutApp.Session.Logon

' Pętla - mail powinien się wysyłać do każdej osoby osobno
For i = 6 To Range("c65536").End(xlUp).Row
Set OutMail = OutApp.createitem(0)

' Adres mailowy odbiorcy pobierany kolejno w każdej iteracji pętli
    strto = Cells(i, 1)
	
' Tytuł maila
    strsub = "Raport cykliczny"

' Treść maila: zdanie wstępu oraz podsumowanie najważniejszych informacji
    strbody = _
        "Poniżej znajduje się raport. W załączniku zestawienie szczegółowe." & vbNewLine & vbNewLine & _
        Cells(i, 10) & vbNewLine & vbNewLine & _
        Cells(1, 2) & ": " & Cells(i, 2) & vbNewLine & Cells(1, 3) & ": " & Cells(i, 3) & vbNewLine & Cells(1, 4) & ": " & Cells(i, 4) & vbNewLine & Cells(1, 5) & ": " & Cells(i, 5) & vbNewLine & Cells(1, 6) & ": " & Cells(i, 6) & _
        vbNewLine & Cells(1, 7) & ": " & Cells(i, 7) & _
        vbNewLine & vbNewLine & _
        "Raport wygenerowany automatycznie." & vbNewLine & vbNewLine

' Wysyłka maila
    With OutMail
        .To = strto
        .Subject = strsub
        .Body = strbody
        .Attachments.Add Filename
        .Send
        '.Display 'Zamiast Send użyć do testowania
    End With
    On Error Resume Next

' Przejście do kolejnej osoby z listy
Next

' Zmiana przypisania
Set OutMail = Nothing
Set OutApp = Nothing


End Sub