# VBA
VBA - (Visual Basics for Aplications) is język programowania stworzony przez Microsoft, używany głównie do automatyzacji zadań w aplikacjach pakietu Office. 

VBA
mamy 2 rodzaje procedur:

Procedura Sub oraz Function (Obie w Public/Private)
Sub - Nie zwraca nic 
Function - Cos tam zwraca

BTW MODUŁ TO TAKI KURWA SKRYPCIOR 
By dodać moduł należy wybrać z paska menu Insert -> Module 
Pole Name w sekcji Properties i wpisz nową nazwę dla modułu 

Dyrektywa ,,Option Compare Database"
Każdy nowo utworzony moduł zaczyna się od tej linii (jak np. <!Doctype HTML> w htmlu)

Option compare Database ma 3 opcje:
Binary - porównywanie łańcuchów operia się na binarnej reprezentacji znaków

Text - oparta jest an lokalnym ustawieniu systemu instrukcja określające, że porównanie ciągów znaków ma być zgodne z wielkością liter ("a" -> "A")
Database - używana na poziomie modułu do wymuszenia jawnej deklaracji wszystkich zmiennych w danym module. Użyj Option Explicit, aby uniknąć nieprawidłowego wpisania nazwy istniejącej zmiennej lub pomyłki w kodzie, gdy zakres zmiennej jest niejasny.


------------------
  OKNO IMMEDIATE
------------------
Okno immediate umożliwia wykonywanie obliczeń w trybie interaktywnym: 
NA PRZYKŁAD: 
wpisanie "? 6*6" wzróci 36.
Oprócz tego masz specjalne zapytania 
? date - dzisiejsza data
? time - obecny czas
? now - data i czas

Komentarze
"'(jebać pindla)"


TYPY DANYCH:
Byte - mała liczba (0-255)
Integer - Liczba (-32768 do 32767)
Long - W chuj długa liczba
Single - Liczba zmiennoprzecinkowa
Double - Liczba zmiennoprzecinkowa ale długa 
Currency - pieniążki

String - Napis
Date - data 

Logiczne:
Boolean - Prawda/Fałsz
Object - Obiekt w bazie 
Variant - Dowolny Typ


PO CO WYBRAĆ TYP:
Optymalizacja pamięci, Wydajność Kodu, Zapobieganie błędom

DAO - (Data Access Objects) - dzięki temu gówno możesz brać dane z plików .mdb i .accdb, oraz Microsoft Jet jak i ACE. Dao to zestaw obiektów dzięki którym manipulujemy danymi.

ADO - (ActiveX Data Objects) - Nowsze kurwa ale z każdom bazą (SQL, Oracle)


Tworzenie nowej procedury:
Aby utworzyć procedurę (nowy moduł) należy: Insert Module (nazwa nie może się powtarzać)

------------------
  KOD
------------------

Dim - Deklarowanie Zmiennych np. (Dim [Nazwa] As [Typ Danych] np. Dim zmienna As Integer, później dodajemy zmienna = coś)
MsgBox - wyświetlanie czegoś elo belo 

if ____ then: (kurwa if no każdy wie co to if) 

Przykładowy Kod:

Option Compare Database

Sub P()

End Sub

Sub Main()
	Dim a As Integer
	a = 67
	'(jebać pindla)
	MsgBox(a)
End Sub
