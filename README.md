# VBA
VBA - (Visual Basics for Aplications) is język programowania stworzony przez Microsoft, używany głównie do automatyzacji zadań w aplikacjach pakietu Office. 

VBA <br>
mamy 2 rodzaje procedur:

Procedura Sub oraz Function (Obie w Public/Private) <br>
Sub - Nie zwraca nic  <br>
Function - Cos tam zwraca <br>

BTW MODUŁ TO TAKI KURWA SKRYPCIOR <br>
By dodać moduł należy wybrać z paska menu Insert -> Module <br>
Pole Name w sekcji Properties i wpisz nową nazwę dla modułu  <br>

Dyrektywa ,,Option Compare Database" <br>
Każdy nowo utworzony moduł zaczyna się od tej linii (jak np. <!Doctype HTML> w htmlu) <br>

Option compare Database ma 3 opcje: <br>
Binary - porównywanie łańcuchów operia się na binarnej reprezentacji znaków <br>

Text - oparta jest an lokalnym ustawieniu systemu instrukcja określające, że porównanie ciągów znaków ma być zgodne z wielkością liter ("a" -> "A") <br>
Database - używana na poziomie modułu do wymuszenia jawnej deklaracji wszystkich zmiennych w danym module. Użyj Option  <br>
Explicit, aby uniknąć nieprawidłowego wpisania nazwy istniejącej zmiennej lub pomyłki w kodzie, gdy zakres zmiennej jest niejasny. <br>


------------------ <br>
  OKNO IMMEDIATE <br>
------------------ <br>
Okno immediate umożliwia wykonywanie obliczeń w trybie interaktywnym: <br>
NA PRZYKŁAD: <br>
wpisanie "? 6*6" wzróci 36. <br>
Oprócz tego masz specjalne zapytania  <br>
? date - dzisiejsza data <br>
? time - obecny czas <br>
? now - data i czas <br>

Komentarze <br>
"'(jebać pindla)" <br>


TYPY DANYCH: <br>
Byte - mała liczba (0-255) <br>
Integer - Liczba (-32768 do 32767) <br>
Long - W chuj długa liczba <br>
Single - Liczba zmiennoprzecinkowa <br>
Double - Liczba zmiennoprzecinkowa ale długa <br>
Currency - pieniążki <br>

String - Napis <br>
Date - data  <br>

Logiczne: <br> 
Boolean - Prawda/Fałsz <br>
Object - Obiekt w bazie  <br>
Variant - Dowolny Typ <br>


PO CO WYBRAĆ TYP: <br>
Optymalizacja pamięci, Wydajność Kodu, Zapobieganie błędom <br>

DAO - (Data Access Objects) - dzięki temu gówno możesz brać dane z plików .mdb i .accdb, oraz Microsoft Jet jak i ACE. Dao to zestaw obiektów dzięki którym manipulujemy danymi.
<br>
ADO - (ActiveX Data Objects) - Nowsze kurwa ale z każdom bazą (SQL, Oracle)


Tworzenie nowej procedury: <br>
Aby utworzyć procedurę (nowy moduł) należy: Insert Module (nazwa nie może się powtarzać)

------------------ <br>
  KOD <br>
------------------ <br>

Dim - Deklarowanie Zmiennych np. (Dim [Nazwa] As [Typ Danych] np. Dim zmienna As Integer, później dodajemy zmienna = coś) <br>
MsgBox - wyświetlanie czegoś elo belo <br>

if ____ then: (kurwa if no każdy wie co to if) <br>

Przykładowy Kod: <br>

Option Compare Database <br>

Sub P() <br>

End Sub <br>

Sub Main() <br>
	Dim a As Integer <br>
	a = 67 <br>
	'(jebać pindla) <br>
	MsgBox(a) <br>
End Sub <br>
