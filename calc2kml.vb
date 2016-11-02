REM  *****  BASIC  *****

Sub Main

Dim FileNo As Integer
Dim CurrentLine As String
Dim Filename As String

Dim Doc As Object
Dim Sheet As Object
Dim Cell As Object
 

Dim Testo as String
Dim LF as String
Dim Row_idx as Integer
Dim Col_idx as Integer
Dim Cell_content as String
Dim LineString_label as String
Dim LineString_SoFar as String

LF = Chr(10)
LineString_label = ""
LineString_SoFar = ""

' ########## apro i file e inizio a scrivere gli header  #########

Doc = ThisComponent				'il documento Calc
Sheet = Doc.Sheets (0)			'il primo foglio del documento Calc

Cell = Sheet.getCellByPosition(1, 0)    ' nella cella B1 c'è il path per il file di output
 
Filename = Cell.String & "textfile_out.kml"  			' Define file name 
FileNo = FreeFile           						    ' Establish free file handle
 
Open Filename For Output As #FileNo         ' Open file (writing mode)


Testo = "<?xml version='1.0' encoding='UTF-8'?>  "&LF 			& _
		"<kml xmlns='http://www.opengis.net/kml/2.2'>  "&LF    & _
		"	<Document> " &LF                                   & _
		"		<name>Dark fiber BPV</name>      "&LF          & _
		"		<description><![CDATA[]]></description>  "&LF  & _
		"		<Folder> " &LF								   &_
		"			<name>Dark fiber BPV</name>      "

Print #FileNo, Testo 

' ########## inizio a ciclare sui dati del file  #########
Row_idx = 2			'inizio dalla terza riga
Col_idx = 0

Cell = Sheet.getCellByPosition(Col_idx, Row_idx)


Do While Cell.type <> com.sun.star.table.CellContentType.EMPTY    ' ciclo sulle righe fino a che non trovo una riga con la prima cella vuota
	
	Select Case Cell.String
	Case "Point"
		' il template è il seguente:
		'	<Placemark>
		'		<name>Via Antonio Meucci, 5</name>
		'		<styleUrl>#icon-1899-0288D1-nodesc</styleUrl>
		'		<Point>
		'			<coordinates>10.973436300000003,45.40517589999999,0.0</coordinates>
		'		</Point>
		'	</Placemark>
		
		Testo = "			<Placemark>"&LF
		
		Cell = Sheet.getCellByPosition(1, Row_idx)
		Testo = Testo & "				<name>" & Cell.String & "</name>" &LF
		
		Cell = Sheet.getCellByPosition(2, Row_idx)
		Testo = Testo & "				<styleUrl>" & Cell.String & "</styleUrl>" &LF
		Testo = Testo & "				<Point>" &LF
		
		Cell = Sheet.getCellByPosition(3, Row_idx)
		Testo = Testo & "					<coordinates>" &  Cell.String & ","
		
		Cell = Sheet.getCellByPosition(4, Row_idx)
		Testo = Testo &  Cell.String & ","
		
		Cell = Sheet.getCellByPosition(5, Row_idx)
		Testo = Testo &  Cell.String & "</coordinates>" &LF
		Testo = Testo & "				</Point>" &LF
		Testo = Testo & "			</Placemark>"
		
		Print #FileNo, Testo 
		
		
	Case "Linestring"
		' il template è il seguente:
		'<Placemark>
		'		<name>Line 4</name>
		'		<styleUrl>#line-000000-1-nodesc</styleUrl>
		'		<LineString>
		'			<tessellate>1</tessellate>
		'			<coordinates>10.9781368,45.4213045,0.0 10.9962647,45.4385209,0.0 11.0464198,45.4338913,0.0</coordinates>
		'		</LineString>
		'	</Placemark>
		
		Cell = Sheet.getCellByPosition(1, Row_idx)
		
		If StrComp(Cell.String , LineString_label, 1) <> 0  Then ' se la etichetta di Linestring è diversa da quella precedente
			' in questo caso ho due possibilità: o sono alla prima istanza di Linestring oppure ho cambiato etichetta
			
			If StrComp("" , LineString_label, 1) <> 0 Then   ' la etichetta precedente non è vuota, quindi ho cambiato etichetta
				' Chiudo la linea e ne stampo tutto il contenuto della linea
				
				LineString_SoFar = LineString_SoFar & "</coordinates>" &LF
				LineString_SoFar = LineString_SoFar & "				</LineString>" &LF
				LineString_SoFar = LineString_SoFar & "			</Placemark>"	
				
				Print #FileNo, LineString_SoFar 	
				
				' ... ma devo anche iniziare a leggere la nuova stringa	
				End If
		
				' o sono alla prima istanza (ho saltato l'IF precedente) oppure ho cambiato etichetta (passando dall'IF precedente)
				' in ogni caso devo leggere la nuova istanza
				
				LineString_label = Cell.String	' imposto la nuova (o prima) label
				
				LineString_SoFar = "			<Placemark>"&LF
				LineString_SoFar = LineString_SoFar & "				<name>" & Cell.String & "</name>" &LF
				
				Cell = Sheet.getCellByPosition(2, Row_idx)
				LineString_SoFar = LineString_SoFar & "				<styleUrl>" & Cell.String & "</styleUrl>" &LF
				LineString_SoFar = LineString_SoFar & "				<LineString>" &LF
				LineString_SoFar = LineString_SoFar & "					<tessellate>1</tessellate>" &LF
				
				Cell = Sheet.getCellByPosition(3, Row_idx)
				LineString_SoFar = LineString_SoFar & "					<coordinates>" & Cell.String & ","
		
				Cell = Sheet.getCellByPosition(4, Row_idx)
				LineString_SoFar = LineString_SoFar & Cell.String & ","
		
				Cell = Sheet.getCellByPosition(5, Row_idx)
				LineString_SoFar = LineString_SoFar & Cell.String
		
		
		Else   ' se la etichetta di Linestring è uguale a quella precedente
			'continuo a comporre la Linestring precedente
			
				Cell = Sheet.getCellByPosition(3, Row_idx)
				LineString_SoFar = LineString_SoFar & "	" & Cell.String & ","
		
				Cell = Sheet.getCellByPosition(4, Row_idx)
				LineString_SoFar = LineString_SoFar & Cell.String & ","
		
				Cell = Sheet.getCellByPosition(5, Row_idx)
				LineString_SoFar = LineString_SoFar & Cell.String
		
		
		End If
		

		
	End Select
		
	' avanzo di una riga e mi riporto all'inizio
	Row_idx = Row_idx + 1
	Col_idx = 0
	Cell = Sheet.getCellByPosition(Col_idx, Row_idx)
Loop   ' il loop sulle righe del foglio

' controllo se sono rimase linestring da finire di stampare

If Len (LineString_SoFar) >0 Then    ' c'è una Linestring ancora non stampata

	LineString_SoFar = LineString_SoFar & "</coordinates>" &LF
	LineString_SoFar = LineString_SoFar & "				</LineString>" &LF
	LineString_SoFar = LineString_SoFar & "			</Placemark>"	
				
	Print #FileNo, LineString_SoFar

End If

Testo = "		</Folder> " &LF  & _
		"	</Document> " &LF    & _
		"</kml>" 
Print #FileNo, Testo

Close #FileNo                 					 ' Close file

MsgBox "Fatto!"

End Sub
