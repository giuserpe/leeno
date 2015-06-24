
rem ***** BASIC *****

sub Struttura_ComputoM____OK ' sul computo (ok sulla 3.7)
	'>>>>>>>>>>>>>>>>>>>>>>>>>>> per evitare che duplichi la struttura
	 Togli_Struttura
	 '<<<<<<<<<<<<<<<<<<<<<<<<<<<
	Ripristina_statusLine
	Barra_Apri_Chiudi_5("1 Sto creando le strutture... Pazienta...", 10)
	oSheet = ThisComponent.currentController.activeSheet 'oggetto sheet
	iSheet = oSheet.RangeAddress.sheet ' index della sheet
'	sAttributo = Trova_Attr_Sheet
 If osheet.Name <> "COMPUTO" Then	
 			 msgbox "#5 Questo comando si può usare solo" & CHR$(10)_
 			 &" in una tabella di COMPUTO!", 16, "AVVISO!"
			exit sub
	end if
 	if ThisComponent.currentcontroller.activesheet.name = "COMPUTO" then
		s_R = ""
	end if
	if ThisComponent.currentcontroller.activesheet.name = "CONTABILITA" then
		s_R = "_R"
	end if		
	if right( (oSheet.GetCellByPosition(0 ,5).CellStyle), 2) = "_R" or	_
				right( (oSheet.GetCellByPosition(0 ,6).CellStyle), 2) = "_R" or _
				right( (oSheet.GetCellByPosition(0 ,7).CellStyle), 2) = "_R" or _
				right( (oSheet.GetCellByPosition(0 ,8).CellStyle), 2) = "_R" then
				s_R = "_R"	
	end if	

	
	lStartRow = oSheet.GetCellByPosition( 0 , 0)
'	lLastUrowNN = getLastUsedRow(oSheet)
	
	sString$ = "Fine Computo"
	
	oEnd=uFindString(sString$, oSheet) 
	If isNull (oEnd) or isEmpty (oEnd) then 
						lLastUrowNN = getLastUsedRow(oSheet)-1
					else
						lLastUrowNN=oEnd.RangeAddress.EndRow -1
	end if
	IF lLastUrowNN	>= 600 and lLastUrowNN	<= 1000 then
		if msgbox ("Questa tabella è piuttosto corposa e ci vorrà un po' di tempo! " & CHR$(10)_
		 	& " (alcuni minuti...) " & CHR$(10) & CHR$(10)_
			&" PROSEGUO ?... " & CHR$(10) _
 			& CHR$(10) & "",36, "Creazione STRUTTURE") = 7 then
 			exit sub
 	end if
	end if
	IF lLastUrowNN	> 1001 then
		if msgbox ("Questa tabella è piuttosto corposa e ci vorrà un po' di tempo! " & CHR$(10)_
			& " ( Parecchi minuti...) " & CHR$(10) & CHR$(10)_
			&" PROSEGUO ?... " & CHR$(10) _
 			& CHR$(10) & "",36, "Creazione STRUTTURE") = 7 then
 			exit sub
 	end if
	end if
	'lLastUrowNN=oEnd.RangeAddress.EndRow-1
	ThisComponent.Sheets.getByName("S1").GetCellByPosition(7,280).string = NOW 'debug
	Ripristina_statusLine
	Barra_Apri_Chiudi_5("2 Sto creando le strutture... Pazienta...", 15)
	iLivelliVisibili = ThisComponent.Sheets.getByName("S1").GetCellByPosition(7,297).value 
	if iLivelliVisibili=0 then 'per avere comunque un valore di default nel caso il riferimento sia vuoto
			iLivelliVisibili = 2
	end if
For I = 3 to lLastUrowNN
Dim oCRA As New com.sun.star.table.CellRangeAddress
		'		print left((oSheet.GetCellByPosition(1 , (I)).CellStyle),7) 'debug
		' se è dentro una intestazione di capito o sottoCap
		if	oSheet.GetCellByPosition(1 , (I)).CellStyle = "livello-1-sopra" & s_R or _
		 		oSheet.GetCellByPosition(1 , (I)).CellStyle = "Livello-1-scritta" & s_R or _
		 		oSheet.GetCellByPosition(1 , (I)).CellStyle = "livello-1-sotto_" & s_R or _
			 	oSheet.GetCellByPosition(1 , (I)).CellStyle = "livello2 valuta" & s_R or _
			 	oSheet.GetCellByPosition(1 , (I)).CellStyle = "livello-2-sotto_" & s_R or _
			 	oSheet.GetCellByPosition(1 , (I)).CellStyle = "livello2 sopra" & s_R or _
			 	oSheet.GetCellByPosition(1 , (I)).CellStyle = "Reg_prog" or _
			 	oSheet.GetCellByPosition(1 , (I)).CellStyle = "Default" or _
			 	oSheet.GetCellByPosition(1 , (I)).CellStyle = "COMP_BASE" or _
			 	left((oSheet.GetCellByPosition(1 , (I)).CellStyle),7) = "Reg-SAL" then	 'questa per usare stile gerarchico
	'	ThisComponent.CurrentController.Select(oSheet.GetCellByPosition(1 ,I )) 'debug	
	'	print "cap " & oSheet.GetCellByPosition(1 ,I ).CellStyle
				goto prossima ' passa alla riga dopo
			else 
				iRI=I
				Do while oSheet.GetCellByPosition(1 , (I)).CellStyle = "Comp-Bianche sopra" & s_R OR _
						oSheet.GetCellByPosition(1 , (I)).CellStyle = "comp Art-EP" & s_R OR _
						oSheet.GetCellByPosition(1 , (I)).CellStyle = "Comp-Bianche in mezzo" & s_R OR _
						oSheet.GetCellByPosition(1 , (I)).CellStyle = "comp sotto Bianche" & s_R ' ' prima avevo OR
				'	ThisComponent.CurrentController.Select(oSheet.GetCellByPosition(1 ,I )) 'debug	
				'	print oSheet.GetCellByPosition(1 ,I ).CellStyle
						if oSheet.GetCellByPosition(1 , (I)).CellStyle = "Comp-Bianche in mezzo" & s_R then
							iRIs = I
							Do while oSheet.GetCellByPosition(1 , (I)).CellStyle = "Comp-Bianche in mezzo" & s_R 
								'ThisComponent.CurrentController.Select(oSheet.GetCellByPosition(1 ,I )) 'debug
								'print 'ThisComponent.CurrentController.Select(oSheet.GetCellByPosition(1 ,I )).CellStyle'debug
								I = I + 1 
							loop
							if iRIs > I-1 then
									ThisComponent.CurrentController.Select(oSheet.getCellByPosition(1,iRI))
									beep
									oCRA.EndRow = iRIs 
									
									Select case msgbox ("Da queste parti lo stile delle righe non sembra ordinato! "& CHR$(10)_
											& "Io ho raggruppato ugualmente... ma se la struttura non ti sembrerà ordinata non dare la colpa a me... -;)"_
											&" " & CHR$(10) _
 											& CHR$(10) & "",49, "") 				
										case 6 'SI
										case 7
							 				exit sub
						 	 			case 2
						 					exit sub
 								 	end select	
								else
									oCRA.EndRow = I-1	
							end if
							oCRA.Sheet =iSheet
							oCRA.StartColumn = 	4
							oCRA.StartRow = iRIs
							oCRA.EndColumn = 9
							oSheet.group(oCRA,1,1)
							' magari provare autoOutline (sarà + veloce?)
							oSheet.showLevel(iLivelliVisibili-1,1)
							'oSheet.hideDetail(oCRA)
						end if
						I=I+1
				loop
				if iRI > I-1 then
						ThisComponent.CurrentController.Select(oSheet.getCellByPosition(1,iRI))
						beep
						oCRA.EndRow = iRi
						msgbox "In questo punto lo stile delle righe non sembra ordinato! "& CHR$(10)_
								& "Io ho raggruppato ugualmente... ma se la struttura non ti sembrerà ordinata non dare la colpa a me... -;)"	
					else
						oCRA.EndRow = I-1	
				end if
				oCRA.Sheet =iSheet
				oCRA.StartColumn = 	5
				oCRA.StartRow = iRI
				oCRA.EndColumn = 5
				'oCRA.EndRow = I-1
				oSheet.group(oCRA,1)
			'	oSheet.hideDetail(oCRA)
				oSheet.showLevel(iLivelliVisibili-1,1)
			'	oSheet.showLevel(1,1)

				I = I+1 
		end if
		prossima:
	'	ThisComponent.CurrentController.Select(oSheet.GetCellByPosition(1 ,I )) 'debug	
	'	print oSheet.GetCellByPosition(1 ,I ).CellStyle
	'	if i = 10 then ' and i <= lLastUrowNN / 4 + 1 then
		'	Ripristina_statusLine
		Barra_Apri_Chiudi_5( stxt, 50)
			sTxt = "Sto creando le strutture e sono alla riga " & i
			Barra_Apri_Chiudi_5( stxt, 50)
	'	end if
		goto salta_debug

		Ripristina_statusLine
		Barra_Apri_Chiudi_5(sTxt, lLastUrowNN/5*100/lLastUrowNN)
		if i >= lLastUrowNN / 4 then ' and i <= lLastUrowNN / 4 + 1 then
			Ripristina_statusLine
			Barra_Apri_Chiudi_5(sTxt, lLastUrowNN/4*100/lLastUrowNN)
		end if
		if i >= lLastUrowNN/4*2 then 'and i <= lLastUrowNN/4*2 + 1 then
			Ripristina_statusLine
			Barra_Apri_Chiudi_5(sTxt , lLastUrowNN/4*2*100/lLastUrowNN)
		end if
		if i >= lLastUrowNN/4*3 then 'and i <= lLastUrowNN/4*3 + 1 then
			Ripristina_statusLine
			Barra_Apri_Chiudi_5(sTxt, lLastUrowNN/4*3*100/lLastUrowNN)
		end if

		salta_debug:
next I

	oCRA.Sheet =iSheet
	oCRA.StartColumn = 	4
	oCRA.StartRow = 1
	oCRA.EndColumn = 8
	oCRA.EndRow = 100
	oSheet.group(oCRA,0)
	Ripristina_statusLine
	ThisComponent.Sheets.getByName("S1").GetCellByPosition(7,281).string = NOW 'debug
end sub




FUNCTION Circoscrive_Voce_Computo_Att___ (ByVal lrow As Long) 'individua un record di Computo

'---------------------------------------------------------------------------
							'restituisce il range'
dim lrowS as long
dim lrowE as long
	oSheet = ThisComponent.currentController.activeSheet 
	lcol = 0
	oCell = oSheet.GetCellByPosition( 0 , lrow)
	nCurRow = lrow
	if templateTipo = "ATT" then lcol5 = 3
	if templateTipo = "5C" then lcol5 = 2
	Do while (Trova_Attr_N (oCell, oSheet)) <> "Start_voce_COMPUTO"_
	and ((oSheet.GetCellByPosition( 0, nCurRow).string <> "")_
			or	 (oSheet.GetCellByPosition( 0,nCurRow).string <>"")_
			 or (oSheet.GetCellByPosition( 4,nCurRow).string <> "")_
			 or (oSheet.GetCellByPosition( 5,nCurRow).string <> "")_
			 or (oSheet.GetCellByPosition( 6,nCurRow).string <> "")_
			 or (oSheet.GetCellByPosition( 7,nCurRow).string <> "")_
			 or (oSheet.GetCellByPosition( 8,nCurRow).string <> "")_
			 or ((oSheet.GetCellByPosition( 1, nCurRow+1).string <> "") = false)_
			 or (oSheet.GetCellByPosition( 9,nCurRow).string <> ""))
			 	if nCurRow < 2 then
			 		exit do
			 	end if
				nCurRow = nCurRow-1
				oCell=oSheet.GetCellbyPosition(0 , nCurRow)
			'	 ThisComponent.CurrentController.Select(oCell) 'debug
			'	 print
	loop
'print "wow " & lrow	
	lrowS = nCurRow
	oCell = oSheet.GetCellByPosition( 0 , lrowS+1)
	oCell2 = oSheet.GetCellByPosition( lcol5 , lrowS+1)
	Do while (Trova_Attr_N (oCell, oSheet)) <> "End_voce_COMPUTO"'_
		 'And	oCell2.string <> "SOMMANO "
				nCurRow = nCurRow+1
				oCell=oSheet.GetCellbyPosition( 0, nCurRow)
				oCell2 = oSheet.GetCellByPosition( lcol5 , nCurRow)
	loop
	lrowE = nCurRow

 oRangeVoceC = osheet.getCellRangeByPosition (0,lrowS,250,lrowE )

	Circoscrive_Voce_Computo_Att= oRangeVoceC
	
 ' dis_080212 ThisComponent.CurrentController.Select(oRangeVoceC) 'NON eliminare

end Function
SUB CREA_VOCE_ANALISI non usata ' sceglie se usare la macro vecchia o nuova (ovvero cerca la presenza di attributi di cella)
													' TEMO SIA INUTILIZZATA
oSheet = ThisComponent.currentController.activeSheet

	lrow= Range2Cell ' queste 4 righe per ridurre a cella iniziale una eventuale 
	if lrow = -1 then
			 
			exit sub
	end if

	oCell = oSheet.GetCellByPosition( 0 , lrow)
	Do while oSheet.GetCellByPosition( 0, lrow).string = ""
			lrow = lrow-1
			oCell = oSheet.GetCellByPosition( 0, lrow )
		'	ThisComponent.CurrentController.Select(oCell)' debug
		'	print "ecco"
	Loop
 ' 	end if
 	oCell = oSheet.GetCellByPosition( 3 , lrow)
 oCustomXmlAttributes = oCell.UserDefinedAttributes
' xray oCustomXmlAttributes
 If oCustomXmlAttributes.hasElements = false then
print 1
				NuovaAnalisi
		else
print 2
			NuovaAnalisi_Att
	end if
END Sub

Sub Svuota_Computo_VECCHIA () 'svuota tutto (o quasi) il doc di computo ' doradora - da cancellare

'svuota Elenco prezzi
'xray ThisComponent.Sheets
	if ThisComponent.Sheets.hasByName("Elenco Prezzi") = false then
		msgbox "questa macro si usa soltanto in LeenO...",48
		Exit sub
	end if
	Ripristina_statusLine
	barra_apri_chiudi_4

	'>>>>>>>>>>>>>>>>>>>>>>
	Verifica_chiudi_preview
	'<<<<<<<<<<<<<<<<<<<<<<
	
	if msgbox (" La macchina sta per Svuotare questo computo,!"& CHR$(10)_
			& " ma prima salvo il doc corrente, e poi ne salvo una copia con nuovo nome!"& CHR$(10)_
			& " PROSEGUO? ", 4,""& CHR$(10)) = 7 then
		'	&"",4, ""& CHR$(10)) = 7 then
			
		exit sub	
	end if
 oDoc = ThisComponent
 ' Get the document's controller.
 oDocCtrl = oDoc.getCurrentController()
 ' Get the frame from the controller.
 oDocFrame = oDocCtrl.getFrame()	
	
	' salviamo comunque il doc corrente
	oDispatchHelper = createUnoService( "com.sun.star.frame.DispatchHelper" )
	oDispatchHelper.executeDispatch( oDocFrame, ".uno:Save", "", 0, Array() )
 
 	Barra_chiudi_sempre_4
	Barra_Apri_Chiudi_5("Sto svuotando questo documento... Pazienta... ", 30)
	'svuota Elenco Prezzi	
	oSheet = ThisComponent.Sheets.getByName("Elenco Prezzi")
	oEnd=uFindString("Fine elenco", oSheet) 
	lrowFine=oEnd.RangeAddress.EndRow-1
	If lrowFine > 2 then
'		oSheet.rows.removeByIndex (2, lrowFine)
		oRange = osheet.getCellRangeByPosition (0,1,13,lrowFine )
		ThisComponent.CurrentController.Select(oRange)

cancella_dati
		oRange = osheet.getCellRangeByPosition (0,2,0,lrowFine )
		ThisComponent.CurrentController.Select(oRange)
		elimina_riga
thisComponent.currentController.Select(thisComponent.CreateInstance("com.sun.star.sheet.SheetCellRanges")) 'deseleziona
	end If

	Barra_chiudi_sempre_4
	Barra_Apri_Chiudi_5("2 Sto svuotando questo documento... Pazienta... ", 40)


	ThisComponent.Sheets.getByName("S1").GetCellByPosition(7,327).value = 0 
	ThisComponent.Sheets.getByName("S1").GetCellByPosition(7,334).value = 0

'	'sString$ = "Fine elenco"
'	oEnd=uFindString(sString$, oSheet) 
'	lrowFine=oEnd.RangeAddress.EndRow-1
'	oRange = osheet.getCellRangeByPosition (0,1,10,lrowFine )
'	Flags = com.sun.star.sheet.CellFlags.STRING _
'			 + com.sun.star.sheet.CellFlags.HARDATTR _
'				+ com.sun.star.sheet.CellFlags.VALUE
'	oRange.clearContents(Flags)
'	Barra_chiudi_sempre_4
'	Barra_Apri_Chiudi_5("3 Sto svuotando questo documento... Pazienta... ", 50)
	' magari adatta le righe (in certe sistuazioni rimaneva una riga rossa molto alta
	lrowFine= getLastUsedRow(oSheet)
	oCell=oSheet.getCellRangeByPosition(0, 0, 9, lrowFine)
	ThisComponent.CurrentController.Select(oCell)
	Adatta_Altezza_riga
thisComponent.currentController.Select(thisComponent.CreateInstance("com.sun.star.sheet.SheetCellRanges")) 'deseleziona	
'	Adatta_h_riga_intera_tabella("Elenco Prezzi")
fissa(0,1)

''svuota Analisi
	oSheet = ThisComponent.Sheets.getByName("Analisi di Prezzo")
'	lrowFine= getLastUsedRow(oSheet)
	oEnd=uFindString("Fine ANALISI", oSheet) 
	lrowFine=oEnd.RangeAddress.EndRow-1
	If lrowFine > 2 then
'		oSheet.rows.removeByIndex (2, lrowFine)
		oRange = osheet.getCellRangeByPosition (0,1,0,lrowFine )
		ThisComponent.CurrentController.Select(oRange)
		elimina_righe
thisComponent.currentController.Select(thisComponent.CreateInstance("com.sun.star.sheet.SheetCellRanges")) 'deseleziona
	end If
	
	Barra_chiudi_sempre_4
	Barra_Apri_Chiudi_5("4 Sto svuotando questo documento... Pazienta... ", 60)

	' magari adatta le righe (in certe sistuazioni rimaneva una riga rossa molto alta
	lrowFine= getLastUsedRow(oSheet)
	oCell=oSheet.getCellRangeByPosition(0, 0, 10, lrowFine)
	ThisComponent.CurrentController.Select(oCell)
	Adatta_Altezza_riga
thisComponent.currentController.Select(thisComponent.CreateInstance("com.sun.star.sheet.SheetCellRanges")) 'deseleziona
fissa(0,1)

'svuota Computo
	oSheet = ThisComponent.Sheets.getByName("COMPUTO")
	oEnd=uFindString("Fine Computo", oSheet) 
	lrowFine =oEnd.RangeAddress.EndRow-1

'	if oSheet.GetCellByPosition( 18, lrowfine ).cellstyle = "Comp TOTALI num" then
'		'meglio sarebbe cancellarla... ma proviamo così...
'		osheet.getCellRangeByPosition (0,lrowfine,45,lrowfine).CellStyle = "Default"
'	end if

	lrowFine = lrowFine-1 'oEnd.RangeAddress.EndRow-2
	If lrowFine > 2 then
'		oSheet.rows.removeByIndex (2, lrowFine)
		oRange = osheet.getCellRangeByPosition (0,2,0,lrowFine )
		ThisComponent.CurrentController.Select(oRange)
		elimina_righe
thisComponent.currentController.Select(thisComponent.CreateInstance("com.sun.star.sheet.SheetCellRanges")) 'deseleziona
	end if	
	Barra_chiudi_sempre_4	
	Barra_Apri_Chiudi_5("5 Sto svuotando questo documento... Pazienta... ", 70)

Rifa_Somme_TOT_Computo
	' magari adatta le righe (in certe sistuazioni rimaneva una riga rossa molto alta
	oSheet = ThisComponent.Sheets.getByName("COMPUTO")
	oEnd=uFindString("Fine Computo", oSheet) 
	lrowFine =oEnd.RangeAddress.EndRow
	oCell=oSheet.getCellRangeByPosition(0, 0, 43, lrowFine)
	ThisComponent.CurrentController.Select(oCell)
	Adatta_Altezza_riga
thisComponent.currentController.Select(thisComponent.CreateInstance("com.sun.star.sheet.SheetCellRanges")) 'deseleziona
fissa(0,1)

	'svuota CONTABILITA
	Svuota_CONTABILITA_esegui 
	
	'aggiunge una riga vuota "quasi" in testa e la formatta... NON è importante per il codice... ma pare ergonomico
	insRows (2, 1)
'	oSheet.getRows.insertByIndex(2, 1)'
	oSheet.getCellRangeByPosition(0, 2, 49 , 2).cellstyle = "Reg_prog"
	Adatta_h_riga_intera_tabella("CONTABILITA")
	if	sState = "chiusa" then
		osheet.isVisible = false
	end if '
	
	
Ripristina_Validita_Lista	
	salva_temp
	
	'rendi correntexx
	Scrivi_Globale
	Ripristina_statusLine
	Vai_a_M1

	msgbox "Questo documento è stato in gran parte svuotato dai sui dati!..!"
end Sub
Function Rifa_Somma_Analisi (lCol as long, lrow as long) 'dopo che si è iinserita riga - NON USATA
' di % (Molto probab non serve più)
' 
'print lrow 
'dim lcol as long
dim lrowE as long
Dim oCell as object
Dim oCellB as object

	oSheet = ThisComponent.currentController.activeSheet
	 lrowE = lrow
'	 lcol = oCell.CellAddress.column
	 oCell = oSheet.GetCellByPosition( lcol , lrow)
	 ThisComponent.CurrentController.Select(oCell)

	oCellB = oSheet.GetCellByPosition( 6 , lrow)
	 sCol= ColumnNameOf(lcol)

	oCellC = oSheet.GetCellByPosition( 7 , lrow)
	 sColC= ColumnNameOf(lcol+1)
	 
lrow = lrow+2
'cactus
	 do while oSheet.GetCellByPosition( lcol , lrow).string <> "-" And _
	 			(Trova_Attr_N (oSheet.GetCellByPosition( lcol , 0), oSheet)) <> "End_voce_ANALISI" And _
	 			oSheet.GetCellbyPosition( 7, lrow ).CellStyle = "An-sfondo-basso dx"
	 	 	lrow = lrow+1
			ThisComponent.CurrentController.Select(oCell)
'print "!AA"
	 loop

	 lrow = lrow+2	
	 sFormula = "=SUM(" & sCol & lrowE+2 & ":" & sCol & lrow-1 & ")"
	 oCellB.setformula(sFormula )
	'	 ThisComponent.CurrentController.Select(oCellB)
'	 print	sFormula	 

	 sFormulaC = "=SUM(" & sColC & lrowE+2 & ":" & sColC & lrow-1 & ")"
	 oCellC.setformula(sFormulaC )
		 ThisComponent.CurrentController.Select(ocellc)
'	 print	sFormula_i
	 
'print sformula
end Function




'*************************************************************************************************
Function Rifa_Somma_Analisi___________(lCol as long) 'dopo che si è iinserita riga di % (Molto probab non serve più)
' NON USATA (e non so se funziona correttamente...)
dim lrow as long
'dim lcol as long
dim lrowE as long
Dim oCell as object
Dim oCellB as object
'print "lcol " & lCol
'oSheet = thiscomponent.Sheets.getByName ("COMPUTO")
	oSheet = ThisComponent.currentController.activeSheet
	oCell = ThisComponent.CurrentSelection
'	Orange = ThisComponent.CurrentSelection
	 lrow = oCell.CellAddress.row
	 lrowE = lrow
'	 lcol = oCell.CellAddress.column
	 oCell = oSheet.GetCellByPosition( lcol , lrow)
'	 ThisComponent.CurrentController.Select(oCell)
	' print
	 xA = oCell.string
	oCellB = oCell
	 sCol= ColumnNameOf(lcol)
'	 print sCol
	 do while xA <> "-"
	 	lrow = lrow-1
	 	oCell = oSheet.GetCellByPosition( lcol , lrow)
	 	xA = oCell.string
	 loop
	
	 lrow = lrow+2	
	 sFormula = "=SUM(" & sCol & lrowE & ":" & sCol & lrow & ")"
	 oCellB.setformula(sFormula )

end Function

Sub xxxxxxxx
xray ThisComponent.getCurrentSelection(

oSheet = ThisComponent.Sheets.getByName("COMPUTO-")
		iCellAttr = com.sun.star.sheet.CellFlags.OBJECTS
		osheet.getCellRangeByPosition (0,0,140,3).ClearContents(iCellAttr)

End Sub


Sub Svuota_CONTABILITA_esegui_bart (optional sSommari as String) 'non usata
	'ATTENZIONE, se arriva da dialog il parametro non è mai NULL
	' perché intercetta il clic del mouse o cose del genere...
	if thisComponent.Sheets.hasByName("CONTABILITA") then
		oSheet = ThisComponent.Sheets.getByName("CONTABILITA")
		if osheet.isVisible = false then
			osheet.isVisible = TRUE
			sState = "chiusa"
		end if
		oEnd=uFindString("T O T A L E", oSheet) 
		lrowFine =oEnd.RangeAddress.EndRow-1
		'L'ULTIMA (PRIMA DELLA RIGA ROSSA LA SVUOTO SOLTANTO)
		iCellAttr = _
			com.sun.star.sheet.CellFlags.VALUE + _
			com.sun.star.sheet.CellFlags.DATETIME + _
			com.sun.star.sheet.CellFlags.STRING + _
			com.sun.star.sheet.CellFlags.ANNOTATION + _
			com.sun.star.sheet.CellFlags.FORMULA + _
			com.sun.star.sheet.CellFlags.OBJECTS + _
			com.sun.star.sheet.CellFlags.HARDATTR + _
			com.sun.star.sheet.CellFlags.EDITATTR '+ _
			com.sun.star.sheet.CellFlags.STYLES 
		osheet.getCellRangeByPosition (0,lrowfine,48,lrowfine).ClearContents(iCellAttr)
		osheet.getCellRangeByPosition (0,lrowfine,48,lrowfine).CellStyle = "Reg_prog"
		lrowFine = lrowFine-1 ' 'oEnd.RangeAddress.EndRow-2 

		If lrowFine >= 2 then
			oSheet.rows.removeByIndex (12, 22)'lrowFine)
		end if	
	end if
	' magari adatta le righe (in certe sistuazioni rimaneva una riga rossa molto alta
	Adatta_h_riga_intera_tabella("CONTABILITA")
	if	sState = "chiusa" then
		osheet.isVisible = false
	end if '

	if sSommari = "true" then
		' pulisce anche i sommari in Computo
		oSheet = ThisComponent.Sheets.getByName("COMPUTO")
		iColEnd = getLastUsedCol(oSheet)
		iColStart = 44
		if iColEnd-iColStart+1 > 0 then
			oSheet.Columns.removeByindex(iColStart,iColEnd-iColStart+1)	
		end if
	end if
		'aggiunge una riga vuota "quasi" in testa e la formatta... NON è importante per il codice... ma pare ergonomico
	insRows (2,2) 'insertByIndex non funziona
'	oSheet.getRows.insertByIndex(2, 2)'
	oSheet.getCellRangeByPosition(0, 2, 48 , 3).cellstyle = "Reg_prog"
'	oSheet.getCellRangeByPosition(0, 2, 48 , 2).cellstyle = "comp In testa"
	oSheet.getCellRangeByPosition(0, 4, 48 , 4).cellstyle = "Comp TOTALI"
	oSheet.GetCellByPosition(2,2).setstring("QUESTA RIGA NON VIENE STAMPATA")
'	oSheet.GetCellByPosition(25,2).FORMULA="=Z" & lrowFine+3
'	oSheet.GetCellByPosition(25,2).FORMULA="=ROUND(SUBTOTAL(9;Z4:Z" & lrowFine+2 & ");2)"
	oSheet.GetCellByPosition(15,2).FORMULA="=ROUND(SUBTOTAL(9;P3:P4);2)"
'	oSheet.getCellbyPosition(15,2).cellstyle ="comp In testa"
'	oSheet.GetCellByPosition(25,lrowFine+2).FORMULA="=ROUND(SUBTOTAL(9;Z4:Z" & lrowFine+2 & ");2)"
	oSheet.GetCellByPosition(15,4).FORMULA="=ROUND(SUBTOTAL(9;P3:P4);2)"
end sub

SUB Pesca_cod__non_usata 'Pesca_cod__ 'quella di partenza (TUTTO INIZIA DA QUESTA)
'print "uffaaaaa"
	'sceglie tra le due routine
'	questo ambaradan serve per decidere il contesto (si può fare di meglio... ma questa funziona)
'xray smemopesca '<empty>
'print ThisComponent.currentcontroller.activesheet.name 'sMemoPesca
' select Case ThisComponent.currentcontroller.activesheet.name
 ' 		Case "CONTABILITA"
' 			sMemoPesca = "cod_reg"
 ' 			Pesca_cod_per_reg_A
 ' 		Case
' End select
	'	PRINT "prima" & " " & sMemoPesca
	'	exit sub
		If ThisComponent.currentcontroller.activesheet.name="CONTABILITA" then 
		' AND _
		'	sMemoPesca = Nothing then 'empty ' then
			sMemoPesca = "cod_reg"
			Pesca_cod_per_reg_A
			exit sub
			'print sMemoPesca
		end if
		If ThisComponent.currentcontroller.activesheet.name="COMPUTO" then ''AND _
			'	sMemoPesca = empty then
			if sMemoPesca = "cod_reg" then
				Pesca_cod_per_reg_A
				'sMemoPesca = "cod"
				exit sub
			end if
		'	print "è vuoto? " & sMemoPesca
		'	xray sMemoPesca
			if sMemoPesca = "cod" or isempty(sMemoPesca) or sMemoPesca = "" then
		'	print "dentro"
				Pesca_cod_0
				exit sub
			end if
		end if
		If ThisComponent.currentcontroller.activesheet.name="Analisi di Prezzo" then ''AND _
			'sMemoPesca = empty then
			if sMemoPesca = "cod" then
				Pesca_cod_2
				sMemoPesca = empty
			'	print sMemoPesca
				exit sub
			end if
			if isEmpty(sMemoPesca) then
				Pesca_cod_1
				exit sub			
			end if
		end if
		If ThisComponent.currentcontroller.activesheet.name="Elenco Prezzi" then
				Pesca_cod_2
		exit sub
	end if
 
	'PRINT sMemoPesca
	select case sMemoPesca 
		case "cod" 
		'print 1
			Pesca_cod_0
		case "cod_reg"
			Pesca_cod_per_reg_A
	end select
end sub
SUB Pesca_cod_bart_non_usata 'quella di partenza (TUTTO INIZIA DA QUESTA)
'	print "globals " & sGVV & "-" & sGorigine & "-" & sGDove
	Select case ThisComponent.currentcontroller.activesheet.name
		Case "CONTABILITA"
			sGVV = "va"
			lrow = range2cell 'cerca_partenza
			if 	ThisComponent.Sheets.getByName("CONTABILITA").GetCellByPosition(44, lrow).string = "" and _
					ThisComponent.Sheets.getByName("CONTABILITA").GetCellByPosition(1, lrow).cellstyle= "comp Art-EP_R" then
				PRINT "VAI1"
					sGDove = "Elenco Prezzi"
			 		sGVV = "va"	
					sGorigine = "CONTABILITA"	
					cerca_partenza
					'porta su EP ????
				PRINT "VAI2"
					Pesca_cod_1
				PRINT "VAI3"		
					exit sub
				Else
				Print "eccomi"
					sGDove = "COMPUTO"
			 		sGVV = "va"	
					sGorigine = "CONTABILITA"			
					Pesca_cod__per_reg_A_1
					exit sub
					'porta su Computo 		
			end if
		Case "COMPUTO"
			 if sGorigine = "CONTABILITA" then
			 		sGVV = "viene" 'incolla
			 		sGDove = "CONTABILITA"
			 		sGorigine = "COMPUTO"
			 		' esegue viene appropriato
					'preleva solo i dati (da fare)
			 		Pesca_cod__per_reg_A_2
					sGVV = ""
					sGDove = ""		
				 	sGorigine =	""				 		
			 		
			 	else ' 
			 		sGVV = "va"
			 		sGDove = "Elenco Prezzi"
			 		sGorigine = "COMPUTO"
				 	'porta su EP
				 	Pesca_cod_1
			 end if
			 
		Case "Elenco Prezzi"
			'esegui vieni appropriato
			if sGorigine = "CONTABILITA" then
					'preleva solo i dati
					Print 
					Pesca_solo_dati_metti_in_contab
				else
					Pesca_cod_2
			end if
			sGVV = ""
			sGDove = ""		
		 	sGorigine =	""			
		 	
		Case "Analisi di Prezzo"
			sGVV = "va"
			sGDove = "Elenco Prezzi"		
		 	sGorigine =	"Analisi di Prezzo"
		 	'porta su EP
		 	Pesca_cod_1
		 			 
	end select
	
'	print "azzerate? " & sGVV & "-" & sGorigine & "-" & sGDove
	
	
end sub
Sub inseriscirighesopra (InsertPos as long, nRows as long)
		dim document as object
		dim dispatcher as object
'	InsertRow = ThisComponent.getCurrentSelection().celladdress.row
rem questa me l'ha passata A. Vallortigara
	document = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	dim args3(0) as new com.sun.star.beans.PropertyValue
	args3(0).Name = "ToPoint"
	args3(0).Value = "A" & InsertPos+1 'cella di inserimento
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args3())
	for i=1 to nRows
		dispatcher.executeDispatch(document, ".uno:InsertRows", "", 0, Array())
	next
end Sub'********************************************************

SUB SCEMO_____
oSheet = ThisComponent.currentController.activeSheet
 oSelezione = thisComponent.getCurrentSelection()
'checkzelle=HasUnoInterfaces(oSelezione, "com.sun.star.table.XCell" )
'if checkzelle then 
		lrow = oSelezione.celladdress.row
'		lcol = oSelezione.celladdress.column		
	oRange=Circoscrive_Voce_Computo_Att (lrow)
'	ThisComponent.CurrentController.Select(oRange)
END SUB
'_____________________________________________________________________________
Function Circoscrive_Voce_Computo$ (ByVal lrow As Long) 'individua un record di Computo MOLTO vecchio
									' ancora nn è chiaro se serve... ma l'ho aggiornata nel gennaio del 2011 per S5
'---------------------------------------------------------------------------
	
'msgbox "Per favore se vi appare questa finestra scrivetemi citando: Function Circoscrive_Voce_Computo$"
dim lrowS as long
dim lrowE as long
	 oSheet = ThisComponent.currentController.activeSheet 
	lrowS = lRow
	lcol = 0
	lcolS = lcol
	oCellStart = oSheet.GetCellByPosition( lcol , lrow)
 ' ThisComponent.CurrentController.Select(oCell)
 
	If oSheet.GetCellByPosition(2 , lrow ).string = "SOMMANO " then 
 		goto sommano_trovato 
	end if 

	Do while oSheet.GetCellByPosition(2 , lrow ).string <> "SOMMANO " 'or _
 		'	oSheet.GetCellByPosition(2 , lrow ).string <> "SOMMANO" 
 	 lrow = lrow+1
 	 ThisComponent.CurrentController.Select(oSheet.GetCellByPosition( 2 , lrow))
 	 print
 	Loop
 	lrow = lrow+1
	sommano_trovato:

 ' ThisComponent.CurrentController.Select(oCell) 
 lrowE = lrow-1 
 '	xA = ThisComponent.getcurrentselection.getstring
 '	xA = oCell.string	
 	Do while oSheet.GetCellByPosition(0 , lrow ).string = "" ' xA = ""
 		lrow = lrow-1
 	 	'oCell = oSheet.GetCellByPosition( 0 , lrow)
 	 	'xA = oCell.string	
	loop
 	'	oCelle=thisComponent.getCurrentSelection().getCellAddress() 
	 ' lrowS=oCelle.Row - 1
		lrowS= lrow -1
	'	print lrows
	oRangeVoceC = osheet.getCellRangeByPosition (0,lrowS,31,lrowE )
	'	lista = "0, " & lrowS & ",31 ," & lrowE 
	Circoscrive_Voce_Computo= oRangeVoceC
' ThisComponent.CurrentController.Select(oRangeVoceC) 'maria
 'print 

end Function
'================================================================================= 
Sub Formula_magica___

		oSheet = ThisComponent.currentController.activeSheet
		iRow= Range2Cell

	if oSheet.GetCellByPosition( 2, iRow ).cellstyle = "comp 1-a" or _
		oSheet.GetCellByPosition( 5, iRow ).cellstyle = "comp 1-a"	then	
		sSimbDividente = " >| "

		sComp=oSheet.GetCellByPosition( 2, iRow ).string
		if InStr(1, sComp, sSimbDividente) <> 0 then
			sComp = Left(sComp,InStr(1, sComp, sSimbDividente)-1) ' Restituisce la stringa 
		end if

		sString4 = oSheet.GetCellByPosition( 4, iRow ).formula
		sString5 = oSheet.GetCellByPosition( 5, iRow ).formula
		sString6 = oSheet.GetCellByPosition( 6, iRow ).formula
		sString7 = oSheet.GetCellByPosition( 7, iRow ).formula
		sString8 = oSheet.GetCellByPosition( 8, iRow ).formula

		if sString4 <> "" then 
						 sString4 =	"(" & sString4 &")"
						 iTag = 1
		End If	

		if sString5 <> "" then 
						 sString5 =	"(" & sString5 &")" 
						 iTag = 1
		end if	
		if sString4 <> "" and sString5 <> "" and sString6 <> "" then 
						 sString5 =	"*" & sString5 	
				if itag = 1 then 
					sString5 =	"*" & sString5
				end if			
		end if	
	'	sString5 = sString4 & sString5
		
			
		if sString6 <> "" then 
						 sString6 =	"(" & sString6 &")" 
						 iTag = 1
		end if
		if sString5 <> "" and sString6 <> "" and sString7 <> "" then 
						 sString6 =	"*" & sString6 
				else
				if itag = 1 then 
					sString6 =	"*" & sString6
				end if			
		end if	

		
		if sString7 <> "" then 
						 sString7 =	"(" & sString7 &")" 
						 iTag = 1
		end if	
		if sString6 <> "" and sString7 <> "" and sString8 <> "" then 
						 sString7 =	"*" & sString7 
				else
				if itag = 1 then 
					sString7 =	"*" & sString7
				end if			
		end if	
	'	sString7 = sString6 & sString7			
		
		if sString8 <> "" then 
						 sString8 =	"(" & sString8 &")" 
						 iTag = 1
		end if	
		if sString7 <> "" or sString8 <> "" then 
						 sString8 =	"*" & sString8
						else
				if itag = 1 then 
					sString8 =	"*" & sString8
				end if			
		end if
	'	sString8 = sString7 & sString8	
		sString8 = sString4 & sString5 & sString6 & sString7 & sString8
		'Source = "il mio cane"
		Search = "="
		NewPart = ""		
		sString8 = Replace_G (sString8 , Search , NewPart )			
		sTringTUTTA = sComp & sSimbDividente & sString8 

'print sTringTUTTA

		oSheet.GetCellByPosition( 2, iRow ).Formula = sTringTUTTA
	end if'

End Sub
Sub Inserisci_Utili_old ' o Oneri di sicurezza o Maggiorazione %
	nome_sheet = thisComponent.currentcontroller.activesheet.name
	if SE_contabilita = 0 and nome_sheet <> "CONTABILITA" then
		exit sub
	end if

	'_____________________
	chiudi_dialoghi ' chiude tutti i dialoghi
	'_____________________
	
	If thisComponent.Sheets.hasByName("S1") Then ' se la sheet esiste
		If ThisComponent.Sheets.getByName("S1").GetCellByPosition(7,326).value = 3 then
				Inserisci_Utili_giuserpe
			else
				Inserisci_Utili_Classica
		end if
	end if
end sub
Sub Inserisci_Incid_manodopera()' obsoleta ... cancellare poi...

dim nome as string
Dim oDoc As Object
Dim oSheets As Object
dim oCelle As Object
Dim CellRangeAddress As New com.sun.star.table.CellRangeAddress
Dim CellAddress As New com.sun.star.table.CellAddress
dim lrow as integer
dim lcol as integer
dim lrow2 as integer
Dim oView As Object
Dim nome_sheet as string
Dim OcalcSheet as Object
Dim I as long
Dim oSheet_num as integer
Dim iflag
	oDoc=thisComponent
	oDoc.SupportsService("com.sun.star.sheet.SpreadsheetDocument")
	oCelle=oDoc.getCurrentSelection().getCellAddress()
	lrow=oCelle.Row	
	oSheets = odoc.Sheets
	oView = ThisComponent.CurrentController
	nome_sheet = oView.GetActiveSheet.Name
	
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
	iflag =	CNTR_Analisi (lrow)' controlla se la riga è buona
			sFlag = "A"

	if iflag = 1 then
		exit sub
	end if
	oCalcSheet = ThisComponent.currentController.activeSheet
'	oCalcSheet = oSheets.GetByIndex(0)
'	For I = 0 to oSheets.Count -1 
'		oCalcSheet = oSheets(I) 'recuperiamo la tabella
'		if oCalcSheet.Name = nome_sheet Then
'			oSheet_num = I
'		end if
'	Next I
'	oSheets = oDoc.Sheets (oCalcSheet)
'	CellRangeAddress.Sheet = oCalcSheet 
'	CellRangeAddress.StartColumn = 0
'	CellRangeAddress.StartRow = lrow
'	CellRangeAddress.EndColumn = 250 
'	CellRangeAddress.EndRow = lrow
'	print lrow
'	oSheets.insertCells(CellRangeAddress, com.sun.star.sheet.CellInsertMode.ROWS)' inserisce delle righe vuote
'exit sub
'print
'	lrow2 = lrow +1
'	CellAddress.Sheet = oSheet_num 
'	CellRangeAddress.StartRow = lrow2
'	CellRangeAddress.EndRow = lrow2
'	CellAddress.Column = 0
'	CellAddress.Row = lrow
'	oSheets.copyRange(CellAddress, CellRangeAddress)

	if sflag = "A" then ' ????
			lcol = 1
		else
			lcol = 3
	end if
'	oCell = oSheets.GetCellByPosition( lcol , lrow+1)	
'	ThisComponent.CurrentController.Select(oCell)
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	oSheet = ThisComponent.currentController.activeSheet ' sheet corrente 
 oCelle=thisComponent.getCurrentSelection().getCellAddress() 
 lrow=oCelle.Row 
 oCell = oSheet.GetCellByPosition(0,lrow ) 
 ThisComponent.CurrentController.Select(oCell)
 
 '''''''' copia gli utili %
	oDoc = ThisComponent
	DocView=oDoc.getCurrentController()
	oSheet1=oDoc.NamedRanges.utili.ReferredCells ' utili è il nome del range
	oCelle=oDoc.getCurrentSelection().getCellAddress() ' indirizzo cella attiva (qui)
	oSheet2 = oDoc.Sheets.getByName(oDoc.currentcontroller.activesheet.name) ' sheet corrente dove incollare
	oQuellRangeAddresse = osheet1.getRangeAddress
 ' oSheet2.copyRange(oCelle,oQuellRangeAddresse)
 	oCell = oSheet.GetCellByPosition( 7 , lrow)	
	ThisComponent.CurrentController.Select(oCell)
'	print
	Range_Somma_locale_analisi_incidenza (6)
	oCell.cellstyle="An-lavoraz-dx% ins"
	msgbox "Controlla la formula... potrebbe non essere giusta!!!"
END SUB
' Questa Ã¨ molto traballante... da rivedere completamente...

Function Disattiva_button 'viene richiamata dopo l'esecuzione 
print "oops... non credevo venisse usata"
	oSheet = ThisComponent.currentController.activeSheet
	oDpage = oSheet.DrawPage
 	oform = oDpage.Forms.getbyname("WW-Standard")
 	oCtrlModel = oform.getbyname("PushButton2")
 	msgbox oCtrlModel.Enabled
 	oCtrlModel.Enabled = False
' msgbox oCtrlModel.Enabled	
 '	Wait 3000
end Function 
sub Zoo_a_default ' da cancellare
print "Focus.Zoo_a_default - Ma non era disattivata?"	
	oSheet = ThisComponent.Sheets.getByName("S1")
	orange = oSheet.getCellRangeByPosition(0,0,0,37)
	oContr = ThisComponent.CurrentController	 
	oContr.select(oRange)
	oContr.ZoomType=OPTIMAL
	oContr.select (oSheet.getCellByPosition(1,1))


	oSheet = ThisComponent.Sheets.getByName(_
	ThisComponent.currentcontroller.activesheet.name)
'	orange = oSheet.getCellRangeByPosition(0,0,15,37)
	orange = oSheet.getCellRangeByPosition(10,0,16,37)
	oContr = ThisComponent.CurrentController	 
	oContr.select(oRange)
'xray oContr '.ZoomType
	oContr.ZoomType=OPTIMAL
	oContr.select (oSheet.getCellByPosition(10,0))
'	oContr.select (oSheet.getCellByPosition(0,0))

end sub
SUB Pesca_cod_0 ' probabilmente non usata...
'0) verifica contesto
' decide se azionare Pesca_cod_1 oppure Pesca_cod_2

	oSheet = ThisComponent.currentController.activeSheet
	sSheetName = ThisComponent.currentcontroller.activesheet.name
	If sSheetName="COMPUTO" or sSheetName="Analisi di Prezzo" then
	print "3"
		Pesca_cod_1
		exit sub
	end if
	If sSheetName="Elenco Prezzi" then
		Pesca_cod_2
		exit sub
	end if
'	msgbox "questa macro va usata solo su Computo o su EP"
END SUB
sub Struttura_ComputoCONT 'PROVA PROVA per il foglio contabilita NON USATA
print "questo" 
	'>>>>>>>>>>>>>>>>>>>>>>>>>>> per evitare che duplichi la struttura
	 Togli_Struttura
	 '<<<<<<<<<<<<<<<<<<<<<<<<<<<
	Ripristina_statusLine
	Barra_Apri_Chiudi_5("#1 Sto creando le strutture... Pazienta...", 10)
	oSheet = ThisComponent.currentController.activeSheet 'oggetto sheet
	iSheet = oSheet.RangeAddress.sheet ' index della sheet
'	sAttributo = Trova_Attr_Sheet
 If osheet.Name <> "COMPUTO" Then	
 			 msgbox "#8 Questo comando si può usare solo" & CHR$(10)_
 			 &" in una tabella di COMPUTO!", 16, "AVVISO!"
			exit sub
	end if
 	if ThisComponent.currentcontroller.activesheet.name = "COMPUTO" then
		s_R = ""
	end if
	if ThisComponent.currentcontroller.activesheet.name = "CONTABILITA" then
		s_R = "_R"
	end if		
	if right( (oSheet.GetCellByPosition(0 ,5).CellStyle), 2) = "_R" or	_
				right( (oSheet.GetCellByPosition(0 ,6).CellStyle), 2) = "_R" or _
				right( (oSheet.GetCellByPosition(0 ,7).CellStyle), 2) = "_R" or _
				right( (oSheet.GetCellByPosition(0 ,8).CellStyle), 2) = "_R" then
				s_R = "_R"	
	end if	

	
	lStartRow = oSheet.GetCellByPosition( 0 , 0)
'	lLastUrowNN = getLastUsedRow(oSheet)
	
	sString$ = "Fine Computo"
	
	oEnd=uFindString(sString$, oSheet) 
	If isNull (oEnd) or isEmpty (oEnd) then 
						lLastUrowNN = getLastUsedRow(oSheet)-1
					else
						lLastUrowNN=oEnd.RangeAddress.EndRow -1
	end if
	IF lLastUrowNN	>= 600 and lLastUrowNN	<= 1000 then
		if msgbox ("Questa tabella è piuttosto corposa e ci vorrà un po' di tempo! " & CHR$(10)_
		 	& " (alcuni minuti...) " & CHR$(10) & CHR$(10)_
			&" PROSEGUO ?... " & CHR$(10) _
 			& CHR$(10) & "",36, "Creazione STRUTTURE") = 7 then
 			exit sub
 	end if
	end if
	IF lLastUrowNN	> 1001 then
		if msgbox ("Questa tabella è piuttosto corposa e ci vorrà un po' di tempo! " & CHR$(10)_
			& " ( Parecchi minuti...) " & CHR$(10) & CHR$(10)_
			&" PROSEGUO ?... " & CHR$(10) _
 			& CHR$(10) & "",36, "Creazione STRUTTURE") = 7 then
 			exit sub
 	end if
	end if
	'lLastUrowNN=oEnd.RangeAddress.EndRow-1
	ThisComponent.Sheets.getByName("S1").GetCellByPosition(7,280).string = NOW 'debug
	Ripristina_statusLine
	Barra_Apri_Chiudi_5("2 Sto creando le strutture... Pazienta...", 15)
	iLivelliVisibili = ThisComponent.Sheets.getByName("S1").GetCellByPosition(7,297).value 
	if iLivelliVisibili=0 then 'per avere comunque un valore di default nel caso il riferimento sia vuoto
			iLivelliVisibili = 2
	end if
For I = 3 to lLastUrowNN
Dim oCRA As New com.sun.star.table.CellRangeAddress
	iRiga = i
'	do while oSheet.GetCellByPosition(0 , i).CellStyle = "Livello-1-scritta" & s_R
			
		'		print left((oSheet.GetCellByPosition(1 , (I)).CellStyle),7) 'debug
		' se è dentro una intestazione di capito o sottoCap
		if	oSheet.GetCellByPosition(1 , (I)).CellStyle = "Comp Start Attributo_R" & s_R or _
		 		oSheet.GetCellByPosition(1 , (I)).CellStyle = "Livello-1-scritta" & s_R or _
		 		oSheet.GetCellByPosition(1 , (I)).CellStyle = "livello-1-sotto_" & s_R or _
			 	oSheet.GetCellByPosition(1 , (I)).CellStyle = "livello2 valuta" & s_R or _
			 	oSheet.GetCellByPosition(1 , (I)).CellStyle = "livello-2-sotto_" & s_R or _
			 	oSheet.GetCellByPosition(1 , (I)).CellStyle = "livello2 sopra" & s_R or _
			 	oSheet.GetCellByPosition(1 , (I)).CellStyle = "Default" or _
			 	left((oSheet.GetCellByPosition(1 , (I)).CellStyle),7) = "Reg-SAL" then	 'questa per usare stile gerarchico
	'	ThisComponent.CurrentController.Select(oSheet.GetCellByPosition(1 ,I )) 'debug	
	'	print "cap " & oSheet.GetCellByPosition(1 ,I ).CellStyle
				goto prossima ' passa alla riga dopo
			else 
				iRI=I
				Do while oSheet.GetCellByPosition(1 , (I)).CellStyle = "Comp-Bianche sopra" & s_R OR _
						oSheet.GetCellByPosition(1 , (I)).CellStyle = "comp Art-EP" & s_R OR _
						oSheet.GetCellByPosition(1 , (I)).CellStyle = "Comp-Bianche in mezzo" & s_R OR _
						oSheet.GetCellByPosition(1 , (I)).CellStyle = "comp sotto Bianche" & s_R ' ' prima avevo OR
				'	ThisComponent.CurrentController.Select(oSheet.GetCellByPosition(1 ,I )) 'debug	
				'	print oSheet.GetCellByPosition(1 ,I ).CellStyle
						if oSheet.GetCellByPosition(1 , (I)).CellStyle = "Comp-Bianche in mezzo" & s_R then
							iRIs = I
							Do while oSheet.GetCellByPosition(1 , (I)).CellStyle = "Comp-Bianche in mezzo" & s_R 
							'	ThisComponent.CurrentController.Select(oSheet.GetCellByPosition(1 ,I )) 'debug
							'	print 'ThisComponent.CurrentController.Select(oSheet.GetCellByPosition(1 ,I )).CellStyle'debug
								I = I + 1 'genni
							loop
							if iRIs > I-1 then
									ThisComponent.CurrentController.Select(oSheet.getCellByPosition(1,iRI))
									beep
									oCRA.EndRow = iRIs
									msgbox "1 Da queste parti lo stile delle righe non sembra ordinato! " & I-1 & CHR$(10)_
											& "Io ho raggruppato ugualmente... ma se la struttura non ti sembrerà ordinata non dare la colpa a me... -;)"	
								else
									oCRA.EndRow = I-1	
							end if
							oCRA.Sheet =iSheet
							oCRA.StartColumn = 	4
							oCRA.StartRow = iRIs
							oCRA.EndColumn = 9
							oSheet.group(oCRA,1,1)
							oSheet.showLevel(iLivelliVisibili-1,1)
							'oSheet.hideDetail(oCRA)
						end if
						I=I+1
				loop
				'print "michiaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
				if iRI > I-1 then
						ThisComponent.CurrentController.Select(oSheet.getCellByPosition(1,iRI))
						beep
						oCRA.EndRow = iRi
						msgbox "In questo punto lo stile delle righe non sembra ordinato! "& CHR$(10)_
								& "Io ho raggruppato ugualmente... ma se la struttura non ti sembrerà ordinata non dare la colpa a me... -;)"	
					else
						oCRA.EndRow = I-1	
				end if
				oCRA.Sheet =iSheet
				oCRA.StartColumn = 	5
				oCRA.StartRow = iRI
				oCRA.EndColumn = 5
				'oCRA.EndRow = I-1
				oSheet.group(oCRA,1)
			'	oSheet.hideDetail(oCRA)
				oSheet.showLevel(iLivelliVisibili-1,1)
			'	oSheet.showLevel(1,1)

				I = I+1 
		end if
		prossima:
	'	ThisComponent.CurrentController.Select(oSheet.GetCellByPosition(1 ,I )) 'debug	
	'	print oSheet.GetCellByPosition(1 ,I ).CellStyle
	'	if i = 10 then ' and i <= lLastUrowNN / 4 + 1 then
		'	Ripristina_statusLine
			sTxt = "Sto creando le strutture e sono alla riga " & i
		'	Barra_Apri_Chiudi_5( stxt, 50)
	'	end if
		goto salta_debug

		Ripristina_statusLine
		Barra_Apri_Chiudi_5(sTxt, lLastUrowNN/5*100/lLastUrowNN)
		if i >= lLastUrowNN / 4 then ' and i <= lLastUrowNN / 4 + 1 then
			Ripristina_statusLine
			Barra_Apri_Chiudi_5(sTxt, lLastUrowNN/4*100/lLastUrowNN)
		end if
		if i >= lLastUrowNN/4*2 then 'and i <= lLastUrowNN/4*2 + 1 then
			Ripristina_statusLine
			Barra_Apri_Chiudi_5(sTxt , lLastUrowNN/4*2*100/lLastUrowNN)
		end if
		if i >= lLastUrowNN/4*3 then 'and i <= lLastUrowNN/4*3 + 1 then
			Ripristina_statusLine
			Barra_Apri_Chiudi_5(sTxt, lLastUrowNN/4*3*100/lLastUrowNN)
		end if

		salta_debug:
next I

	oCRA.Sheet =iSheet
	oCRA.StartColumn = 	4
	oCRA.StartRow = 1
	oCRA.EndColumn = 8
	oCRA.EndRow = 100
	oSheet.group(oCRA,0)
	Ripristina_statusLine
	'ThisComponent.Sheets.getByName("S1").GetCellByPosition(7,281).string = NOW 'debug
end sub
SUB Accoda_Marche(livelli as long, lrow as long) ' Questa non credo sia utilizzata... ma per il momento rimane
'print "Dovrebbe esere disattivata... (Sub Accoda_Marche)"
'exit sub
	
 oSheet = ThisComponent.currentController.activeSheet ' sheet corrente 
 
 oCell = oSheet.GetCellByPosition( 2, lrow-1 ) 
 ' ThisComponent.CurrentController.Select(oCell)'debug	
 oSheetTemp = thisComponent.sheets.getbyname("Temp") 
 dim iIncMan as double
 
 	if oSheetTemp.getCellByPosition(5,1).string = "incid. manodopera" or _
 		oSheetTemp.getCellByPosition(6,1).string <> "" then
	'	 sIncMan = oSheetTemp.GetCellByPosition( 8, lrow).getvalue
	end if

 sIncMan = oSheet.GetCellByPosition( 8, lrow).getvalue
 sSicur = oSheet.GetCellByPosition( 9, lrow).getvalue
 ' print sIncMan
 
	if oSheet.GetCellByPosition( 7, lrow).value = isNotANumber then
			Prezzo = oSheet.GetCellByPosition( 7, lrow).getstring
				else
			Prezzo = oSheet.GetCellByPosition( 7, lrow).getvalue
	end if
	sUM = oSheet.GetCellByPosition( 6, lrow).string
	sDescr0 = oSheet.GetCellByPosition( 4, lrow).string
	sAlfaNum0 = oSheet.GetCellByPosition( 2, lrow).string
'	print sAlfaNum0
	sAlfaNum1 = oSheet.GetCellByPosition( 2, lrow-1).string
'	print sAlfaNum1
	lAlfa0 = Len (sAlfaNum0)
	oCell = oSheet.GetCellByPosition( 2, lrow )
	lAlfaA = Len(oCell.string)
	sAlfac1 = oSheet.GetCellByPosition( 2, lrow-1).string
	sCategoria = oSheet.GetCellByPosition( 0, lrow).string
	oSheetTemp = thisComponent.sheets.getbyname("Temp")
 
 '	print lAlfa0
'	print lAlfaA
'print livelli
	Select Case livelli

	Case 1 
	
	msgbox "il livello in questo caso deve essere impostato a 2"
	exit sub
	'	lAlfaA = Len(sAlfaNum1)
	'	Do while lAlfa0 = lAlfaA 
	'			lrow = lrow-1
	'			oCell = oSheet.GetCellByPosition( 2, lrow )
				'lAlfaA = Len( oCell.string)
	'			sAlfaNum1 = oSheet.GetCellByPosition( 2, lrow).string	
	''			ThisComponent.CurrentController.Select(oCell)'debug	
	'			print "AAA"		
	'	loop 
	'	lrow = lrow+1
	'	sDescr1 = oSheet.GetCellByPosition( 4, lrow-1).string

	Case 2 ' per Marche
print lrow
			lrow = lrow-1
			oCell = oSheet.GetCellByPosition( 2, lrow )
			lAlfaA = Len( oCell.string)
			sAlfac1 = oSheet.GetCellByPosition( 2, lrow).string
print sAlfac1
'print oSheet.GetCellByPosition( 2, lrow+1).string
'print oSheet.GetCellByPosition( 2, lrow).string
			if oSheet.GetCellByPosition( 6, lrow-1).string = "" AND _
			 	 sAlfac1 <> oSheet.GetCellByPosition( 2, lrow-1).string	then 'oSheet.GetCellByPosition( 2, lrow).string then
		'	print "diretto"
				goto voce_completa
			end if 
			Do while sAlfac1 = oSheet.GetCellByPosition( 2, lrow-1).string
				if oSheet.GetCellByPosition( 6, lrow).string = "" then
					sDescr1 = oSheet.GetCellByPosition( 4, lrow).string
					print sDescr1
					goto voce_completa
				end if 
				lrow = lrow-1	
				ThisComponent.CurrentController.Select(oSheet.GetCellByPosition( 2, lrow))'debug	
				print "proseguo"
			loop 
			print lrow & " in uscita"
		'	lrow = lrow +1
			sDescr1 = oSheet.GetCellByPosition( 4, lrow).string
	print 	sDescr1
			If sAlfac1 = "P" then ' se è milano...	
		'	print "MI"
					goto voce_completa
			end if
			goto voce_completa
			oCell = oSheet.GetCellByPosition( 2, lrow )
			lAlfa0 = lAlfaA
			lAlfaA = Len( oCell.string)
	
			Do while lAlfaA = lAlfa0 
				lrow = lrow-1
				oCell = oSheet.GetCellByPosition( 2, lrow )
				lAlfaA = Len( oCell.string)
				sAlfac1 = oSheet.GetCellByPosition( 3, lrow).string
			'		ThisComponent.CurrentController.Select(oCell)'debug	
			'	print sAlfac1

			loop
			IF lAlfaA > lAlfa0 then
		'		goto voce_completa
			end if
			sDescr2 = oSheet.GetCellByPosition( 4, lrow).string			
			oCell = oSheet.GetCellByPosition( 2, lrow )

	Case 3 '
	'	msgbox "il livello in questo caso deve essere impostato a 2"
	'	exit sub
'	print "sono in tip 3"
			lrow = lrow-1
			oCell = oSheet.GetCellByPosition( 2, lrow )
			lAlfaA = Len( oCell.string)
			sAlfac1 = oSheet.GetCellByPosition( 3, lrow).string

			Do while lAlfaA = lAlfa0 AND sAlfac1 <> "P" 
				lrow = lrow-1
				oCell = oSheet.GetCellByPosition( 2, lrow )
				lAlfaA = Len( oCell.string)
				sAlfac1 = oSheet.GetCellByPosition( 3, lrow).string			
			'	ThisComponent.CurrentController.Select(oCell)'debug	
			'	print sAlfac1
			loop 
			sDescr1 = oSheet.GetCellByPosition( 4, lrow).string
'	print "1 " & sDescr1
			If sAlfac1 = "P" then ' se è milano...
	
					sDescr1 = oSheet.GetCellByPosition( 4, lrow).string
					goto voce_completa
			end if
			
'		PRINT "SECONDO"
			'lrow = lrow+1 
			oCell = oSheet.GetCellByPosition( 2, lrow )
		'	ThisComponent.CurrentController.Select(oCell)'debug	
	'	print "dove"	
		'	lAlfa1 = lAlfaA
			lAlfa1 = Len( oCell.string)
			lrow = lrow-1
			oCell = oSheet.GetCellByPosition( 2, lrow )
			lAlfaB = Len( oCell.string)
			IF lAlfaB > lAlfa1 then
			'		print lAlfaB & " ancora " & lAlfa1
					Do while lAlfaB >= lAlfa1
						lrow = lrow-1
						oCell = oSheet.GetCellByPosition( 2, lrow )
						lAlfaB = Len( oCell.string)
					'	ThisComponent.CurrentController.Select(oCell)'debug	
				'		print "3c"
					loop
					sDescr2 = oSheet.GetCellByPosition( 4, lrow).string
				else
				'	ThisComponent.CurrentController.Select(oCell)'debug	
				'		print "Altern"
					sDescr2 = oSheet.GetCellByPosition( 4, lrow).string
			end if
		
			IF lAlfaA > lAlfa0 then
			
		''		goto voce_completa
			end if
		'	sDescr2 = oSheet.GetCellByPosition( 4, lrow).string
		'	print "2 " & sDescr2			
			oCell = oSheet.GetCellByPosition( 0, 0 )
	'			oCell = oSheet.GetCellByPosition( 2, lrow )


	End select
	 

	voce_completa:

	If sDescr1 = sDescr2 then ' su milano ci sono voci dove 
	'la descr è ripetuta tal quale... e possiamo eliminarla subito.
		sDescr2 = "" '
	end if
	If sDescr0 = sDescr1 then
		sDescr1 = ""
	end if	
	
	oUltimo_indirizzo_conosciuto = ThisComponent.CurrentSelection

' tutti i dati sono adesso stivati nelle variabili procedo a ripulire e incollare
	oSheet = oSheetTemp
	Thiscomponent.currentcontroller.setactivesheet(oSheet)


	Flags = com.sun.star.sheet.CellFlags.STRING + _
			com.sun.star.sheet.CellFlags.VALUE + _
			com.sun.star.sheet.CellFlags.FORMULA
	oRange = oSheet.getCellRangeByPosition (1,2,7,2)
 	oRange.clearContents(Flags) '@@@ pare non cancellare tutto... aggiungere flag?
'print sCategoria
	if sCategoria <> "" then
'		oSheet.getCellByPosition(5,5).string = sCategoria 
	end if

	SCompleta1 = sAlfaNum0
'	oSheet.getCellByPosition(1,2).string = SCompleta1

	if Len( sDescr3) <> 0 then
		 sDescr3 = sDescr3 & CHR(13)
	end if
	if Len( sDescr2) <> 0 then
		 sDescr2 = sDescr2 & CHR(13)
	end if
	if Len( sDescr1) <> 0 then
		 sDescr1 = sDescr1 & CHR(13)
	end if
	
	' introdotto con TV
	' azioni diverse per TV e T3
'	print oSheet.getCellByPosition(5,1).string
	if oSheet.getCellByPosition(5,1).string = "incid. manodopera" or _
		oSheet.getCellByPosition(6,1).string <> "" then
			if sIncMan<> 0 then
				oSheet.getCellByPosition(5,2).value = sIncMan 'SCompleta0
			end if		
'PRINT "DDDDD"'
'			SCompleta5 = "=_CONCATENATE(""("";G5; "")""; "" ""; " & " " & """" & sAlfaNum0 & """" & ")"
			SCompleta5 = "=CONCATENATE(""("";G5; "")""; "" ""; " & " " & """" & sAlfaNum0 & """" & ")"
			oSheet.getCellByPosition(6,2).formula = SCompleta5
		Else
	'		SCompleta5 = "=CONCATENATE(""("";F5; "")""; "" ""; " & " " & """" & sAlfaNum0 & """" & ")"
			SCompleta5 = "=CONCATENATE(""("";F5; "")""; "" ""; " & " " & """" & sAlfaNum0 & """" & ")"
			oSheet.getCellByPosition(5,2).formula = SCompleta5
	end if
	
	if oSheet.getCellByPosition(7,1).value <> "" then
		if sSicur <> 0 then 
			oSheet.getCellByPosition(7,2).value = sSicur
		end if
	end if
	
	SCompleta2 = sDescr3 & sDescr2 & sDescr1 & sDescr0
 	oSheet.getCellByPosition(2,2).string=SCompleta2

	SCompleta3 = sUM
	oSheet.getCellByPosition(3,2).string=SCompleta3 

	SCompleta4 = Prezzo
	oSheet.getCellByPosition(4,2).value = SCompleta4 'Attenzione!	
	' il prezzo DEVE essere un numero VERO (non stringa)

 oCell=oSheet.getCellByPosition(2,2) 
 ' ThisComponent.CurrentController.Select(oCell)
 ' Adatta_Altezza_riga 	
 sQualeCella = "$C$3"
 Seleziona_Cella (sQualeCella)
 Adatta_Altezza_riga 	
 ' pRINT "ACCODA NORMAL"
end sub
Sub Sposta_Voce_Analisi ' riferita alla ver 6 di circoscrivi' ' sposta una analisi di prezzo in nuova posizione...
dim lSRow as long
dim oErow as long
dim StartRow as long
dim lrow as long
	lrow= Range2Cell ' queste 4 righe per ridurre a cella iniziale una eventuale 
	if lrow = -1 then
		 
		exit sub
	end if
	oSheet = ThisComponent.currentController.activeSheet
	oCell = oSheet.GetCellByPosition( 0 , lrow)' errata selezione di un range
 	ThisComponent.CurrentController.Select(oCell)
'	print "alt"
	
	sStRange = CircoscrivileAnalisi_555 (lrow)
'	oRangeVoce = CircoscrivileAnalisi_6$ (lrow)
'xray sStRange
	If IsNull(sStRange) Then 
 		ThisComponent.CurrentController.Select(oCell)
		msgbox "Devi essere all'interno di una voce... Altrimenti non so che cosa vuoi spostare... RIPROVA!!++ "
		exit sub
	end if

'	lrowS = oRangeVoce.RangeAddress.Startrow '+1
'	lrowE = oRangeVoce.RangeAddress.Endrow
	
'	ThisComponent.CurrentController.Select(oRangeVoce)
	
'	oOldSelection = ThisComponent.CurrentSelection
	
'	Riprova:
'______________________________________________________-	
	sTitolo = " Click sulla riga dove spostare l'analisi (ESC per Annullare, NO Click su X ) "
	SelectedRange = getRange(sTitolo) ' richiama il listeners
 	if SelectedRange = "" or _
 	 	SelectedRange = "ANNULLA" then
 	 	ThisComponent.currentController.removeRangeSelectionListener(oRangeSelectionListener)
 	 	exit sub
 	end if
	StartRow = getRigaIniziale(SelectedRange)
	
	'''''''''''''''''''''''''''''''''''''''''''''
 	sString$ = "Fine ANALISI" ' in caso di click fuori zona...
	oEnd=uFindString(sString$, oSheet)
	lrowF=oEnd.CellAddress.Row 

	If lrowF < StartRow-1 then
		oCellK = oSheet.GetCellByPosition( 0 , StartRow)
		ThisComponent.CurrentController.Select(oCellK)
		msgbox " Hai selezionato una destinazione ESTERNA all'area " & CHR$(10)_
		& " definita dalla riga rossa di chiusura... "& CHR$(10) & CHR$(10)_
		& " e questo non è consentito!..."
		ThisComponent.CurrentController.Select(sStRange)
		exit sub
	end if
''''''''''''''''''''''''''''''''''''''''''''''''''''

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
	StartRowM = Mettiti_esattamente_tra_due_Analisi (StartRow)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
print "alfa " & StartRowM 
'xray sStRange
ThisComponent.CurrentController.Select(sStRange)
'print
	Sposta_range_buono (StartRowM) ',sStRange)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'	StartRowM = Mettiti_esattamente_tra_due_Analisi (StartRow)
'	if StartRowM = "ciucca" then
'		Print "Cerca di mirare meglio..."
'		goto Riprova
'	end if
'	if StartRowM = 0 then
'		oSheet = ThisComponent.currentController.activeSheet ' sheet corrente 
 ''		oCell = oSheet.GetCellByPosition( 0, StartRow-1 )
' 		ThisComponent.CurrentController.Select(oCell)
'		msgbox ("Errore! Credo tu abbia selezionato DENTRO una analisi... "_
'				&"Devi essere più preciso... clicca TRA due analisi... "_
'				&"Altrimenti non so dove vuoi spostare.. RIPROVA!! ", "Guarda dove clikki!")
'		goto Riprova
'		ThisComponent.CurrentController.Select(oOldSelection)			
'	end if
'	ThisComponent.CurrentController.Select(oOldSelection)	
'	StartRow = StartRowM ''''
 ' 	Sposta_range_buono(StartRow+1)
	'----------------------------------------------------------
END SUB

'==========================================================================


SUB Tronca_Altezza_Voci_Computo_NON_USATA '1 e Contabilità NON USATA
		oSheet = ThisComponent.currentController.activeSheet
	'	Print 	right( (oSheet.GetCellByPosition(0 ,5).CellStyle), 2)
 		if right( (oSheet.GetCellByPosition(0 ,5).CellStyle), 2) = "_R" or	_
				right( (oSheet.GetCellByPosition(0 ,6).CellStyle), 2) = "_R" or _
				right( (oSheet.GetCellByPosition(0 ,7).CellStyle), 2) = "_R" or _
				right( (oSheet.GetCellByPosition(0 ,8).CellStyle), 2) = "_R" then
				s_R = "_R"	
			else
				s_R = ""
		end if	' ha impostato un flag
	oActiveCell1 = thisComponent.getCurrentSelection()
	lStartRow = oSheet.GetCellByPosition( 1 , 0)
	numV = 1
	
	'trovo la fine dei dati
	oEnd=uFindString("TOTALI COMPUTO", oSheet) 
	If isNull (oEnd) or isEmpty (oEnd) then 
			lLastUrowNN = getLastUsedRow(oSheet)
		else
			lLastUrowNN=oEnd.RangeAddress.EndRow '-1
	end if
	
	lrow= Range2Cell
	if lrow = -1 then exit Sub
	if lrow > lLastUrowNN Then lrow = lLastUrowNN-3

	for i = lrow to lLastUrowNN
			if oSheet.GetCellByPosition(1 ,i).CellStyle = "comp Art-EP" then 
				lrow =i
				exit for
			end if
	next
		
	'controllo se la riga corrente	è quella base
	if	oSheet.GetCellByPosition(1 , lrow).CellStyle = "comp Art-EP" & s_R then
		goto verifica 
	end if

	' Altrimenti questo ciclo cerca nella voce la riga base
	oRangeVC = Circoscrive_Voce_Computo_Att(lrow)
 For i = oRangeVC.RangeAddress.StartRow to oRangeVC.RangeAddress.EndRow
		if	oSheet.GetCellByPosition(1 , (i)).CellStyle = "comp Art-EP" & s_R then	 
			 lrow = i
		end if
	Next i
	verifica:
	' se la riga 'base' NON è ottimizzata in altezza
	if oSheet.GetCellByPosition(1 ,lrow).Rows.OptimalHeight = false then
		' le allunga
	 	sString$ = "Fine Computo"
		oEnd=uFindString(sString$, oSheet) 
		If isNull (oEnd) or isEmpty (oEnd) then 
				lLastUrowNN = getLastUsedRow(oSheet)
			else
				lLastUrowNN=oEnd.RangeAddress.EndRow '-1
		end if
		oRange = oSheet.getCellRangeByPosition (1,2,5,lLastUrowNN)
		oRange.Rows.OptimalHeight = true
		ThisComponent.CurrentController.Select(oActiveCell1)
		exit sub
	end if

	If thisComponent.Sheets.hasByName("S1") Then 
			If ThisComponent.Sheets.getByName("S1").GetCellByPosition(7,310).value = 0 then
					lAltezzaRiga = 1200
				else
					lAltezzaRiga =_
					ThisComponent.Sheets.getByName("S1").GetCellByPosition(7,310).value * 1000
			end if				 	
		else
			lAltezzaRiga = 1300
	end if
			
	lStartRow = oSheet.GetCellByPosition( 0 , 0)
	
	sString$ = "Fine Computo"
	oEnd=uFindString(sString$, oSheet) 
	If isNull (oEnd) or isEmpty (oEnd) then 
			lLastUrowNN = getLastUsedRow(oSheet)
		else
			lLastUrowNN=oEnd.RangeAddress.EndRow-1
	end if
	
	If thisComponent.Sheets.hasByName("S1") Then '???
			iLivelliVisibili = ThisComponent.Sheets.getByName("S1").GetCellByPosition(7,297).value '?????????
			if iLivelliVisibili=0 then 'per avere comunque un valore di default nel caso il riferimento sia vuoto
					iLivelliVisibili = 2
			end if
		else
			iLivelliVisibili = 2
	end if
		
	For i = 2 to lLastUrowNN
		'Dim oCRA As New com.sun.star.table.CellRangeAddress
		if	oSheet.GetCellByPosition(1 , (I)).CellStyle = "comp Art-EP" & s_R then	 
			 oSheet.GetCellByPosition(1 , (I)).rows.Height = lAltezzaRiga
		end if
	next I
	ThisComponent.CurrentController.Select(oActiveCell1) 
	thisComponent.currentController.Select(thisComponent.CreateInstance("com.sun.star.sheet.SheetCellRanges")) 'unselect ranges 	
	' toglie la selezione 	
END SUB 'fine di: Tronca_Altezza_Voci_Computo
'#########################################################################

Sub Svuota_TUTTO ()
'modificata Giuseppe Vizziello 2015
'svuota tutto (o quasi) il doc di computo ' doradora
	if ThisComponent.Sheets.hasByName("Elenco Prezzi") = false then
		msgbox "Questo comando si usa soltanto in LeenO...",48, "ATTENZIONE!"
		Exit sub
	end if
	Ripristina_statusLine
	barra_apri_chiudi_4

Verifica_chiudi_preview
	if msgbox ("Stai per SVUOTARE questo computo!"& CHR$(10) & CHR$(10)_
			& "SEI DAVVERO SICURO DI VOLER PROSEGUIRE?" & CHR$(10) & " ", 4,"ATTENZIONE!!!") = 7 Then
		exit sub	
	end if
	oDoc = ThisComponent
 ' Get the document's controller.
	oDocCtrl = oDoc.getCurrentController()
 ' Get the frame from the controller.
	oDocFrame = oDocCtrl.getFrame()	
	
' salviamo comunque il doc corrente
'	oDispatchHelper = createUnoService( "com.sun.star.frame.DispatchHelper" )
'	oDispatchHelper.executeDispatch( oDocFrame, ".uno:Save", "", 0, Array() )
 
 	Barra_chiudi_sempre_4
	Barra_Apri_Chiudi_5("Sto svuotando questo documento... Pazienta... ", 30)
rem ----------------------------------------------------------------------
'svuota Elenco Prezzi	
	oSheet = ThisComponent.Sheets.getByName("Elenco Prezzi")
	oEnd=uFindString("Fine elenco", oSheet) 
	lrowFine=oEnd.RangeAddress.EndRow-1
'GoTo salta:
	If lrowFine > 2 Then
		oRangeEP = ThisComponent.NamedRanges.getByName("elenco_prezzi").getReferredCells.RangeAddress
		lrowEPI = oRangeEP.StartRow+1
		lrowEPF = oRangeEP.EndRow-1
		lcolEPI = oRangeEP.StartColumn
		lcolEPF = oRangeEP.EndColumn
		ThisComponent.CurrentController.Select(oSheet.GetCellRangeByPosition(lcolEPI,lrowEPI,lcolEPF,lrowEPF))
		elimina_righe
		thisComponent.currentController.Select(thisComponent.CreateInstance("com.sun.star.sheet.SheetCellRanges")) 'deseleziona
	end If
'salta:

rem ----------------------------------------------------------------------
	Barra_chiudi_sempre_4
	Barra_Apri_Chiudi_5("2 Sto svuotando questo documento... Pazienta... ", 40)


	ThisComponent.Sheets.getByName("S1").GetCellByPosition(7,327).value = 0 
	ThisComponent.Sheets.getByName("S1").GetCellByPosition(7,334).value = 0

	lrowFine= getLastUsedRow(oSheet)
	oCell=oSheet.getCellRangeByPosition(0, 0, 9, lrowFine)
	ThisComponent.CurrentController.Select(oCell)
	Adatta_Altezza_riga
	thisComponent.currentController.Select(thisComponent.CreateInstance("com.sun.star.sheet.SheetCellRanges")) 'deseleziona	
'	Adatta_h_riga_intera_tabella("Elenco Prezzi")
fissa(0,2)

''svuota Analisi
	oSheet = ThisComponent.Sheets.getByName("Analisi di Prezzo")
	lrowFine= getLastUsedRow(oSheet)

	
	lrowFine=oEnd.RangeAddress.EndRow-1
	If lrowFine > 2 then
'		oSheet.rows.removeByIndex (2, lrowFine)
		oRange = osheet.getCellRangeByPosition (0,1,0,lrowFine )
		ThisComponent.CurrentController.Select(oRange)
		elimina_righe
		thisComponent.currentController.Select(thisComponent.CreateInstance("com.sun.star.sheet.SheetCellRanges")) 'deseleziona
	end If
	
	Barra_chiudi_sempre_4
	Barra_Apri_Chiudi_5("4 Sto svuotando questo documento... Pazienta... ", 60)

	' magari adatta le righe (in certe sistuazioni rimaneva una riga rossa molto alta
	lrowFine= getLastUsedRow(oSheet)
	oCell=oSheet.getCellRangeByPosition(0, 0, 10, lrowFine)
	ThisComponent.CurrentController.Select(oCell)
	elimina_righe
	Adatta_Altezza_riga
	
	thisComponent.currentController.Select(thisComponent.CreateInstance("com.sun.star.sheet.SheetCellRanges")) 'deseleziona
fissa(0,2)

rem RIPULISCI COMPUTO
	Barra_chiudi_sempre_4
	Barra_Apri_Chiudi_5("5 Sto svuotando questo documento... Pazienta... ", 70)
	oSheet = ThisComponent.Sheets.getByName("COMPUTO")
inizializza_computo

	oSheet = ThisComponent.Sheets.getByName("CONTABILITA")	
	Svuota_CONTABILITA
'	Svuota_CONTABILITA_esegui 
	
	Adatta_h_riga_intera_tabella("CONTABILITA")
	if	sState = "chiusa" then
		osheet.isVisible = false
	end if '
	
	
Ripristina_Validita_Lista
	'rendi correntexx
Scrivi_Globale
Ripristina_statusLine
Visualizza_normale_esegui
	msgbox "Questo documento è stato in gran parte svuotato dai sui dati!..!"
end sub
