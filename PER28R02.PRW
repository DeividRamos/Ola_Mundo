#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOPCONN.CH"
#INCLUDE "RWMAKE.CH"
/* testes
*/

/*=========================================================================*\
|| ####################################################################### ||
|| # Sistema:  PROGRAMA PER28R02                                         # ||
|| #                                                                     # ||
|| # Objetivo: Impressão customizada de Orçamento (OS) MOD2 para         # ||
|| #           orçamento de atendimento/conserto.  Formato PDF           # ||
|| #                                                                     # ||
|| # Autor:    FILIPE OLEGINI                                            # ||
|| # Data:     29/01/2015                                                # ||
|| ####################################################################### ||
\*=========================================================================*/

User Function PER28R02()
	//define objeto de impressao utilizado em todo relatorio
	private oPrint
		
	//variaveis de controle de linhas
	Private nColuna     := 80
	Private nLinha      := 300
	Private nSaltoLinha := 50
	
	Private nTotal   := 0
	Private nTotIPI  := 0
	Private nTotCIPI := 0
	Private nItem    := 0
	
	CriaPer()
		
	oPrint  := FWMSPrinter():New("Orcamento_" + AllTrim(AB3->AB3_NUMORC) + ".pdf",6,.T.,,.T.)
  
	oPrint:SetPortrait()    //seta impressao paisagem
	oPrint:SetPaperSize(9)  //seta tamanho papel
	oPrint:cPathPDF := ALLTRIM(GETTEMPPATH())
	oPrint:SetViewPDF(.T.)
	
	If !Pergunte("PER28R02",.T.)
      MsgAlert("Textos não imformados")
   EndIf
	
	Processa({||ImpRel()}, "Gerando Orçamento")
		
	//SET Filter TO
	//oPrint:Preview()
	Ms_Flush()
	
Return

//------------------------------------------------
// busca os dados e monta as partes do relatorio
//------------------------------------------------
Static Function ImpRel()
	
	//define as fontes  
	Private oFont10  := TFont():New( "Arial",,10,,.F.,,,,,.F. )
	Private oFont12  := TFont():New( "Arial",,14,,.F.,,,,,.F. )	
	Private oFont18  := TFont():New( "Arial",,18,,.F.,,,,,.F. )
			
	cAliasAB3 := GetNextAlias() 
	cQuery := " SELECT * "
	cQuery += " FROM " + RetSqlName("AB3") + " B3 "
	cQuery += " INNER JOIN " + RetSqlName("AB4") + " B4 ON (B3.AB3_FILIAL = B4.AB4_FILIAL AND B3.AB3_NUMORC = B4.AB4_NUMORC "
	cQuery +=                                             " AND  B3.D_E_L_E_T_ = B4.D_E_L_E_T_) "
	//cQuery += " INNER JOIN " + RetSqlName("AA3") + " A3 ON (SUBSTRING(B3.AB3_FILIAL,1,2) = A3.AA3_FILIAL AND B3.AB3_CODCLI = A3.AA3_CODCLI AND B3.AB3_LOJA = A3.AA3_LOJA AND "
	cQuery += " INNER JOIN " + RetSqlName("AA3") + " A3 ON (B3.AB3_CODCLI = A3.AA3_CODCLI AND B3.AB3_LOJA = A3.AA3_LOJA AND "
	cQuery +=                                        " B4.AB4_CODPRO = A3.AA3_CODPRO AND B4.AB4_NUMSER = A3.AA3_NUMSER AND B3.D_E_L_E_T_ = A3.D_E_L_E_T_ ) "
	cQuery += " WHERE B3.AB3_FILIAL = '" + AB3->AB3_FILIAL+ "'"
	cQuery += " AND B3.AB3_NUMORC = '" + AB3->AB3_NUMORC + "'"
	cQuery += " AND B3.D_E_L_E_T_ <> '*'"
	
	cQuery := ChangeQuery(cQuery)
	
	dbUseArea(.T., "TOPCONN", TCGENQRY(, , cQuery), cAliasAB3, .F., .T.)
	dbSelectArea(cAliasAB3)
	
	If (cAliasAB3)->(!EOF())
		
		dbSelectArea("SA1")
		dbSetOrder(1)
		dbSeek(xFilial("SA1") + (cAliasAB3)->AB3_CODCLI + (cAliasAB3)->AB3_LOJA)
		
		nItem := 0
		
		While (cAliasAB3)->(!EOF())
			
			oPrint:StartPage()
			
			nItem ++
			
			RCab()
			
			If StoD((cAliasAB3)->AA3_DTGAR) < dDataBase
				RGItens(.F.)
			Else
				RGItens(.T.)
			EndIf
			
		   //RLItens()
			RCItens()
			
			nTotal   := 0
	      nTotIPI  := 0
	      nTotCIPI := 0
			
			cAliasAB5 := GetNextAlias() 
	      cQuery := " SELECT B5.AB5_CODPRO, B1.B1_DESC, B5.AB5_QUANT, B5.AB5_VUNIT, B5.AB5_TOTAL, B1.B1_IPI  "
	      cQuery += " FROM " + RetSqlName("AB5") + " B5 "
	      cQuery += " INNER JOIN " + RetSqlName("AA5") + " A5 ON (SUBSTRING(B5.AB5_FILIAL,1,2) = A5.AA5_FILIAL AND B5.AB5_CODSER = A5.AA5_CODSER AND B5.D_E_L_E_T_ = A5.D_E_L_E_T_) "
	      cQuery += " INNER JOIN " + RetSqlName("SB1") + " B1 ON (B5.AB5_CODPRO = B1.B1_COD AND B5.D_E_L_E_T_ = B1.D_E_L_E_T_)  "
	      cQuery += " WHERE B5.D_E_L_E_T_<> '*' "
	      cQuery += " AND B5.AB5_FILIAL = '" + (cAliasAB3)->AB3_FILIAL + "'"
	      cQuery += " AND B5.AB5_NUMORC = '" + (cAliasAB3)->AB3_NUMORC + "'"
	      cQuery += " ORDER BY B5.AB5_ITEM "
	      
	      cQuery := ChangeQuery(cQuery)
	
			dbUseArea(.T., "TOPCONN", TCGENQRY(, , cQuery), cAliasAB5, .F., .T.)
			dbSelectArea(cAliasAB5)
			
			If (cAliasAB5)->(!EOF())
				
				While (cAliasAB5)->(!EOF())
					
				 	RItens()
				 	
				 	nTotal   += (cAliasAB5)-> AB5_TOTAL
	      		nTotIPI  += ((cAliasAB5)->AB5_TOTAL * (cAliasAB5)->B1_IPI) / 100
	      		
					
					(cAliasAB5)->(dbSkip())
				EndDo
			Else
				dbCloseArea((cAliasAB5))	
			EndIf
			nTotCIPI += nTotal + nTotIPI
			
			RLItens()
			
			RTItens()
			
			//posiciona a linha no final
			nLinha := 1400			
			RRod()
			
			
			oPrint:EndPage()
	  		(cAliasAB3)->(dbSkip())			
		EndDo
	
	Else
	   dbCloseArea((cAliasAB3))
	EndIf
	
	oPrint:Print()  
	
Return

Static Function RCab()
	
	//imagem esquerda
   oPrint:SayBitmap(50, 80, "PER28R01_1.bmp", 350, 180)
   
   //imagem direita
   oPrint:SayBitmap(50, 2000, "PER28R01_2.bmp", 200, 180)
	
	//linha
	oPrint:Line(265, 80, 265, 2300)
	
	oPrint:Say(nLinha, nColuna, SA1->A1_NOME, oFont12)
	oPrint:Say(nLinha, nColuna + 2000, DtoC(StoD((cAliasAB3)->AB3_EMISSA)), oFont12)
	nLinha += nSaltoLinha
	
	oPrint:Say(nLinha, nColuna, "CNPJ: " + Transform(SA1->A1_CGC,"@R 99.999.999/9999-99"), oFont12)
	nLinha += nSaltoLinha
	
	oPrint:Say(nLinha, nColuna, AllTrim(SA1->A1_END), oFont12)
	nLinha += nSaltoLinha
	
	oPrint:Say(nLinha, nColuna, AllTrim(SA1->A1_MUN) + " - " + AllTrim(SA1->A1_EST) + " - " + " - BRASIL", oFont12)
	nLinha += nSaltoLinha
	
	oPrint:Say(nLinha, nColuna, "CEP:" + Transform(SA1->A1_CGC,"@R 99999-999"), oFont12)
	nLinha += nSaltoLinha
	
	oPrint:Say(nLinha, nColuna, SA1->A1_CONTATO, oFont12)
	nLinha += ( nSaltoLinha * 2 )
	
	oPrint:Say(nLinha, nColuna, "Ref.: Cotação Nº " + AllTrim((cAliasAB3)->AB3_NUMORC) + " de " + DtoC(StoD((cAliasAB3)->AB3_EMISSA)) , oFont12)
	nLinha += ( nSaltoLinha * 2 )	
	
	oPrint:Say(nLinha, nColuna, "Prezados(as) senhores(as), ", oFont12)
	nLinha += nSaltoLinha
	
	oPrint:Say(nLinha, nColuna, "Estamos enviando a(s) cotação(ões) referente ao item solicitado para apreciação e aprovação." , oFont12)
	nLinha += nSaltoLinha
		
Return

//monta o cabeçalho dos itens fora de garantia
Static Function RGItens(lGarantia)
	nLinha += nSaltoLinha
	
	oPrint:Line(nLinha - 40, 80, nLinha  - 40, 2300) //linha
	
	//imprime a colunas
	oPrint:Line(nLinha - 40, 80, nLinha + 10 , 80) //coluna 1
	oPrint:Line(nLinha - 40, 2300, nLinha + 10 , 2300) //coluna 7
	
	//imprime os dados 
	If lGarantia
		oPrint:Say(nLinha, 1010, "ITENS EM GARANTIA", oFont12)
	Else
		oPrint:Say(nLinha, 960, "ITENS FORA DA GARANTIA", oFont12)
	EndIf
	
	nLinha += nSaltoLinha 
		
	oPrint:Line(nLinha - 40, 80, nLinha - 40, 2300) //linha
	
	//imprime os dados 
	oPrint:Say(nLinha, 90, "ITEM", oFont12)
	oPrint:Say(nLinha, 200, "Nº SÉRIE", oFont12)
	oPrint:Say(nLinha, 600, "MODELO", oFont12)
	oPrint:Say(nLinha, 1000, "DESCRIÇÃO", oFont12)
	oPrint:Say(nLinha, 2000, "QTDADE", oFont12)
	nLinha += nSaltoLinha
	
	//imprime os dados 
	oPrint:Say(nLinha, 90, Alltrim(Str(nItem)) + ")", oFont12)
	oPrint:Say(nLinha, 200, AllTrim((cAliasAB3)->AA3_NUMSER), oFont12)
	oPrint:Say(nLinha, 600, AllTrim((cAliasAB3)->AA3_CODPRO), oFont12)
	oPrint:Say(nLinha, 1000, AllTrim(Posicione("SB1",1,xFilial("SB1") + (cAliasAB3)->AA3_CODPRO,"B1_DESC")), oFont12)
	//oPrint:Say(nLinha, 2000, "1" + AllTrim(Posicione("SB1",1,xFilial("SB1") + (cAliasAB3)->AA3_CODPRO,"B1_UM")), oFont12)
	oPrint:Say(nLinha, 2000, "01", oFont12)
	nLinha += nSaltoLinha + 50
	
Return

//monta o cabeçalho dos itens
Static Function RLItens()
	
	oPrint:Line(nLinha - 40, 80, nLinha  - 40, 2300) //linha
	
Return  

//monta o cabeçalho dos itens
Static Function RCItens()
	
	oPrint:Say(nLinha, nColuna + 1100, "ITENS" , oFont18)
	nLinha += nSaltoLinha
	
	oPrint:Line(nLinha - 40, 80, nLinha  - 40, 2300) //linha
	
		//imprime a colunas
	oPrint:Line(nLinha - 40, 80, nLinha + 10 , 80) //coluna 1
	oPrint:Line(nLinha - 40, 350, nLinha  + 10 , 350) //coluna 2
	oPrint:Line(nLinha - 40, 1200, nLinha + 10 , 1200) //coluna 3
	oPrint:Line(nLinha - 40, 1450, nLinha + 10 , 1450) //coluna 4
	oPrint:Line(nLinha - 40, 1750, nLinha + 10 , 1750) //coluna 5
	oPrint:Line(nLinha - 40, 2050, nLinha + 10 , 2050) //coluna 6
	oPrint:Line(nLinha - 40, 2200, nLinha + 10 , 2200) //coluna 7
	oPrint:Line(nLinha - 40, 2300, nLinha + 10 , 2300) //coluna 7
	
	//imprime os dados 
	oPrint:Say(nLinha, 90, "CÓDIGO", oFont12)
	oPrint:Say(nLinha, 360, "DESCRIÇÃO", oFont12)
	oPrint:Say(nLinha, 1210, "QTDADE", oFont12)
	oPrint:Say(nLinha, 1480, "VLR. UNITÁRIO", oFont12)
	oPrint:Say(nLinha, 1800, "VLR. TOTAL", oFont12)
	oPrint:Say(nLinha, 2060, "(%) IPI", oFont12)
	oPrint:Say(nLinha, 2210, "ST", oFont12)
	
	oPrint:Line(nLinha - 40, 80, nLinha - 40, 2300) //linha
	 
	nLinha += nSaltoLinha
	
Return

//imprime as linhas
Static Function RItens()
	
	//imprime a colunas
	oPrint:Line(nLinha - 40, 80, nLinha + 10 , 80) //coluna 1
	oPrint:Line(nLinha - 40, 350, nLinha  + 10 , 350) //coluna 2
	oPrint:Line(nLinha - 40, 1200, nLinha + 10 , 1200) //coluna 3
	oPrint:Line(nLinha - 40, 1450, nLinha + 10 , 1450) //coluna 4
	oPrint:Line(nLinha - 40, 1750, nLinha + 10 , 1750) //coluna 5
	oPrint:Line(nLinha - 40, 2050, nLinha + 10 , 2050) //coluna 6
	oPrint:Line(nLinha - 40, 2200, nLinha + 10 , 2200) //coluna 7
	oPrint:Line(nLinha - 40, 2300, nLinha + 10 , 2300) //coluna 7
	
	//imprime os dados 
	oPrint:Say(nLinha, 90, (cAliasAB5)->AB5_CODPRO, oFont12)
	oPrint:Say(nLinha, 360, AllTrim((cAliasAB5)->B1_DESC), oFont12)
	oPrint:Say(nLinha, nColuna + 1210, Transform((cAliasAB5)->AB5_QUANT,"@E 999,999.99"), oFont12)
	oPrint:Say(nLinha, nColuna + 1480, Transform((cAliasAB5)->AB5_VUNIT,"@E 999,999.99"), oFont12)
	oPrint:Say(nLinha, nColuna + 1800, Transform((cAliasAB5)->AB5_TOTAL,"@E 999,999.99"), oFont12)
	oPrint:Say(nLinha, nColuna + 2010, Transform((cAliasAB5)->B1_IPI,"@E 999.99"), oFont12)
	
	oPrint:Line(nLinha - 40, 80, nLinha - 40, 2300) //linha

	nLinha += nSaltoLinha

Return

//imprime as linhas
Static Function RTItens()     

	oPrint:Line(nLinha - 40, 1450, nLinha + 10 , 1450) //coluna 4
	oPrint:Line(nLinha - 40, 2050, nLinha + 10 , 2050) //coluna 6
	oPrint:Line(nLinha - 40, 2300, nLinha + 10 , 2300) //coluna 7
	
	//imprime os dados 
	oPrint:Say(nLinha, 1500, "VLR. TOTAL", oFont12)
	oPrint:Say(nLinha, 2060, AllTrim(Transform(nTotal,"@E 999,999.99")), oFont12)
	
	oPrint:Line(nLinha - 40, 1450, nLinha - 40, 2300) //linha

	nLinha += nSaltoLinha
	
	oPrint:Line(nLinha - 40, 1450, nLinha + 10 , 1450) //coluna 4
	oPrint:Line(nLinha - 40, 2050, nLinha + 10 , 2050) //coluna 6
	oPrint:Line(nLinha - 40, 2300, nLinha + 10 , 2300) //coluna 7
	
	//imprime os dados 
	oPrint:Say(nLinha, 1500, "VLR. IPI", oFont12)
	oPrint:Say(nLinha, 2060, AllTrim(Transform(nTotIPI,"@E 999,999.99")), oFont12)
	
	oPrint:Line(nLinha - 40, 1450, nLinha - 40, 2300) //linha

	nLinha += nSaltoLinha
	
	oPrint:Line(nLinha - 40, 1450, nLinha + 10 , 1450) //coluna 4
	oPrint:Line(nLinha - 40, 2050, nLinha + 10 , 2050) //coluna 6
	oPrint:Line(nLinha - 40, 2300, nLinha + 10 , 2300) //coluna 7
	
	//imprime os dados 
	oPrint:Say(nLinha, 1500, "VLR. TOTAL COM IPI", oFont12)
	oPrint:Say(nLinha, 2060, AllTrim(Transform(nTotCIPI,"@E 999,999.99")), oFont12)
	
	oPrint:Line(nLinha - 40, 1450, nLinha - 40, 2300) //linha
	
	oPrint:Line(nLinha+10, 1450, nLinha+10, 2300) //linha

	nLinha += nSaltoLinha
	
Return

Static Function RRod()
	
	oPrint:Say(nLinha, nColuna, "Impostos inclusos: ICMS.", oFont12)
	nLinha += nSaltoLinha
	
	oPrint:Say(nLinha, nColuna, "Impostos a incluir: IPI a incluir sobre peças, embalagens e frete.", oFont12)
	nLinha += nSaltoLinha
	
	oPrint:Say(nLinha, nColuna, "Valores expressos: " + AllTrim(mv_par01), oFont12)
	nLinha += nSaltoLinha
	
	oPrint:Say(nLinha, nColuna, "Frete: " + AllTrim(mv_par02), oFont12)
	nLinha += nSaltoLinha
	
	oPrint:Say(nLinha, nColuna, "Pagamento: " + AllTrim(mv_par03), oFont12)
	nLinha += nSaltoLinha
	
	oPrint:Say(nLinha, nColuna, "Faturamento mínimo: R$ 400,00.", oFont12)
	nLinha += nSaltoLinha
	
	oPrint:Say(nLinha, nColuna, "Prazo de devolução previsto após aprovação: " + AllTrim(mv_par04)  , oFont12)
	nLinha += nSaltoLinha	
	
	oPrint:Say(nLinha, nColuna, "Garantia: 90 dias corridos sobre defeitos em peças e serviços aplicados." , oFont12)
	nLinha += nSaltoLinha
	
	oPrint:Say(nLinha, nColuna, "Validade da proposta: 15 dias corridos." , oFont12)
 	nLinha += nSaltoLinha	
	
	oPrint:Say(nLinha, nColuna, "Observações:" , oFont12)
 	nLinha += nSaltoLinha 
 
	oPrint:Say(nLinha, nColuna+50, "a. Horário de recepção/expedição de peças e equipamentos: 08h30min as 17h00min." , oFont12)
 	nLinha += nSaltoLinha
 
	oPrint:Say(nLinha, nColuna+50, "b. Aprovação do orçamento poderá ser feita por fax ou e-mail." , oFont12)
 	nLinha += nSaltoLinha
 
	oPrint:Say(nLinha, nColuna+50, "c. Caso seja necessário a emissão de pedido de compra, favor informar no ato da aprovação." , oFont12)
 	nLinha += nSaltoLinha
 
	oPrint:Say(nLinha, nColuna+50, "d. Na falta de manifestação de aprovação ou reprovação do orçamento de manutenção dos equipamentos enviados para laboratório" , oFont12)
 	nLinha += nSaltoLinha
 	
 	oPrint:Say(nLinha, nColuna+50, "    pelo período superior a 15 (quinze) dias corridos a partir da data de envio do orçamento, o material será devolvido ao " , oFont12)
 	nLinha += nSaltoLinha
 	
 	oPrint:Say(nLinha, nColuna+50, "    cliente com frete à cobrar.", oFont12)
 	nLinha += nSaltoLinha
 	
	oPrint:Say(nLinha, nColuna+50, "e. Os valores totais apresentados na cotação em anexo não consideram ICMS/ST, eventualmente incidentes, esses valores são " , oFont12)
 	nLinha += nSaltoLinha
 	 
	oPrint:Say(nLinha, nColuna+50, "    antecipação de  impostos que serão repassados conforme determina a legislação e protocolo de cada estado." , oFont12)
 	nLinha += (nSaltoLinha * 2)
	
 	oPrint:Say(nLinha, nColuna, "Atenciosamente" , oFont12)
 	nLinha += (nSaltoLinha * 2)
 	
 	oPrint:Say(nLinha, nColuna, "Marcos Eduardo" , oFont12)

  //linha
	oPrint:Line(nLinha+10, 80, nLinha+10, 2300)
 	nLinha += nSaltoLinha
 	
 	oPrint:Say(nLinha, 800, "DIGICON S.A. - Controle Eletrônico para Mecânica " , oFont12)
 	nLinha += nSaltoLinha
 		 	
 	oPrint:Say(nLinha, 320, "SEDE / FABRICA: Rua Nissin Castiel, 640 - Cx. Postal 131 - CEP 94045-420 - Gravataí - RS - Brasil - Fone: 051 3489.8700 - Fax: 051 3489.1110" , oFont10)
 	nLinha += nSaltoLinha
 	
 	//oPrint:Say(nLinha, nColuna+850, "CEP 94000-970 - Fone: (51) 3489-8700" , oFont10)
 	//nLinha += nSaltoLinha
 	
 	oPrint:Say(nLinha, 400, "FILIAL SP: Rua São Paulo, 82 - Alphaville - CEP 06465-130 - Barueri - SP - Brasil - Fone: 011 3738.3500 - Fax: 011 3738.3500" , oFont10)
 	nLinha += nSaltoLinha
 	
 	//oPrint:Say(nLinha, nColuna+850, "CEP 06465-130 - Fone: (11) 4133-4100" , oFont10)
 	//nLinha += nSaltoLinha
 	
 	oPrint:Say(nLinha, nColuna+980, "www.digicon.com.br" , oFont10)
 	nLinha += nSaltoLinha
Return

//inclusao da perguntas
Static Function CriaPer()
	Local cPerg := "PER28R02"
	
	//cria perguntas
   PutSx1(cPerg,"01","Valores Expressos","","","mv_ch1","C",99,0,1,"C","","","","","mv_par01","","","","","","","","","","","","","","","","",{"",""},{"",""},{"",""})
   PutSx1(cPerg,"02","Frete"            ,"","","mv_ch2","C",99,0,1,"C","","","","","mv_par02","","","","","","","","","","","","","","","","",{"",""},{"",""},{"",""})
   PutSx1(cPerg,"03","Pagamento"        ,"","","mv_ch3","C",99,0,1,"C","","","","","mv_par03","","","","","","","","","","","","","","","","",{"",""},{"",""},{"",""})
   PutSx1(cPerg,"04","Prazo de entrega" ,"","","mv_ch4","C",99,0,1,"C","","","","","mv_par04","","","","","","","","","","","","","","","","",{"",""},{"",""},{"",""})

Return
