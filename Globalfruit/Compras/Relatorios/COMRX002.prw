#INCLUDE 'TOTVS.CH'
#INCLUDE 'TBICONN.CH'
#INCLUDE 'RPTDEF.CH'


#DEFINE REL_VERT_STD 18
#DEFINE REL_START  65
//#DEFINE REL_END 560 //Paisagem
#DEFINE REL_END 700 //Retrato
#DEFINE REL_RIGHT 820
#DEFINE REL_LEFT 10

/*/{Protheus.doc} COMRX002
Relatorio de conferencia a cega do documento de entrada
@type function
@version
@author Rafael França
@since 09/02/2021
@return return_type, return_description
/*/
User Function COMRX002()

	Local _cPerg    := "COMRX002B"

	Private oPrint
	Private cSubTitle	:= ""
	Private nPag 		:= 0
	Private nLin 		:= 0
	Private oFonte 		:= u_xValFonte(12,,,,"Arial")
	Private oFonteN 	:= u_xValFonte(12,.T.,,,"Arial")
	Private oFonte10 	:= u_xValFonte(10,,,,"Arial")
	Private oFonte10N 	:= u_xValFonte(10,.T.,,,"Arial")

	//If !lLogin
	//PREPARE ENVIRONMENT EMPRESA "01" FILIAL "01" MODULO "FIN"
	//EndIf

	// Cria e abre a tela de pergunta
	ValidPerg( _cPerg )
	If !Pergunte(_cPerg)
		ApMsgStop("Operação cancelada pelo usuário!")
		Return
	EndIf

	FwMsgRun(Nil, { || fProcPDF() }, "Processando", "Emitindo relatorio em PDF..." )

Return

/*/{Protheus.doc} fProcPdf
imprimir relatorio em pdf
@type function
@version
@author Rafael França
@since 09/02/2021
@return return_type, return_description
/*/
Static Function fProcPdf()

	Local nRegAtu	:= 0
	Local nTotReg	:= 0
	Local cDir 		:= Alltrim(MV_PAR05) + "\"
	Local cLote 	:= ""

	Private cTmp1   	:= GetNextAlias()
	Private cDoc 		:= MV_PAR01
	Private cSerie		:= MV_PAR02
	Private cFornece 	:= MV_PAR03
	Private cLoja		:= MV_PAR04
	Private cNomFor		:= "" //Alltrim(Posicione("SA2",1,xfilial("SA2")+cFornece+cLoja,"A2_NOME"))
	Private cProduto 	:= ""
	Private lOk       	:= .T.

	// Query para buscar as informações
	GetData(cDoc,cSerie,cFornece,cLoja)

	// Carrega regua de processamento
	Count To nTotReg
	ProcRegua( nTotReg )

	If nTotReg == 0
		MsgInfo("Não existem registros a serem impressos, favor verificar os parametros","COMRX002")
		(cTmp1)->(DbCloseArea())
		Return
	EndIf

	(cTmp1)->(DbGoTop())

	cFileName 	:= "CONFERENCIA_RECEBIMENTO_"+cDoc+cSerie+cFornece+cLoja+"_COMRX002"
	oPrint := FWMSPrinter():New(cFileName, IMP_PDF, .F., cDir, .T.)
	oPrint:SetPortrait()//Retrato
	//	oPrint:SetLandScape()//Paisagem
	oPrint:SetPaperSize(DMPAPER_A4)
	oPrint:cPathPDF := cDir


	While (cTmp1)->(!Eof())

If lOk
		cNomFor		:= Alltrim((cTmp1)->NOME)
		ImpProxPag()//Monta cabeçario da primeira e proxima pagina
		lOk      	:= .F.
		EndIf

		nRegAtu++
		// Atualiza regua de processamento
		IncProc( "Imprimindo Registro " + cValToChar( nRegAtu ) + " De " + cValToChar( nTotReg ) + " [" + StrZero( Round( ( nRegAtu / nTotReg ) * 100 , 0 ) , 3 ) +"%]" )

		oPrint:Say( nLin,020, (cTmp1)->PRODUTO			  				  ,oFonte)
		oPrint:Say( nLin,070, SUBSTRING((cTmp1)->DESCRICAOP,1,40)		  ,oFonte)
		oPrint:Say( nLin,300, (cTmp1)->UNIDADE				  			  ,oFonte)
		oPrint:Say( nLin,350, "______" 						  			  ,oFonte)
		oPrint:Say( nLin,400, "___________________"		  				  ,oFonte)
		oPrint:Say( nLin,510, "____/____/____"							  ,oFonte)

		nLin += REL_VERT_STD

		If nLin > REL_END
			u_XRODAPE(@oPrint,"COMRX002.PRW","")
			oPrint:EndPage()
			ImpProxPag()//Monta cabeçario da proxima pagina
		EndIf

		cLote   := (cTmp1)->LOTE
		cProduto := (cTmp1)->PRODUTO

		(cTmp1)->(DbSkip())

		If  cLote == "L" .AND. cProduto <> (cTmp1)->PRODUTO

			//Espaço para preencher mais um lote
		oPrint:Say( nLin,350, "______" 						  			,oFonte)
		oPrint:Say( nLin,400, "___________________"		  				,oFonte)
		oPrint:Say( nLin,510, "____/____/____"							,oFonte)

			nLin += REL_VERT_STD

		EndIf

	EndDo

	oPrint:Line(nLin,10,nLin,580,CLR_HGRAY,"-9")

	nLin += (REL_VERT_STD * 3)

	//Imprime linha das assinaturas e observações

	oPrint:Say( nLin,020, "OBSERVACAO: __________________________________________________________________________________________________________________" ,oFonte10N)

	nLin += REL_VERT_STD

	oPrint:Say( nLin,020, "________________________________________________________________________________________________________________________________" ,oFonte10N)

	nLin += (REL_VERT_STD * 2)

	oPrint:Say( nLin,020, "NOME E ASSINATURA DO CONFERENTE: ____________________________________________________________________________________________" ,oFonte10N)

	nLin += (REL_VERT_STD * 2)

	oPrint:Say( nLin,020, "DATA: ____/____/____            HORA: ____:____" ,oFonte10N)

	nLin += REL_VERT_STD

	u_XRODAPE(@oPrint,"COMRX002.PRW","")
	oPrint:EndPage()
	oPrint:Preview()
	(cTmp1)->(DbCloseArea())

Return

/*/{Protheus.doc} ImpProxPag
    Imprime cabeçlho da proxima pagina
    @author  Rafael França
    @since   09/02/2021
/*/

Static Function ImpProxPag()

	nPag++
	oPrint:StartPage()
	cSubTitle := "FORNECEDOR: " + cNomFor
	nLin := u_XCABECA(@oPrint, "CONFERENCIA RECEBIMENTO - NOTA FISCAL: " + cDoc + cSerie , cSubTitle  , nPag)

	oPrint:Say( nLin,020, "PRODUTO",oFonteN)
	oPrint:Say( nLin,070, "DESCRICAO",oFonteN)
	oPrint:Say( nLin,300, "UM",oFonteN)
	oPrint:Say( nLin,350, "QUANT.",oFonteN)
	oPrint:Say( nLin,400, "LOTE",oFonteN)
	oPrint:Say( nLin,510, "VALIDADE",oFonteN)

	oPrint:line(nLin+5,REL_LEFT,nLin+5,REL_RIGHT )

	nLin += REL_VERT_STD

Return

/*/{Protheus.doc} GetData
    Busca dados no banco
    @author  Rafael França
    @since   09/02/2021
/*/

Static Function GetData(cDoc,cSerie,cFornece,cLoja)

cFiltro := "%AND D1_DOC = '" + cDoc + "' AND D1_SERIE = '" + cSerie + "' AND D1_FORNECE = '" + cFornece + "' AND D1_LOJA = '" + cLoja + "' %"

BeginSql Alias cTmp1

//Query com o tratamento das informações da tabela SD1 de entrada
SELECT 'ENTRADA' AS TIPO,D1_FILIAL AS FILIAL,D1_TIPO AS TIPONF
, D1_DOC AS NFISCAL, D1_SERIE AS SERIE, D1_ITEM AS ITEM, D1_COD AS PRODUTO, B1_RASTRO AS LOTE, D1_UM AS UNIDADE
, AH_UMRES AS MEDIDA, B1_DESC AS DESCRICAOP, D1_QUANT AS QUANTIDADE,D1_VUNIT AS VLUNIT, D1_TOTAL AS VLTOTAL, D1_LOCAL AS ARMAZEM
, D1_LOTECTL AS LOTE, D1_DTVALID AS VALIDADE, D1_DTDIGIT AS ENTRADA, D1_EMISSAO AS EMISSAO
, D1_TES AS TES
, D1_FORNECE AS CLIFOR, D1_LOJA AS LOJA
, CASE WHEN D1_TIPO NOT IN ('D','B') THEN (SELECT A2_NOME FROM %table:SA2% WHERE D1_FORNECE = A2_COD AND D1_LOJA = A2_LOJA AND %table:SA2%.D_E_L_E_T_ = '' )
WHEN D1_TIPO IN ('D','B') THEN (SELECT A1_NOME FROM %table:SA1% WHERE D1_FORNECE = A1_COD AND D1_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS NOME
, D1_PEDIDO AS PEDIDO, D1_ITEMPC AS ITEMPD
, CASE D1_TIPO WHEN 'D' THEN 'DEVOLUCAO' WHEN 'N' THEN 'NORMAL' WHEN 'C' THEN 'COMPLEMENTO VALOR' WHEN 'B' THEN 'BENEFICIAMENTO' WHEN 'P' THEN 'COMPLEMENTO IPI' WHEN 'I' THEN 'COMPLEMENTO ICMS' END AS TIPONF1
, F1_CHVNFE AS CHAVENFE
FROM %table:SD1%
INNER JOIN %table:SF1% ON D1_FILIAL = F1_FILIAL AND D1_DOC = F1_DOC AND D1_FORNECE = F1_FORNECE AND D1_LOJA = F1_LOJA AND D1_SERIE = F1_SERIE AND D1_EMISSAO = F1_EMISSAO AND %table:SF1%.D_E_L_E_T_ = ''
INNER JOIN %table:SB1% ON D1_COD = B1_COD AND %table:SB1%.D_E_L_E_T_ = ''
//INNER JOIN %table:SF4% ON D1_TES = F4_CODIGO AND %table:SF4%.D_E_L_E_T_ = ''
//INNER JOIN %table:NNR% ON D1_FILIAL = NNR_FILIAL AND D1_LOCAL = NNR_CODIGO AND %table:NNR%.D_E_L_E_T_ = '' //GF NNR Exclusiva
//INNER JOIN %table:NNR% ON D1_LOCAL = NNR_CODIGO AND %table:NNR%.D_E_L_E_T_ = '' //Rec NNR Compartilhada
INNER JOIN %table:SAH% ON D1_UM = AH_UNIMED AND %table:SAH%.D_E_L_E_T_ = ''
WHERE %table:SD1%.D_E_L_E_T_ = ''
%exp:cFiltro%

EndSql //FINALIZO A MINHA QUERY

Return

/*/{Protheus.doc} ValidPerg
//TODO Funï¿½ï¿½o que cria as perguntas.
@author Eduardo Cevoli
@since 01/06/2020
@version 1.0
@return ${return}, ${return_description}
@param cPerg, characters, descricao
@type function
/*/

Static Function ValidPerg(cPerg)

	Local aArea	:= GetArea()
	Local aRegs	:= {}
	Local i,j

	DbSelectArea("SX1")
	SX1->(DbSetOrder(1))

	cPerg := PADR(cPerg,10)

	//          Grupo Ordem Desc Por               Desc Espa   Desc Ingl  Variavel  Tipo  Tamanho  Decimal  PreSel  GSC  Valid   Var01       Def01     DefSpa01  DefEng01  CNT01  Var02  Def02     DefSpa02  DefEng02  CNT02  Var03  Def03  DefEsp03  DefEng03  CNT03     Var04  Def04  DefEsp04  DefEng04  CNT04  Var05  Def05  DefEsp05  DefEng05  CNT05  F3        PYME  GRPSXG   HELP  PICTURE  IDFIL
	aAdd(aRegs,{cPerg,"01", "Nota Fiscal"		 , "",         "",        "mv_ch1", "C",  09,      00,      0,      "G", "",     "mv_par01", "",       "",       "",       "",    "",    "",       "",       "",       "",    "",    "",    "",       "",       "",       "",    "",    "",       "",       "",    "",    "",    "",       "",       "",    "SF101",       "",   "",      "",   "",      ""   })
	aAdd(aRegs,{cPerg,"02", "Serie"				 , "",         "",        "mv_ch2", "C",  03,      00,      0,      "G", "",     "mv_par02", "",       "",       "",       "",    "",    "",       "",       "",       "",    "",    "",    "",       "",       "",       "",    "",    "",       "",       "",    "",    "",    "",       "",       "",    "",       "",   "",      "",   "",      ""   })
	aAdd(aRegs,{cPerg,"03", "Fornecedor"		 , "",         "",        "mv_ch3", "C",  06,      00,      0,      "G", "",     "mv_par03", "",       "",       "",       "",    "",    "",       "",       "",       "",    "",    "",    "",       "",       "",       "",    "",    "",       "",       "",    "",    "",    "",       "",       "",    "SA2",       "",   "",      "",   "",      ""   })
	aAdd(aRegs,{cPerg,"04", "Loja"				 , "",         "",        "mv_ch4", "C",  02,      00,      0,      "G", "",     "mv_par04", "",       "",       "",       "",    "",    "",       "",       "",       "",    "",    "",    "",       "",       "",       "",    "",    "",       "",       "",    "",    "",    "",       "",       "",    "",       "",   "",      "",   "",      ""   })
	aAdd(aRegs,{cPerg,"05", "Destino do(s) Arq.?", "",         "",        "mv_ch5", "C",  99,      00,      0,      "G", "",     "mv_par05", "",       "",       "",       "",    "",    "",       "",       "",       "",    "",    "",    "",       "",       "",       "",    "",    "",       "",       "",    "",    "",    "",       "",       "",    "",       "",   "",      "",   "",      ""   })

	For i:=1 to Len(aRegs)
		If !dbSeek(PADR(cPerg,10)+aRegs[i,2])
			RecLock("SX1",.T.)
			For j:=1 to FCount()
				If j <= Len(aRegs[i])
					FieldPut(j,aRegs[i,j])
				Endif
			Next
			MsUnlock()
		Endif
	Next

	RestArea(aArea)

Return()

/*/{Protheus.doc} xValFonte
	DescriÃ§Ã£o: Realiza o encapsulamento da funÃ§Ã£o TFONT
	@author    Rafael França
	@version   1.00
	@since     09/02/2021
/*/

User Function xValFonte(nTam,lBold,lLine,lItalic,cFont)

	Local oFonte

	lBold 	:= If ( lBold	==	Nil	,	.F.	, 	lBold	)
	lLine	:= If ( lLine	==	Nil	,	.F.	,	lLine	)
	lItalic	:= If ( lItalic	==	Nil	,	.F.	,	lItalic )
	cFont	:= If ( cFont	==	Nil	,"Arial"	,	cFont )

	oFonte:=TFont():New( cFont,,nTam,,lBold,,,,,lLine,lItalic)

Return oFonte

/*/{Protheus.doc} XCABECA
	Monta um cabeÃ§alho prÃ©-definido de acordo com a orientaÃ§Ã£o do objeto
	@author    Rafael França
	@version   1.0
	@since     09/02/2021
/*/

User Function XCABECA(oPrint,cTitle,cSubTitle,nPage, lBlackWhite)

	Local oFont24 := u_xValFonte(16,,,,"Arial")
	Local oFont14 := u_xValFonte(14,,,,"Arial")
	Local cData := DTOS(DATE()) //DTOC NÃ£o estÃ¡ funcionando
	Default cSubTitle := ""

	cData := SUBSTRING(cData,7,2) + "/" +  SUBSTRING(cData,5,2)  + "/" + SUBSTRING(cData,1,4)

	If oPrint:GetOrientation() == 1

		oPrint:SayBitmap(30,15,"\system\LGMID10.png",80,40)
		oPrint:SayAlign(35,155,Capital(AllTrim(Posicione("SM0", 1, cEmpAnt+cFilAnt , "M0_NOMECOM"))),oFont14, 580, 20, , 0, 2)
		oPrint:SayAlign(51,155, "ERP | " + oApp:cModDesc ,oFont14,580,20,,0,2)

		oPrint:SayAlign(20,20,"Emitido em: " + cData,oFont14, 535, 20, , 1, 1)
		oPrint:SayAlign(40,20,"Hora: " + Time(),oFont14, 535, 20, , 1, 1)
		oPrint:SayAlign(60,20,"Pagina: " + cValtoChar(nPage),oFont14, 535, 20, , 1, 1)

		oPrint:Line(80,10,80,580,CLR_HGRAY,"-9")

		oPrint:SayAlign(80,010,cTitle,oFont24, 580, 20, /*[ nClrText]*/, 2, 1)
		oPrint:SayAlign(100,010,cSubTitle,oFont24, 580, 20, /*[ nClrText]*/, 2, 1)

		oPrint:Line(118,10,118,580,CLR_HGRAY,"-9")

	Else

		oPrint:Line(20,10,20,820,CLR_HGRAY,"-9")
		oPrint:SayBitmap(30,15,"\system\LGMID10.PNG",80,40)

		oPrint:SayAlign(70,18,Capital(AllTrim(Posicione("SM0", 1, cEmpAnt+cFilAnt , "M0_NOMECOM"))),oFont14, 820, 20, /*[ nClrText]*/, 0, 1)
		oPrint:SayAlign(85,18, "ERP | " + oApp:cModDesc ,oFont14,580,20,,0,2)

		oPrint:SayAlign(30,20,"Emitido em: " + cData,oFont14, 800, 20, /*[ nClrText]*/, 1, 1)
		oPrint:SayAlign(50,20,"Hora: " + Time(),oFont14, 800, 20, /*[ nClrText]*/, 1, 1)
		oPrint:SayAlign(70,20,"Pagina: " + cValtoChar(nPage),oFont14, 800, 20, /*[ nClrText]*/, 1, 1)

		oPrint:SayAlign(40,10,cTitle,oFont24, 830, 20, /*[ nClrText]*/, 2, 1)
		oPrint:SayAlign(65,10,cSubTitle,oFont14, 830, 20, /*[ nClrText]*/, 2, 1)
		oPrint:Line(105,10,105,820,CLR_HGRAY,"-9")
	EndIf

Return 130


/*/{Protheus.doc} XRODAPE
	Monta um rodapÃ© prÃ©-definido de acordo com a orientaÃ§Ã£o do objeto
	@author    Rafael França
	@version   1.0
	@since     09/02/2021
/*/

User Function XRODAPE(oPrint,cFonteBase,cMsgPad)

	Local oFont8 := u_xValFonte(8)

	cMsgPad := If(cMsgPad == Nil,"",AllTrim(cMsgPad) + " ")

	If oPrint:GetOrientation() == 1
		oPrint:Box (815, 10, 830, 580, "-4")
		oPrint:SayAlign(819,20,cMsgPad + u_xInspFonte(cFonteBase),oFont8, 555, 20, /*[ nClrText]*/, 1, 1)
	Else
		oPrint:Box (580, 10, 595, 830, "-4")
		oPrint:SayAlign(584,20,cMsgPad + u_xInspFonte(cFonteBase),oFont8, 805, 20, /*[ nClrText]*/, 1, 1)
	EndIf

Return


User function xInspFonte(cFonte)

	Local cRet			:= ""
	Local aData			:= {}
	Local cData         := DTOS(DATE()) //DTOC NÃ£o estÃ¡ funcionando

	Default __cUserId	:= ""

	cData := SUBSTRING(cData,7,2) + "/" +  SUBSTRING(cData,5,2)  + "/" + SUBSTRING(cData,1,4)

	U_XVERSAO()

	aData := GetAPOInfo(cFonte)//aFontes[nI])

	/*
    Modos de compilaÃ§Ã£o:
    Valor                     DescriÃ§Ã£o
    0 - BUILD_FULL            UsuÃ¡rio tem permissÃ£o para compilar qualquer tipo de fonte
    2 - BUILD_PARTNER         PermissÃ£o de compilaÃ§Ã£o da FÃ¡brica de Software TOTVS
    3 - BUILD_PATCH           AplicaÃ§Ã£o de Patch
    1 - BUILD_USER            UsuÃ¡rio sÃ³ pode compilar User Functions
	*/

	//cRet := aData[1] + "(" + dtoc(aData[4]) + " - " +  aData[5] + ")"
	cRet := aData[1] + "(" + cData + " - " +  aData[5] + ")"
	cRet += " ENV " + AllTrim(GetEnvServer())
	cRet += " VER " + cXVERSAO
	cRet += " USR " + __cUserId
	//cRet += " EMITIDO " + dtoc(DATE()) + " - " + TIME()
	cRet += " EMITIDO " + cData + " - " + TIME()

Return cRet

/*/{Protheus.doc} XVERSAO
	Cria uma variÃ¡vel pÃºblica baseada no CHANGELOG. MD
	@author  Rafael França
	@since   09/02/2021
/*/

User Function XVERSAO()

	Local cChangeLog	:= cValToChar(GetApoRes("CHANGELOG.MD"))
	Local nIni			:= At( "## [", cChangeLog) + 4//remove os caracteres procurados
	Local nFim			:= At( "]", cChangeLog,nIni)
	Public cXVERSAO	:= ""

	If nIni > 0 .AND. nFim > 0
		cXVERSAO	:= SubStr(cChangeLog,nIni,nFim-nIni)
	EndIf

Return cXVERSAO
