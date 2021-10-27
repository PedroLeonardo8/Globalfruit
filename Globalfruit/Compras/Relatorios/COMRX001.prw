#Include "RwMake.ch"
#Include "topconn.ch"

// COMRX001 - Rafael França - 29/06/2020 - Programa criado com objetivo de facilitar a importação das informações (Entradas, Saidas e Internas) do sistema para o BI e analises dos usuarios.
//Rafael França - 09/2020 - Alteração do programa para nova linguagem de exportação para excel Totvs.
//Rafael França - 12/02/2021 - Criação do parametros de fornecedor, cliente, produto e TES para facilitar a conferencia de dados do usuario. Pedido Leomar.
//Rafael França - 03/2021 - Colocado dados de impostos e custos no relatorio.
//Rafael França - 07/04/2021 - Informações adicionais de saida vendedor e região pedido Everson, dados adicionais de cliente fornecedor loja e cnpj pedido Leomar.
//Rafael França - 04/05/2021 - Ordem dos campos de saida de acordo com o solicitação Karoline/Everson.
//Rafael França - 06/05/2021 - Ordem dos campos seguindo Empresa, Filial, Entrada, Nota Fiscal, Serie e Item.

User Function COMRX001

Private cAlias 			:= GetNextAlias() //Declarei meu ALIAS
Private aArea        	:= GetArea()
Private oFWMsExcel
Private oExcel
Private cArquivo    	:= GetTempPath()+'COMRX001.xml'
Private cPerg  			:= "COMRX001"
Private cFiltro			:= ""
Private cNomeTabela		:= ""
Private cCampoImp		:= "%%"

	ValidPerg(cPerg) //INICIA A STATIC FUNCTION PARA CRIAÇÃO DAS PERGUNTAS

	If !Pergunte(cPerg) //Verifica se usuario respondeu as perguntas ou cancelou
		MsgAlert("Operação Cancelada!","Alerta!")
		RestArea(aArea)
		Return
	EndIf

If !ApOleClient("MsExcel")
	MsgStop("Microsoft Excel nao instalado.")  //"Microsoft Excel nao instalado."
	RestArea(aArea)
	Return
EndIf

FwMsgRun(Nil, { || COMRX001A() }, "Processando", "Gerando planilha xml..." )

Return

Static Function COMRX001A

	Local aUsuario	:= {}
	Local cUsuario	:= ""
	//Local cData  	:= ""

If MV_PAR05 == 1 .OR. MV_PAR05 == 4 //Entradas ou todos

//Filtro das Entradas
cFiltro 	:= "%AND D1_FILIAL BETWEEN '" + ALLTRIM(MV_PAR01) + "' AND '" + ALLTRIM(MV_PAR02) + "' AND D1_DTDIGIT BETWEEN '" + DTOS(MV_PAR03) + "' AND '" + DTOS(MV_PAR04) + "' AND D1_FORNECE BETWEEN '" + (MV_PAR06) + "' AND '" + (MV_PAR07) + "' AND D1_COD BETWEEN '" + (MV_PAR10) + "' AND '" + (MV_PAR11) + "' AND D1_COD BETWEEN '" + (MV_PAR12) + "' AND '" + (MV_PAR13) + "' %"
If MV_PAR14 == 1
cCampoImp 	:= "%, D1_BASEICM AS BICMS, D1_PICM AS AICMS, D1_VALICM AS VICMS, D1_ICMSCOM AS CICMS, D1_BASEIPI AS BIPI, D1_IPI AS AIPI, D1_VALIPI AS VIPI " //ICMS e IPI
cCampoImp 	+= ", D1_BASEPIS AS BPIS, D1_ALQPIS AS APIS, D1_VALPIS AS VPIS, D1_BASECOF AS BCOFINS, D1_ALQCOF AS ACOFINS, D1_VALCOF AS VCOFINS, D1_BASECSL AS BCSLL, D1_ALQCSL AS ACSLL, D1_VALCSL AS VCSLL " //PIS, COFINS e CSLL
cCampoImp 	+= ", D1_BASEIRR AS BIRRF, D1_ALIQIRR AS AIRRF, D1_VALIRR AS VIRRF, D1_BASEISS AS BISS, D1_ALIQISS AS AISS, D1_VALISS AS VISS, D1_BASEINS AS BINSS, D1_ALIQINS AS AINSS, D1_VALINS AS VINSS %" //IRRF, ISS e INSS
EndIf

BeginSql Alias cAlias

//Query com o tratamento das informações da tabela SD1 de entrada
SELECT 'ENTRADA' AS TIPO, '' AS EMPRESA,D1_FILIAL AS FILIAL,D1_TIPO AS TIPONF
, D1_DOC AS NFISCAL, D1_SERIE AS SERIE, D1_ITEM AS ITEM, D1_COD AS PRODUTO
, AH_UMRES AS MEDIDA, B1_DESC AS DESCRICAOP, B1_POSIPI AS NCM, D1_QUANT AS QUANTIDADE,D1_VUNIT AS VLUNIT, D1_TOTAL AS VLTOTAL, D1_LOCAL AS ARMAZEM,NNR_DESCRI AS DESCRICAOA
, D1_LOTECTL AS LOTE, D1_DFABRIC AS FABRICACAO, D1_DTVALID AS VALIDADE, D1_DTDIGIT AS ENTRADA, D1_EMISSAO AS EMISSAO
, D1_TES AS TES, F4_TEXTO AS FINALIDADE, F4_ESTOQUE AS ESTOQUE, F4_PODER3 AS PORDER3, F4_DUPLIC AS FINANCEIRO, D1_CF AS CFOP
, D1_FORNECE AS CLIFOR, D1_LOJA AS LOJA
, CASE WHEN D1_TIPO NOT IN ('D','B') THEN (SELECT A2_NOME FROM %table:SA2% WHERE D1_FORNECE = A2_COD AND D1_LOJA = A2_LOJA AND %table:SA2%.D_E_L_E_T_ = '' )
WHEN D1_TIPO IN ('D','B') THEN (SELECT A1_NOME FROM %table:SA1% WHERE D1_FORNECE = A1_COD AND D1_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS NOME
, CASE WHEN D1_TIPO NOT IN ('D','B') THEN (SELECT A2_NREDUZ FROM %table:SA2% WHERE D1_FORNECE = A2_COD AND D1_LOJA = A2_LOJA AND %table:SA2%.D_E_L_E_T_ = '' )
WHEN D1_TIPO IN ('D','B') THEN (SELECT A1_NREDUZ FROM %table:SA1% WHERE D1_FORNECE = A1_COD AND D1_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS NREDUZ
, CASE WHEN D1_TIPO NOT IN ('D','B') THEN (SELECT A2_EST FROM %table:SA2% WHERE D1_FORNECE = A2_COD AND D1_LOJA = A2_LOJA AND %table:SA2%.D_E_L_E_T_ = '' )
WHEN D1_TIPO IN ('D','B') THEN (SELECT A1_EST FROM %table:SA1% WHERE D1_FORNECE = A1_COD AND D1_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS ESTADO
, CASE WHEN D1_TIPO NOT IN ('D','B') THEN (SELECT A2_TIPO FROM %table:SA2% WHERE D1_FORNECE = A2_COD AND D1_LOJA = A2_LOJA AND %table:SA2%.D_E_L_E_T_ = '' )
WHEN D1_TIPO IN ('D','B') THEN (SELECT A1_PESSOA FROM %table:SA1% WHERE D1_FORNECE = A1_COD AND D1_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS PESSOA
, CASE WHEN D1_TIPO NOT IN ('D','B') THEN (SELECT A2_CGC FROM %table:SA2% WHERE D1_FORNECE = A2_COD AND D1_LOJA = A2_LOJA AND %table:SA2%.D_E_L_E_T_ = '' )
WHEN D1_TIPO IN ('D','B') THEN (SELECT A1_CGC FROM %table:SA1% WHERE D1_FORNECE = A1_COD AND D1_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS CNPJ
, D1_PEDIDO AS PEDIDO, D1_ITEMPC AS ITEMPD
, CASE D1_TIPO WHEN 'D' THEN 'DEVOLUCAO' WHEN 'N' THEN 'NORMAL' WHEN 'C' THEN 'COMPLEMENTO VALOR' WHEN 'B' THEN 'BENEFICIAMENTO' WHEN 'P' THEN 'COMPLEMENTO IPI' WHEN 'I' THEN 'COMPLEMENTO ICMS' END AS TIPONF1
//Rafael França - 07/07/20 - Inclusão das informações: Codigo de municipio, Municipio, Canal e Descrição (Cliente)
, CASE WHEN D1_TIPO NOT IN ('D','B') THEN (SELECT A2_COD_MUN FROM %table:SA2% WHERE D1_FORNECE = A2_COD AND D1_LOJA = A2_LOJA AND %table:SA2%.D_E_L_E_T_ = '' )
WHEN D1_TIPO IN ('D','B') THEN (SELECT A1_COD_MUN FROM %table:SA1% WHERE D1_FORNECE = A1_COD AND D1_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS COD_MUN
, CASE WHEN D1_TIPO NOT IN ('D','B') THEN (SELECT A2_MUN FROM %table:SA2% WHERE D1_FORNECE = A2_COD AND D1_LOJA = A2_LOJA AND %table:SA2%.D_E_L_E_T_ = '' )
WHEN D1_TIPO IN ('D','B') THEN (SELECT A1_MUN FROM %table:SA1% WHERE D1_FORNECE = A1_COD AND D1_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS MUNICIPIO
, CASE WHEN D1_TIPO NOT IN ('D','B') THEN (SELECT A2_PAIS FROM %table:SA2% WHERE D1_FORNECE = A2_COD AND D1_LOJA = A2_LOJA AND %table:SA2%.D_E_L_E_T_ = '' )
WHEN D1_TIPO IN ('D','B') THEN (SELECT A1_PAIS FROM %table:SA1% WHERE D1_FORNECE = A1_COD AND D1_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS COD_PAIS
, CASE WHEN D1_TIPO NOT IN ('D','B') THEN '' //GF
WHEN D1_TIPO IN ('D','B') THEN (SELECT A1_ZZCANAL FROM %table:SA1% WHERE D1_FORNECE = A1_COD AND D1_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS COD_CANAL //GF
, CASE WHEN D1_TIPO NOT IN ('D','B') THEN '' //GF
WHEN D1_TIPO IN ('D','B') THEN (SELECT A1_ZZDSCAN FROM %table:SA1% WHERE D1_FORNECE = A1_COD AND D1_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS CANAL //GF
, CASE WHEN D1_TIPO NOT IN ('D','B') THEN '' //GF
WHEN D1_TIPO IN ('D','B') THEN (SELECT A1_REGIAO FROM %table:SA1% WHERE D1_FORNECE = A1_COD AND D1_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS COD_REGIAO //GF
, CASE WHEN D1_TIPO NOT IN ('D','B') THEN '' //GF
WHEN D1_TIPO IN ('D','B') THEN (SELECT A1_DSCREG FROM %table:SA1% WHERE D1_FORNECE = A1_COD AND D1_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS REGIAO //GF
, F1_CHVNFE AS CHAVENFE, F1_USERLGI AS USUARIO
//Dados complementares
, D1_CONTA AS CCONTABIL, D1_CC AS CCUSTO, D1_VALDESC AS DESCONTO, D1_VALFRE AS FRETE, F1_TPFRETE AS TPFRETE, D1_CUSTO AS CUSTO
, D1_NFORI AS NFORIGEM, D1_SERIORI AS SERIEORI, D1_ITEMORI AS ITEMORI, D1_SEGUM AS SEGUNDAUM, D1_QTSEGUM AS QTDSEG
//Impostos
%exp:cCampoImp%
FROM %table:SD1%
INNER JOIN %table:SF1% ON D1_FILIAL = F1_FILIAL AND D1_DOC = F1_DOC AND D1_FORNECE = F1_FORNECE AND D1_LOJA = F1_LOJA AND D1_SERIE = F1_SERIE AND D1_EMISSAO = F1_EMISSAO AND %table:SF1%.D_E_L_E_T_ = ''
INNER JOIN %table:SB1% ON D1_COD = B1_COD AND %table:SB1%.D_E_L_E_T_ = ''
INNER JOIN %table:SF4% ON D1_TES = F4_CODIGO AND %table:SF4%.D_E_L_E_T_ = ''
INNER JOIN %table:NNR% ON D1_FILIAL = NNR_FILIAL AND D1_LOCAL = NNR_CODIGO AND %table:NNR%.D_E_L_E_T_ = '' //GF NNR Exclusiva
//INNER JOIN %table:NNR% ON D1_LOCAL = NNR_CODIGO AND %table:NNR%.D_E_L_E_T_ = '' //NNR Compartilhada
INNER JOIN %table:SAH% ON D1_UM = AH_UNIMED AND %table:SAH%.D_E_L_E_T_ = ''
WHERE %table:SD1%.D_E_L_E_T_ = ''
%exp:cFiltro%
ORDER BY EMPRESA,FILIAL,ENTRADA,NFISCAL,SERIE,ITEM

EndSql //FINALIZO A MINHA QUERY

cNomeTabela := "SD1 Entradas - " + DTOC(MV_PAR03) + " a " + DTOC(MV_PAR04)

    //Criando o objeto que irá gerar o conteúdo do Excel
    oFWMsExcel := FWMSExcel():New()

    //Aba 01 - Mapas
    oFWMsExcel:AddworkSheet("Entradas") //Não utilizar número junto com sinal de menos. Ex.: 1-

		//Criando a Tabela
        oFWMsExcel:AddTable("Entradas",cNomeTabela)

//Criando Colunas
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"TIPO"   		,1,1) //01 //1,1 = Modo Texto  // 2,2 = Valor sem R$  //  3,3 = Valor com R$
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"EMPRESA"		,1,1) //02
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"FILIAL"	   	,1,1) //03
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"TIPONF"	    ,1,1) //04
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"NFISCAL"  		,1,1) //05
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"SERIE"   	  	,1,1) //06
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"ITEM"   	  	,1,1) //07
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"PRODUTO"   	,1,1) //08
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"DESCR_PROD"	,1,1) //09
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"QUANTIDADE"	,2,2) //10
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"MEDIDA"	    ,1,1) //11
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"VLUNIT"	    ,2,2) //12
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"VLTOTAL"  		,2,2) //13
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"ARMAZEM"   	,1,1) //14
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"DESCR_ARMAZ"  	,1,1) //15 // FINAL 1 LINHA DE COLUNAS
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"LOTE" 	 		,1,1) //16
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"FABRICACAO"   	,1,1) //17
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"VALIDADE"   	,1,1) //18
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"DTENTRADA"   	,1,1) //19
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"EMISSAO"   	,1,1) //20
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"TES"			,1,1) //21
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"FINALIDADE"	,1,1) //22
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"ESTOQUE"   	,1,1) //23
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"PODER3"    	,1,1) //24
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"FINANCEIRO"	,1,1) //25
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"CFOP"	   		,1,1) //26 // FINAL 2 LINHA DE COLUNAS
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"CLIFOR"	    ,1,1) //27
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"LOJA"		    ,1,1) //28
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"NOME"  		,1,1) //29
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"NREDUZ"  		,1,1) //30
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"ESTADO"   	  	,1,1) //31
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"PESSOA"   	  	,1,1) //32
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"CNPJ_CPF" 	  	,1,1) //33
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"PEDIDO"   	  	,1,1) //34
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"ITEMPD"   	  	,1,1) //35
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"COD_MUN"   	,1,1) //36
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"MUNICIPIO"  	,1,1) //37
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"COD_PAIS"      ,1,1) //38
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"PAIS"      	,1,1) //39
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"COD_CANAL"     ,1,1) //40 - GF
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"CANAL"   	  	,1,1) //41 - GF
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"COD_REGIAO"    ,1,1) //42
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"REGIAO"   	  	,1,1) //43
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"CHAVENFE" 	  	,1,1) //44 // FINAL 3 LINHA DE COLUNAS
//Dados Complementares
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"CCONTABIL"  	,1,1) //45
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"DESCR_CONTA"   ,1,1) //46
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"CCUSTO"      	,1,1) //47
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"DESCR_CUSTO"   ,1,1) //48
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"DESCONTO" 	  	,3,3) //49
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"FRETE" 	  	,3,3) //50
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"TPFRETE" 	  	,1,1) //51
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"CUSTO" 	  	,3,3) //52 // FINAL 4 LINHA DE COLUNAS
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"NCM"		  	,1,1) //53
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"NF_ORIGEM"	  	,1,1) //54
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"SERIE_ORIGEM" 	,1,1) //55
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"ITEM_ORIGEM"  	,1,1) //56
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"SEGUNDA_UM"  	,1,1) //57
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"QTD_SEGUNDA"  	,2,2) //58
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"USUARIO"  		,1,1) //59 // FINAL 5 LINHA DE COLUNAS
If MV_PAR14 == 1
//Impostos //ICMS e IPI
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"BASEICMS"  	,3,3) //01
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"ALIQICMS"      ,2,2) //02
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"VALORICMS"    	,3,3) //03
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"ICMS_COMPL"    ,3,3) //04
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"BASEIPI"  		,3,3) //05
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"ALIQIPI"      	,2,2) //06
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"VALORIPI"    	,3,3) //07
//PIS, COFINS e CSLL
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"BASEPIS"  		,3,3) //08
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"ALIQPIS"     	,2,2) //09
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"VALORPIS"    	,3,3) //10
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"BASECOFINS"  	,3,3) //11
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"ALIQCOFINS"    ,2,2) //12
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"VALORCOFINS"   ,3,3) //13
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"BASECSLL"  	,3,3) //14
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"ALIQCSLL"      ,2,2) //15
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"VALORCSLL"    	,3,3) //16
//IRRF, ISS e INSS
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"BASEIRRF" 		,3,3) //17
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"ALIQIRRF"     	,2,2) //18
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"VALORIRRF"    	,3,3) //19
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"BASEISS"  		,3,3) //20
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"ALIQISS"    	,2,2) //21
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"VALORISS"   	,3,3) //22
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"BASEINSS"  	,3,3) //23
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"ALIQINSS"      ,2,2) //24
oFWMsExcel:AddColumn("Entradas",cNomeTabela,"VALORINSS"    	,3,3) //25
Endif

While !(cAlias)->(Eof())

	//cUsuario	:= FWLeUserlg("USUARIO", 1)
	//cData  		:= FWLeUserlg("USUARIO", 2)

	cUsuario:= Substr(Embaralha((cAlias)->USUARIO,1),3,6)

	If PswSeek(cUsuario, .T. )
			aUsuario := PswRet() // Retorna vetor com informações do usuário
			cUsuario := Alltrim(aUsuario[1][2])
	EndIf

	//Criando as Linhas
	If MV_PAR14 == 1
    oFWMsExcel:AddRow("Entradas",cNomeTabela,{TIPO,EMPRESA,FILIAL,TIPONF1,NFISCAL,SERIE,ITEM,PRODUTO,DESCRICAOP,QUANTIDADE,MEDIDA,VLUNIT,VLTOTAL,ARMAZEM,DESCRICAOA; //Campo 15
	,LOTE,DTOC(STOD(FABRICACAO)),DTOC(STOD(VALIDADE)),DTOC(STOD(ENTRADA)),DTOC(STOD(EMISSAO)),TES,FINALIDADE,ESTOQUE,PORDER3,FINANCEIRO,CFOP,; //Campo 26
	CLIFOR,LOJA,NOME,NREDUZ,ESTADO,PESSOA,CNPJ,PEDIDO,ITEMPD,COD_MUN,MUNICIPIO,COD_PAIS,Posicione('SYA',1,xFilial('SYA')+COD_PAIS,'YA_DESCR'),COD_CANAL,CANAL,COD_REGIAO,REGIAO,CHAVENFE; //Campo 44
	,CCONTABIL,Posicione("CT1",1,xFilial("CT1")+CCONTABIL,"CT1_DESC01"),CCUSTO,Posicione("CTT",1,xFilial("CTT")+CCUSTO,"CTT_DESC01"),DESCONTO,FRETE,TPFRETE,CUSTO; //Campo 52
	,NCM,NFORIGEM,SERIEORI,ITEMORI,SEGUNDAUM,QTDSEG,cUsuario; //Campo 58
	,BICMS,AICMS,VICMS,CICMS,BIPI,AIPI,VIPI,BPIS,APIS,VPIS,BCOFINS,ACOFINS,VCOFINS,BCSLL,ACSLL,VCSLL,BIRRF,AIRRF,VIRRF,BISS,AISS,VISS,BINSS,AINSS,VINSS}) //Campos Impostos
	else
   oFWMsExcel:AddRow("Entradas",cNomeTabela,{TIPO,EMPRESA,FILIAL,TIPONF1,NFISCAL,SERIE,ITEM,PRODUTO,DESCRICAOP,QUANTIDADE,MEDIDA,VLUNIT,VLTOTAL,ARMAZEM,DESCRICAOA; //Campo 15
	,LOTE,DTOC(STOD(FABRICACAO)),DTOC(STOD(VALIDADE)),DTOC(STOD(ENTRADA)),DTOC(STOD(EMISSAO)),TES,FINALIDADE,ESTOQUE,PORDER3,FINANCEIRO,CFOP,; //Campo 26
	CLIFOR,LOJA,NOME,NREDUZ,ESTADO,PESSOA,CNPJ,PEDIDO,ITEMPD,COD_MUN,MUNICIPIO,COD_PAIS,Posicione('SYA',1,xFilial('SYA')+COD_PAIS,'YA_DESCR'),COD_CANAL,CANAL,COD_REGIAO,REGIAO,CHAVENFE; //Campo 44
	,CCONTABIL,Posicione("CT1",1,xFilial("CT1")+CCONTABIL,"CT1_DESC01"),CCUSTO,Posicione("CTT",1,xFilial("CTT")+CCUSTO,"CTT_DESC01"),DESCONTO,FRETE,TPFRETE,CUSTO; //Campo 52
	,NCM,NFORIGEM,SERIEORI,ITEMORI,SEGUNDAUM,QTDSEG,cUsuario}) //Campo 58
	Endif

	(cAlias)->(dbSkip()) //PASSAR PARA O PRÓXIMO REGISTRO DA MINHA QUERY

Enddo

(cAlias)->(dbClosearea()) //FECHO A TABELA APOS O USO

ENDIF

IF MV_PAR05 == 2 .OR. MV_PAR05 == 4 //Saídas ou todos

//Filtro das saídas
cFiltro 	:= "%AND D2_FILIAL BETWEEN '" + ALLTRIM(MV_PAR01) + "' AND '" + ALLTRIM(MV_PAR02) + "' AND D2_EMISSAO BETWEEN '" + DTOS(MV_PAR03) + "' AND '" + DTOS(MV_PAR04) + "' AND D2_CLIENTE BETWEEN '" + (MV_PAR08) + "' AND '" + (MV_PAR09) + "' AND D2_COD BETWEEN '" + (MV_PAR10) + "' AND '" + (MV_PAR11) + "' AND D2_TES BETWEEN '" + (MV_PAR12) + "' AND '" + (MV_PAR13) + "' %"
//If MV_PAR14 == 1
cCampoImp 	:= "%, D2_BASEICM AS BICMS, D2_PICM AS AICMS, D2_VALICM AS VICMS, D2_ICMSCOM AS CICMS,D2_BRICMS AS BICMSRET, D2_ICMSRET AS ICMSRET, D2_BASEIPI AS BIPI, D2_IPI AS AIPI, D2_VALIPI AS VIPI " //ICMS e IPI
cCampoImp 	+= ", D2_BASEPIS AS BPIS, D2_ALQPIS AS APIS, D2_VALPIS AS VPIS, D2_BASECOF AS BCOFINS, D2_ALQCOF AS ACOFINS, D2_VALCOF AS VCOFINS, D2_BASECSL AS BCSLL, D2_ALQCSL AS ACSLL, D2_VALCSL AS VCSLL " //PIS, COFINS e CSLL
cCampoImp 	+= ", D2_BASEIRR AS BIRRF, D2_ALQIRRF AS AIRRF, D2_VALIRRF AS VIRRF, D2_BASEISS AS BISS, D2_ALIQISS AS AISS, D2_VALISS AS VISS, D2_BASEINS AS BINSS, D2_ALIQINS AS AINSS, D2_VALINS AS VINSS %" //IRRF, ISS e INSS%"
//EndIf

BeginSql Alias cAlias

//Query com o tratamento das informações da tabela SD2 de saída - Aqui é onde separa as informações no banco de dados para gerar a planilha -> D2_FILIAL (Nome do Campo) AS FILIAL (Apelido)
SELECT 'SAIDA' AS TIPO, '10 - GLOBALFRUIT' AS EMPRESA,D2_FILIAL AS FILIAL,D2_TIPO AS TIPONF, D2_DOC AS NFISCAL, D2_SERIE AS SERIE, D2_ITEM AS ITEM, D2_COD AS PRODUTO
, AH_UMRES AS MEDIDA, B1_DESC AS DESCRICAOP, B1_POSIPI AS NCM, D2_QUANT AS QUANTIDADE,D2_PRCVEN AS VLUNIT, D2_TOTAL AS VLTOTAL, D2_LOCAL AS ARMAZEM,NNR_DESCRI AS DESCRICAOA
, D2_LOTECTL AS LOTE, D2_DFABRIC AS FABRICACAO, D2_DTVALID AS VALIDADE, D2_EMISSAO AS ENTRADA, D2_EMISSAO AS EMISSAO
, D2_TES AS TES, F4_TEXTO AS FINALIDADE, F4_ESTOQUE AS ESTOQUE, F4_PODER3 AS PODER3, F4_DUPLIC AS FINANCEIRO, D2_CF AS CFOP
, D2_CLIENTE AS CLIFOR, D2_LOJA AS LOJA
,CASE WHEN D2_TIPO IN ('D','B') THEN (SELECT A2_NOME FROM %table:SA2% WHERE D2_CLIENTE = A2_COD AND D2_LOJA = A2_LOJA AND %table:SA2%.D_E_L_E_T_ = '' )
WHEN D2_TIPO NOT IN ('D','B') THEN (SELECT A1_NOME FROM %table:SA1% WHERE D2_CLIENTE = A1_COD AND D2_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS NOME
,CASE WHEN D2_TIPO IN ('D','B') THEN (SELECT A2_NREDUZ FROM %table:SA2% WHERE D2_CLIENTE = A2_COD AND D2_LOJA = A2_LOJA AND %table:SA2%.D_E_L_E_T_ = '' )
WHEN D2_TIPO NOT IN ('D','B') THEN (SELECT A1_NREDUZ FROM %table:SA1% WHERE D2_CLIENTE = A1_COD AND D2_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS NREDUZ
, CASE WHEN D2_TIPO IN ('D','B') THEN (SELECT A2_EST FROM %table:SA2% WHERE D2_CLIENTE = A2_COD AND D2_LOJA = A2_LOJA AND %table:SA2%.D_E_L_E_T_ = '' )
WHEN D2_TIPO NOT IN ('D','B') THEN (SELECT A1_EST FROM %table:SA1% WHERE D2_CLIENTE = A1_COD AND D2_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS ESTADO
, CASE WHEN D2_TIPO IN ('D','B') THEN (SELECT A2_TIPO FROM %table:SA2% WHERE D2_CLIENTE = A2_COD AND D2_LOJA = A2_LOJA AND %table:SA2%.D_E_L_E_T_ = '' )
WHEN D2_TIPO NOT IN ('D','B') THEN (SELECT A1_PESSOA FROM %table:SA1% WHERE D2_CLIENTE = A1_COD AND D2_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS PESSOA
, CASE WHEN D2_TIPO IN ('D','B') THEN (SELECT A2_CGC FROM %table:SA2% WHERE D2_CLIENTE = A2_COD AND D2_LOJA = A2_LOJA AND %table:SA2%.D_E_L_E_T_ = '' )
WHEN D2_TIPO NOT IN ('D','B') THEN (SELECT A1_CGC FROM %table:SA1% WHERE D2_CLIENTE = A1_COD AND D2_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS CNPJ
, CASE WHEN D2_TIPO IN ('D','B') THEN '' //GF
WHEN D2_TIPO NOT IN ('D','B') THEN (SELECT A1_TIPO FROM %table:SA1% WHERE D2_CLIENTE = A1_COD AND D2_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS TIPOCLI //GF
, D2_PEDIDO AS PEDIDO, D2_ITEMPV AS ITEMPD,F2_MENNOTA AS MENNOTA
, CASE D2_TIPO WHEN 'D' THEN 'DEVOLUCAO' WHEN 'N' THEN 'NORMAL' WHEN 'C' THEN 'COMPLEMENTO VALOR' WHEN 'B' THEN 'BENEFICIAMENTO' WHEN 'P' THEN 'COMPLEMENTO IPI' WHEN 'I' THEN 'COMPLEMENTO ICMS' END AS TIPONF1
//Rafael França - 07/07/20 - Inclusão das informações: Codigo de municipio, Municipio, Canal e Descrição (Cliente)
, CASE WHEN D2_TIPO IN ('D','B') THEN (SELECT A2_COD_MUN FROM %table:SA2% WHERE D2_CLIENTE = A2_COD AND D2_LOJA = A2_LOJA AND %table:SA2%.D_E_L_E_T_ = '' )
WHEN D2_TIPO NOT IN ('D','B') THEN (SELECT A1_COD_MUN FROM %table:SA1% WHERE D2_CLIENTE = A1_COD AND D2_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS COD_MUN
, CASE WHEN D2_TIPO IN ('D','B') THEN (SELECT A2_MUN FROM %table:SA2% WHERE D2_CLIENTE = A2_COD AND D2_LOJA = A2_LOJA AND %table:SA2%.D_E_L_E_T_ = '' )
WHEN D2_TIPO NOT IN ('D','B') THEN (SELECT A1_MUN FROM %table:SA1% WHERE D2_CLIENTE = A1_COD AND D2_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS MUNICIPIO
, CASE WHEN D2_TIPO IN ('D','B') THEN (SELECT A2_PAIS FROM %table:SA2% WHERE D2_CLIENTE = A2_COD AND D2_LOJA = A2_LOJA AND %table:SA2%.D_E_L_E_T_ = '' )
WHEN D2_TIPO NOT IN ('D','B') THEN (SELECT A1_PAIS FROM %table:SA1% WHERE D2_CLIENTE = A1_COD AND D2_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS COD_PAIS
, CASE WHEN D2_TIPO IN ('D','B') THEN '' //GF
WHEN D2_TIPO NOT IN ('D','B') THEN (SELECT A1_ZZCANAL FROM %table:SA1% WHERE D2_CLIENTE = A1_COD AND D2_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS COD_CANAL //GF
, CASE WHEN D2_TIPO IN ('D','B') THEN '' //GF
WHEN D2_TIPO NOT IN ('D','B') THEN (SELECT A1_ZZDSCAN FROM %table:SA1% WHERE D2_CLIENTE = A1_COD AND D2_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS CANAL //GF
, CASE WHEN D2_TIPO IN ('D','B') THEN '' //GF
WHEN D2_TIPO NOT IN ('D','B') THEN (SELECT A1_REGIAO FROM %table:SA1% WHERE D2_CLIENTE = A1_COD AND D2_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS COD_REGIAO //GF
, CASE WHEN D2_TIPO IN ('D','B') THEN '' //GF
WHEN D2_TIPO NOT IN ('D','B') THEN (SELECT A1_DSCREG FROM %table:SA1% WHERE D2_CLIENTE = A1_COD AND D2_LOJA = A1_LOJA AND %table:SA1%.D_E_L_E_T_ = '' ) END AS REGIAO //GF
, F2_VEND1 AS VENDEDOR, F2_CHVNFE AS CHAVENFE, F2_USERLGI AS USUARIO
//Dados complementares
, D2_CONTA AS CCONTABIL, D2_CCUSTO AS CCUSTO, D2_DESCON AS DESCONTO, D2_VALFRE AS FRETE, F2_TPFRETE AS TPFRETE, D2_CUSTO1 AS CUSTO
, D2_NFORI AS NFORIGEM, D2_SERIORI AS SERIEORI, D2_ITEMORI AS ITEMORI
//Impostos
%exp:cCampoImp%
FROM %table:SD2%
INNER JOIN %table:SF2% ON D2_FILIAL = F2_FILIAL AND D2_DOC = F2_DOC AND D2_SERIE = F2_SERIE AND D2_CLIENTE = F2_CLIENTE AND D2_LOJA = F2_LOJA AND %table:SF2%.D_E_L_E_T_ = ''
INNER JOIN %table:SB1% ON D2_COD = B1_COD AND %table:SB1%.D_E_L_E_T_ = ''
INNER JOIN %table:SF4% ON D2_TES = F4_CODIGO AND %table:SF4%.D_E_L_E_T_ = ''
INNER JOIN %table:NNR% ON D2_FILIAL = NNR_FILIAL AND D2_LOCAL = NNR_CODIGO AND %table:NNR%.D_E_L_E_T_ = '' //GF NNR Exclusiva
//INNER JOIN %table:NNR% ON D2_LOCAL = NNR_CODIGO AND %table:NNR%.D_E_L_E_T_ = '' //NNR Compartilhada
INNER JOIN %table:SAH% ON D2_UM = AH_UNIMED AND %table:SAH%.D_E_L_E_T_ = ''
WHERE %table:SD2%.D_E_L_E_T_ = ''
%exp:cFiltro%
ORDER BY EMPRESA,FILIAL,ENTRADA,NFISCAL,SERIE,ITEM

EndSql //FINALIZO A MINHA QUERY

cNomeTabela := "SD2 Saidas - " + DTOC(MV_PAR03) + " a " + DTOC(MV_PAR04)

	IF MV_PAR05 == 2 //Crio o objeto se ele não foi criado
    //Criando o objeto que irá gerar o conteúdo do Excel
    oFWMsExcel := FWMSExcel():New()
	ENDIF

    //Aba 01 - Mapas
    oFWMsExcel:AddworkSheet("Saidas") //Não utilizar número junto com sinal de menos. Ex.: 1-

		//Criando a Tabela
        oFWMsExcel:AddTable("Saidas",cNomeTabela)

//Criando Colunas - Aqui coloco a ordem dos campos que deveram ser gerados na planilha.
//Dados do cabeçalho nota e dados do cliente/fornecedor
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"TIPO"   		,1,1) //01 //1,1 = Modo Texto  // 2,2 = Valor sem R$  //  3,3 = Valor com R$
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"EMPRESA"		,1,1) //02
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"FILIAL"	   	,1,1) //03
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"TIPONF"	    ,1,1) //04
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"NFISCAL"  	,1,1) //05
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"SERIE"   	,1,1) //06
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"EMISSAO"   	,1,1) //07
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"DTSAIDA"   	,1,1) //08
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"CHAVENFE"   	,1,1) //09
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"CLIFOR"	    ,1,1) //10
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"LOJA"	    ,1,1) //11
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"NOME"  		,1,1) //12
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"NREDUZ" 		,1,1) //13
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"CNPJ_CPF"   	,1,1) //14
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"PESSOA"   	,1,1) //15
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"TIPOCLI"   	,1,1) //16
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"COD_MUN"   	,1,1) //17
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"MUNICIPIO"  	,1,1) //18
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"ESTADO"   	,1,1) //19
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"COD_PAIS"   	,1,1) //20
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"PAIS"   	    ,1,1) //21
//Dados itens documento de saida
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"ITEM"   	  	,1,1) //22
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"PRODUTO"   	,1,1) //23
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"DESCR_PROD"	,1,1) //24
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"QUANTIDADE"	,2,2) //25
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"MEDIDA"	    ,1,1) //26
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"NCM"		  	,1,1) //27
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"VLUNIT"	    ,2,2) //28
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"VLTOTAL"  	,2,2) //29
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"ARMAZEM"   	,1,1) //30
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"DESCR_ARMAZ"	,1,1) //31
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"LOTE" 	 	,1,1) //32
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"FABRICACAO" 	,1,1) //33
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"VALIDADE"   	,1,1) //34
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"DESCONTO" 	,3,3) //35
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"FRETE" 	  	,3,3) //36
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"TPFRETE" 	,1,1) //37
//Dados impostos - ICMS e IPI
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"BASEICMS"  	,3,3) //38
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"ALIQICMS"    ,2,2) //39
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"VALORICMS"   ,3,3) //40
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"ICMS_COMPL"  ,3,3) //41
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"BICMS_RET"  	,3,3) //42
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"ICMS_RET"    ,3,3) //43
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"BASEIPI"  	,3,3) //44
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"ALIQIPI"     ,2,2) //45
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"VALORIPI"    ,3,3) //46
//PIS, COFINS e CSLL
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"BASEPIS"  	,3,3) //47
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"ALIQPIS"     ,2,2) //48
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"VALORPIS"    ,3,3) //49
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"BASECOFINS"  ,3,3) //50
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"ALIQCOFINS"  ,2,2) //51
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"VALORCOFINS" ,3,3) //52
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"BASECSLL"  	,3,3) //53
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"ALIQCSLL"    ,2,2) //54
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"VALORCSLL"   ,3,3) //55
//IRRF, ISS e INSS
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"BASEIRRF" 	,3,3) //56
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"ALIQIRRF"    ,2,2) //57
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"VALORIRRF"   ,3,3) //58
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"BASEISS"  	,3,3) //59
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"ALIQISS"    	,2,2) //60
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"VALORISS"   	,3,3) //61
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"BASEINSS"  	,3,3) //62
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"ALIQINSS"    ,2,2) //63
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"VALORINSS"   ,3,3) //64
//Dados fiscais
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"VLFINAL"  	,2,2) //65
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"TES"			,1,1) //66
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"FINALIDADE"	,1,1) //67
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"ESTOQUE"   	,1,1) //68
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"PODER3"  	,1,1) //69
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"FINANCEIRO"	,1,1) //70
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"CFOP"	   	,1,1) //71
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"CUSTO" 	  	,3,3) //72
//Dados contabeis
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"CCONTABIL"  	,1,1) //73
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"DESCR_CONTA" ,1,1) //74
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"CCUSTO"      ,1,1) //75
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"DESCR_CUSTO" ,1,1) //76
//Dados de referencia da nota do cliente
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"NF_ORIGEM"	,1,1) //77
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"SERIE_ORIGEM",1,1) //78
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"ITEM_ORIGEM"	,1,1) //79
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"PED_CLIENTE"	,1,1) //80
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"ITEM_CLIENTE",1,1) //81
//Dados adicionais
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"PEDIDO"   	,1,1) //82
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"ITEMPD"   	,1,1) //83
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"COD_CANAL"   ,1,1) //84 - GF
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"CANAL"   	,1,1) //85 - GF
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"COD_REGIAO"  ,1,1) //86
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"REGIAO"   	,1,1) //87
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"VENDEDOR"	,1,1) //88
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"NOME_VEND"  	,1,1) //89
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"MEN_NOTA"  	,1,1) //90
oFWMsExcel:AddColumn("Saidas",cNomeTabela,"USUARIO"  	,1,1) //91

While !(cAlias)->(Eof())

	//cUsuario	:= FWLeUserlg("USUARIO", 1)
	//cData  		:= FWLeUserlg("USUARIO", 2)

	cUsuario:= Substr(Embaralha((cAlias)->USUARIO,1),3,6)

	If PswSeek(cUsuario, .T. )
			aUsuario := PswRet() // Retorna vetor com informações do usuário
			cUsuario := Alltrim(aUsuario[1][2])
	EndIf

//If MV_PAR14 == 1 - Novo layout sempre imprime os impostos nas notas de saidas
	//Criando as Linhas - Aqui são geradas as linhas na mesma ordem das colunas
    oFWMsExcel:AddRow("Saidas",cNomeTabela,{TIPO,EMPRESA,FILIAL,TIPONF1,NFISCAL,SERIE,DTOC(STOD(EMISSAO)),DTOC(STOD(ENTRADA)),CHAVENFE,CLIFOR,LOJA,NOME,NREDUZ,CNPJ,PESSOA,TIPOCLI,COD_MUN,MUNICIPIO,ESTADO,COD_PAIS,Posicione('SYA',1,xFilial('SYA')+COD_PAIS,'YA_DESCR'); //Dados do cabeçalho e informações cliente/fornecedor
	,ITEM,PRODUTO,DESCRICAOP,QUANTIDADE,MEDIDA,NCM,((VLTOTAL+DESCONTO)/QUANTIDADE),(VLTOTAL+DESCONTO),ARMAZEM,DESCRICAOA,LOTE,DTOC(STOD(FABRICACAO)),DTOC(STOD(VALIDADE)),DESCONTO,FRETE,TPFRETE; //Dados dos itens da nota
	,BICMS,AICMS,VICMS,CICMS,BICMSRET,ICMSRET,BIPI,AIPI,VIPI,BPIS,APIS,VPIS,BCOFINS,ACOFINS,VCOFINS,BCSLL,ACSLL,VCSLL,BIRRF,AIRRF,VIRRF,BISS,AISS,VISS,BINSS,AINSS,VINSS; //Impostos
	,(VLTOTAL+FRETE+ICMSRET),TES,FINALIDADE,ESTOQUE,PODER3,FINANCEIRO,CFOP,CUSTO; //Dados fiscais
	,CCONTABIL,Posicione("CT1",1,xFilial("CT1")+CCONTABIL,"CT1_DESC01"),CCUSTO,Posicione("CTT",1,xFilial("CTT")+CCUSTO,"CTT_DESC01"); //Dados Contabeis
	,NFORIGEM,SERIEORI,ITEMORI,Posicione("SC6",1,xFilial("SC6")+PEDIDO+ITEMPD+PRODUTO,"C6_NUMPCOM"),Posicione("SC6",1,xFilial("SC6")+PEDIDO+ITEMPD+PRODUTO,"C6_ITEMPC"); //Dados de referencia
	,PEDIDO,ITEMPD,COD_CANAL,CANAL,COD_REGIAO,REGIAO,VENDEDOR,Posicione("SA3",1,xFilial("SA3")+VENDEDOR,"A3_NOME"),MENNOTA,cUsuario}) //Dados adicionais
	/* else
    oFWMsExcel:AddRow("Saidas",cNomeTabela,{TIPO,EMPRESA,FILIAL,TIPONF1,NFISCAL,SERIE,ITEM,PRODUTO,DESCRICAOP,QUANTIDADE,MEDIDA,VLUNIT,VLTOTAL,(VLTOTAL+FRETE+ICMSRET-DESCONTO),ARMAZEM,DESCRICAOA; //Campo 16
	,LOTE,DTOC(STOD(FABRICACAO)),DTOC(STOD(VALIDADE)),DTOC(STOD(ENTRADA)),DTOC(STOD(EMISSAO)),TES,FINALIDADE,ESTOQUE,PODER3,FINANCEIRO,CFOP; //Campo 26
	,CLIFOR,LOJA,NOME,NREDUZ,ESTADO,PESSOA,CNPJ,TIPOCLI,PEDIDO,ITEMPD,COD_MUN,MUNICIPIO,COD_PAIS,Posicione('SYA',1,xFilial('SYA')+COD_PAIS,'YA_DESCR'),COD_CANAL,CANAL,COD_REGIAO,REGIAO,VENDEDOR,Posicione("SA3",1,xFilial("SA3")+VENDEDOR,"A3_NOME"),CHAVENFE; //Campo 47
	,CCONTABIL,Posicione("CT1",1,xFilial("CT1")+CCONTABIL,"CT1_DESC01"),CCUSTO,Posicione("CTT",1,xFilial("CTT")+CCUSTO,"CTT_DESC01"),DESCONTO,FRETE,TPFRETE,CUSTO; //Campo 55
	,NCM,NFORIGEM,SERIEORI,ITEMORI,Posicione("SC6",1,xFilial("SC6")+PEDIDO+ITEMPD+PRODUTO,"C6_NUMPCOM"),Posicione("SC6",1,xFilial("SC6")+PEDIDO+ITEMPD+PRODUTO,"C6_ITEMPC"),MENNOTA}) //Campo 61
	EndIf */

	(cAlias)->(dbSkip()) //PASSAR PARA O PRÓXIMO REGISTRO DA MINHA QUERY

Enddo

(cAlias)->(dbClosearea()) //FECHO A TABELA APOS O USO

ENDIF

IF MV_PAR05 == 3 .OR. MV_PAR05 == 4 // Estoque ou todos

//Filtro das saídas
cFiltro := "%AND D3_FILIAL BETWEEN '" + ALLTRIM(MV_PAR01) + "' AND '" + ALLTRIM(MV_PAR02) + "' AND D3_EMISSAO BETWEEN '" + DTOS(MV_PAR03) + "' AND '" + DTOS(MV_PAR04) + "' AND D3_COD BETWEEN '" + (MV_PAR10) + "' AND '" + (MV_PAR11) + "' %"

BeginSql Alias cAlias

SELECT 'INTERNAS' AS TIPO, '' AS EMPRESA, D3_FILIAL AS FILIAL, D3_CF AS TIPO1, D3_DOC AS DOCUMENTO, D3_SEQCALC AS SEQUENCIA, D3_COD AS PRODUTO, B1_DESC AS DESCRICAOP, D3_QUANT AS QUANTIDADE
, D3_CUSTO1 AS CUSTO, AH_UMRES AS MEDIDA, D3_LOCAL AS ARMAZEM, NNR_DESCRI AS DESCRICAOA, D3_TM AS TPMOV, F5_TEXTO AS FINALIDADE,  D3_EMISSAO AS EMISSAO, D3_LOTECTL AS LOTE, D3_DTVALID AS VALIDADE, D3_USUARIO AS USUARIO
FROM %table:SD3%
INNER JOIN %table:SB1% ON B1_COD = D3_COD AND %table:SB1%.D_E_L_E_T_ = ''
INNER JOIN %table:NNR% ON D3_LOCAL = NNR_CODIGO AND %table:NNR%.D_E_L_E_T_ = ''
INNER JOIN %table:SAH% ON D3_UM = AH_UNIMED AND %table:SAH%.D_E_L_E_T_ = ''
INNER JOIN %table:SF5% ON D3_TM = F5_CODIGO AND %table:SF5%.D_E_L_E_T_ = ''
WHERE %table:SD3%.D_E_L_E_T_ = ''
%exp:cFiltro%
ORDER BY EMPRESA, FILIAL, DOCUMENTO, SEQUENCIA

EndSql //FINALIZO A MINHA QUERY

cNomeTabela := "SD3 Internas - " + DTOC(MV_PAR03) + " a " + DTOC(MV_PAR04)

	IF MV_PAR05 == 3 //Crio o objeto se ele não foi criado
    //Criando o objeto que irá gerar o conteúdo do Excel
    oFWMsExcel := FWMSExcel():New()
	ENDIF

    //Aba 01 - Mapas
    oFWMsExcel:AddworkSheet("Internas") //Não utilizar número junto com sinal de menos. Ex.: 1-

		//Criando a Tabela
        oFWMsExcel:AddTable("Internas",cNomeTabela)

//Criando Colunas
oFWMsExcel:AddColumn("Internas",cNomeTabela,"TIPO"   		,1,1) //01 //1,1 = Modo Texto  // 2,2 = Valor sem R$  //  3,3 = Valor com R$
oFWMsExcel:AddColumn("Internas",cNomeTabela,"EMPRESA"		,1,1) //02
oFWMsExcel:AddColumn("Internas",cNomeTabela,"FILIAL"	   	,1,1) //03
oFWMsExcel:AddColumn("Internas",cNomeTabela,"TIPO1"		    ,1,1) //04
oFWMsExcel:AddColumn("Internas",cNomeTabela,"DOCUMENTO"  	,1,1) //05
oFWMsExcel:AddColumn("Internas",cNomeTabela,"SEQUENCIA"		,1,1) //06
oFWMsExcel:AddColumn("Internas",cNomeTabela,"PRODUTO"   	,1,1) //07
oFWMsExcel:AddColumn("Internas",cNomeTabela,"DESCRICAOP"	,1,1) //08
oFWMsExcel:AddColumn("Internas",cNomeTabela,"QUANTIDADE"	,2,2) //09
oFWMsExcel:AddColumn("Internas",cNomeTabela,"MEDIDA"	    ,1,1) //10
oFWMsExcel:AddColumn("Internas",cNomeTabela,"CUSTO"  		,2,2) //11
oFWMsExcel:AddColumn("Internas",cNomeTabela,"EMISSAO"	    ,1,1) //12
oFWMsExcel:AddColumn("Internas",cNomeTabela,"TIPOMOV"	    ,1,1) //13
oFWMsExcel:AddColumn("Internas",cNomeTabela,"FINALIDADE"    ,1,1) //14
oFWMsExcel:AddColumn("Internas",cNomeTabela,"ARMAZEM"   	,1,1) //15
oFWMsExcel:AddColumn("Internas",cNomeTabela,"DESCRICAOA"  	,1,1) //16
oFWMsExcel:AddColumn("Internas",cNomeTabela,"LOTE" 	 		,1,1) //17
oFWMsExcel:AddColumn("Internas",cNomeTabela,"VALIDADE"   	,1,1) //18
oFWMsExcel:AddColumn("Internas",cNomeTabela,"USUARIO"   	,1,1) //19

While !(cAlias)->(Eof())

	//Criando as Linhas
	oFWMsExcel:AddRow("Internas",cNomeTabela,{TIPO,EMPRESA,FILIAL,TIPO1,DOCUMENTO,SEQUENCIA,PRODUTO,DESCRICAOP,QUANTIDADE,MEDIDA,CUSTO,DTOC(STOD(EMISSAO)),TPMOV,FINALIDADE,ARMAZEM,DESCRICAOA,LOTE,DTOC(STOD(VALIDADE)),USUARIO})

	(cAlias)->(dbSkip()) //PASSAR PARA O PRÓXIMO REGISTRO DA MINHA QUERY

Enddo

(cAlias)->(dbClosearea()) 				//FECHO A TABELA APOS O USO

//Crio a tabela de saldos iniciais

cFiltro := "%AND B9_FILIAL BETWEEN '" + ALLTRIM(MV_PAR01) + "' AND '" + ALLTRIM(MV_PAR02) + "' AND B9_COD BETWEEN '" + ALLTRIM(MV_PAR10) + "' AND '" + ALLTRIM(MV_PAR11) + "' %"

BeginSql Alias cAlias

SELECT 'SLD_INICIAL' AS TIPO, '' AS EMPRESA, B9_FILIAL AS FILIAL, B9_COD AS PRODUTO, B1_DESC AS DESCRICAOP, B9_QINI AS QTDINICIAL
, B9_VINI1 AS VLRINICIAL, B9_LOCAL AS ARMAZEM, NNR_DESCRI AS DESCRICAOA, B9_DATA AS DATA
FROM %table:SB9%
INNER JOIN %table:SB1% ON B1_COD = B9_COD AND %table:SB1%.D_E_L_E_T_ = ''
INNER JOIN %table:NNR% ON B9_FILIAL = NNR_FILIAL AND B9_LOCAL = NNR_CODIGO AND %table:NNR%.D_E_L_E_T_ = ''
WHERE %table:SB9%.D_E_L_E_T_ = '' AND (B9_VINI1 <> 0 OR B9_QINI <> 0)
%exp:cFiltro%

EndSql //FINALIZO A MINHA QUERY

cNomeTabela := "SB9 Saldos Iniciais"

	IF MV_PAR05 == 3 //Crio o objeto se ele não foi criado
    //Criando o objeto que irá gerar o conteúdo do Excel
    oFWMsExcel := FWMSExcel():New()
	ENDIF

    //Aba 01 - Mapas
    oFWMsExcel:AddworkSheet("Saldo_Inicial") //Não utilizar número junto com sinal de menos. Ex.: 1-

		//Criando a Tabela
        oFWMsExcel:AddTable("Saldo_Inicial",cNomeTabela)

//Criando Colunas
oFWMsExcel:AddColumn("Saldo_Inicial",cNomeTabela,"TIPO"   		,1,1) //01 //1,1 = Modo Texto  // 2,2 = Valor sem R$  //  3,3 = Valor com R$
oFWMsExcel:AddColumn("Saldo_Inicial",cNomeTabela,"EMPRESA"		,1,1) //02
oFWMsExcel:AddColumn("Saldo_Inicial",cNomeTabela,"FILIAL"	   	,1,1) //03
oFWMsExcel:AddColumn("Saldo_Inicial",cNomeTabela,"PRODUTO"   	,1,1) //07
oFWMsExcel:AddColumn("Saldo_Inicial",cNomeTabela,"DESCRICAOP"	,1,1) //08
oFWMsExcel:AddColumn("Saldo_Inicial",cNomeTabela,"QTDINICIAL"	,2,2) //09
oFWMsExcel:AddColumn("Saldo_Inicial",cNomeTabela,"VLRINICIAL"   ,2,2) //10
oFWMsExcel:AddColumn("Saldo_Inicial",cNomeTabela,"DATA"		    ,1,1) //12
oFWMsExcel:AddColumn("Saldo_Inicial",cNomeTabela,"ARMAZEM"   	,1,1) //15
oFWMsExcel:AddColumn("Saldo_Inicial",cNomeTabela,"DESCRICAOA"  	,1,1) //16

While !(cAlias)->(Eof())

	//Criando as Linhas
	oFWMsExcel:AddRow("Saldo_Inicial",cNomeTabela,{TIPO,EMPRESA,FILIAL,PRODUTO,DESCRICAOP,QTDINICIAL,VLRINICIAL,DTOC(STOD(DATA)),ARMAZEM,DESCRICAOA})

	(cAlias)->(dbSkip()) //PASSAR PARA O PRÓXIMO REGISTRO DA MINHA QUERY

Enddo

(cAlias)->(dbClosearea()) //FECHO A TABELA APOS O USO

ENDIF

    oFWMsExcel:Activate()
    oFWMsExcel:GetXMLFile(cArquivo)

    //Abrindo o excel e abrindo o arquivo xml
    oExcel:= MsExcel():New()            	//Abre uma nova conexão com Excel
    oExcel:WorkBooks:Open(cArquivo)     	//Abre uma planilha
    oExcel:SetVisible(.T.)              	//Visualiza a planilha
    oExcel:Destroy()                    	//Encerra o processo do gerenciador de tarefas

	RestArea(aArea)

Return

//Validação cPerg

Static Function ValidPerg(cPerg)

	Local aArea	:= GetArea()
	Local aRegs	:= {}
	Local i,j

	DbSelectArea("SX1")
	SX1->(DbSetOrder(1))

	cPerg := PADR(cPerg,10)

//Paramtros que criei para o relatorio, não precisa mexer, pode escrever continua exp
AADD(aRegs,{cPerg,"01","Da Filial:			","","","mv_ch01","C",02,0,0,"G","","mv_par01","","","","","","","","","","","","","","","","","","","","","","","","","SM0"})
AADD(aRegs,{cPerg,"02","Até a Filial: 		","","","mv_ch02","C",02,0,0,"G","","mv_par02","","","","","","","","","","","","","","","","","","","","","","","","","SM0"})
AADD(aRegs,{cPerg,"03","Da Data:  			","","","mv_ch03","D",08,0,0,"G","","mv_par03","","","","","","","","","","","","","","","","","","","","","","","","",""})
AADD(aRegs,{cPerg,"04","Até a Data: 		","","","mv_ch04","D",08,0,0,"G","","mv_par04","","","","","","","","","","","","","","","","","","","","","","","","",""})
AADD(aRegs,{cPerg,"05","TP Relatorio: 		","","","mv_ch05","N",01,0,0,"C","","mv_par05","Entradas","","","","","Saidas","","","","","Estoque","","","","","Todos","","","","","","","","",""})
AADD(aRegs,{cPerg,"06","Do Fornecedor:		","","","mv_ch06","C",06,0,0,"G","","mv_par06","","","","","","","","","","","","","","","","","","","","","","","","","SA2"})
AADD(aRegs,{cPerg,"07","Até o Fornecedor: 	","","","mv_ch07","C",06,0,0,"G","","mv_par07","","","","","","","","","","","","","","","","","","","","","","","","","SA2"})
AADD(aRegs,{cPerg,"08","Do Cliente:			","","","mv_ch08","C",06,0,0,"G","","mv_par08","","","","","","","","","","","","","","","","","","","","","","","","","SA1"})
AADD(aRegs,{cPerg,"09","Até o Cliente: 		","","","mv_ch09","C",06,0,0,"G","","mv_par09","","","","","","","","","","","","","","","","","","","","","","","","","SA1"})
AADD(aRegs,{cPerg,"10","Do Produto:			","","","mv_ch10","C",15,0,0,"G","","mv_par10","","","","","","","","","","","","","","","","","","","","","","","","","SB1"})
AADD(aRegs,{cPerg,"11","Até o Produto: 		","","","mv_ch11","C",15,0,0,"G","","mv_par11","","","","","","","","","","","","","","","","","","","","","","","","","SB1"})
AADD(aRegs,{cPerg,"12","Da TES:				","","","mv_ch12","C",03,0,0,"G","","mv_par12","","","","","","","","","","","","","","","","","","","","","","","","","SF4"})
AADD(aRegs,{cPerg,"13","Até a TES: 			","","","mv_ch13","C",03,0,0,"G","","mv_par13","","","","","","","","","","","","","","","","","","","","","","","","","SF4"})
AADD(aRegs,{cPerg,"14","Imprime Impostos:	","","","mv_ch14","N",01,0,0,"C","","mv_par14","Sim","","","","","Não","","","","","","","","","","","","","","","","","","",""})

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
