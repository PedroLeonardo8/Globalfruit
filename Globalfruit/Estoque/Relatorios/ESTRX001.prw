#INCLUDE "TOPCONN.CH"
#INCLUDE "PROTHEUS.CH"

//Rafael França - 18/05/2021 - Relatorio para saber o periodo que os itens ficam em estoque sem uso
//SIGAEST -> Relatorios -> Especificos -> Empenho por OP

User Function ESTRX001()

Private cAlias 		:= GetNextAlias() //Declarei meu ALIAS
Private aArea       := GetArea()
Private oFWMsExcel
Private oExcel
Private cArquivo    := GetTempPath()+'ESTRX001.xml'
Private cPerg  		:= "ESTRX001"
Private cNomeTabela	:= UPPER("Empenho")
Private cNomePlan	:= UPPER("Empenho por OP")
Private cFiltro		:= ""
Private c2Unidade 	:= ""

	ValidPerg(cPerg) //INICIA A STATIC FUNCTION PARA CRIAÇÃO DAS PERGUNTAS

	If !Pergunte(cPerg) //Verifica se usuario respondeu as perguntas ou cancelou
		MsgAlert("Operação Cancelada!")
		RestArea(aArea)
		Return
	EndIf

If !ApOleClient("MsExcel")
	MsgStop("Microsoft Excel nao instalado.")  //"Microsoft Excel nao instalado."
	RestArea(aArea)
	Return
EndIf

FwMsgRun(Nil, { || ESTRX001A() }, "Processando", "Gerando planilha xml..." )

Return

Static Function ESTRX001A

//TRATO AS PERGUNTAS PARA USO NOS FILTROS
If MV_PAR05 == 1
cFiltro		:= "% AND C2_QUJE = 0 AND (C2_NUM + C2_ITEM + C2_SEQUEN) BETWEEN '" + (MV_PAR01) + "' AND '" + (MV_PAR02) + "' %"
Else
cFiltro		:= "% AND C2_QUJE = 0 AND (C2_NUM + C2_ITEM + C2_SEQUEN) BETWEEN '" + (MV_PAR01) + "' AND '" + (MV_PAR02) + "' AND B1_TIPO NOT IN ('PI','PP') %"
EndIf

If MV_PAR06 == 1

//COMEÇO A MINHA CONSULTA SQL
BeginSql Alias cAlias

//QUERY
SELECT C2_FILIAL AS FILIAL, (C2_NUM + C2_ITEM + C2_SEQUEN) AS OP, C2_PRODUTO AS PROD_OP
, C2_LOCAL AS LOCAL_OP, C2_QUANT AS QTD_OP, C2_DATPRI AS PREV_INI, C2_DATPRF AS ENTREGA, C2_EMISSAO AS EMISSAO
, D4_COD AS PROD_EMP, D4_LOCAL AS LOCAL_EMP, D4_DATA AS DT_EMP, D4_QTDEORI AS QTD_EMP, B1_UM AS UNIDADE, D4_QTSEGUM AS QTDSEG, B1_SEGUM AS UNIDADE2, B1_DESC AS DESCR_EMP
FROM %table:SC2%
INNER JOIN %table:SD4% ON %table:SD4%.D_E_L_E_T_ = '' AND (C2_NUM + C2_ITEM + C2_SEQUEN)  = TRIM(D4_OP) AND C2_FILIAL = D4_FILIAL
INNER JOIN %table:SB1% ON %table:SB1%.D_E_L_E_T_ = '' AND D4_COD = B1_COD
WHERE %table:SC2%.D_E_L_E_T_ = ''
%exp:cFiltro%
ORDER BY C2_NUM, C2_ITEM, C2_SEQUEN, D4_COD

EndSql //FINALIZO A MINHA QUERY

    //Criando o objeto que irá gerar o conteúdo do Excel
    oFWMsExcel := FWMSExcel():New()

    //Aba 01 - Nome Guia
    oFWMsExcel:AddworkSheet(cNomePlan) //Não utilizar número junto com sinal de menos. Ex.: 1-
        //Criando a Tabela
        oFWMsExcel:AddTable(cNomePlan,cNomeTabela)
        //Criando Colunas
        oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"OP",1,1) //1,1 = Modo Texto  // 2,2 = Valor sem R$  //  3,3 = Valor com R$
        //oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"PRODUTO",1,1)
        //oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"DESCRICAO",1,1)
        //oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"ARMAZEM",1,1)
        //oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"QTD_OP",2,2)
    	//oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"UNIDADE_OP",1,1)
        oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"PREV_INICIAL",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"PREV_FINAL",1,1)
        oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"PROD_EMP",1,1)
        oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"DESCR_EMP",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"QTD_EMP",2,2)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"UNIDADE_EMP",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"2QTD_EMP",2,2)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"2UNIDADE_EMP",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"LOCAL_EMP",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"ENTREGA",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"RETORNO",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"LOTE",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"CONSUMO",2,2)


While !(cAlias)->(Eof())

If QTDSEG <> 0
c2Unidade := UNIDADE2
Else
c2Unidade := ""
Endif

//oFWMsExcel:AddRow(cNomePlan,cNomeTabela,{OP,PROD_OP,Posicione("SB1",1,xFilial("SB1")+PROD_OP,"B1_DESC"),LOCAL_OP,QTD_OP,Posicione("SB1",1,xFilial("SB1")+PROD_OP,"B1_UM"),DTOC(STOD(PREV_INI)),DTOC(STOD(ENTREGA)),PROD_EMP,DESCR_EMP,QTD_EMP,UNIDADE,LOCAL_EMP,"","","",0})
oFWMsExcel:AddRow(cNomePlan,cNomeTabela,{OP,DTOC(STOD(PREV_INI)),DTOC(STOD(ENTREGA)),PROD_EMP,DESCR_EMP,QTD_EMP,UNIDADE,QTDSEG,c2Unidade,LOCAL_EMP,"","","",0})

	(cAlias)->(dbSkip()) //PASSAR PARA O PRÓXIMO REGISTRO DA MINHA QUERY

Enddo

Else MV_PAR06 == 2

//COMEÇO A MINHA CONSULTA SQL
BeginSql Alias cAlias

//QUERY
SELECT C2_FILIAL AS FILIAL, D4_COD AS PROD_EMP, D4_LOCAL AS LOCAL_EMP, D4_QTDEORI AS QTD_EMP, B1_UM AS UNIDADE, D4_QTSEGUM AS QTDSEG, B1_SEGUM AS UNIDADE2, B1_DESC AS DESCR_EMP
FROM %table:SC2%
INNER JOIN %table:SD4% ON %table:SD4%.D_E_L_E_T_ = '' AND (C2_NUM + C2_ITEM + C2_SEQUEN)  = TRIM(D4_OP) AND C2_FILIAL = D4_FILIAL
INNER JOIN %table:SB1% ON %table:SB1%.D_E_L_E_T_ = '' AND D4_COD = B1_COD
WHERE %table:SC2%.D_E_L_E_T_ = ''
%exp:cFiltro%
ORDER BY D4_COD, D4_LOCAL

EndSql //FINALIZO A MINHA QUERY

    //Criando o objeto que irá gerar o conteúdo do Excel
    oFWMsExcel := FWMSExcel():New()

    //Aba 01 - Nome Guia
    oFWMsExcel:AddworkSheet(cNomePlan) //Não utilizar número junto com sinal de menos. Ex.: 1-
        //Criando a Tabela
        oFWMsExcel:AddTable(cNomePlan,cNomeTabela)
        //Criando Colunas
        oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"PROD_EMP",1,1)
        oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"DESCR_EMP",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"QTD_EMP",2,2)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"UNIDADE_EMP",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"2QTD_EMP",2,2)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"2UNIDADE_EMP",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"LOCAL_EMP",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"ENTREGA",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"RETORNO",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"LOTE",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"CONSUMO",2,2)


While !(cAlias)->(Eof())

If QTDSEG <> 0
c2Unidade := UNIDADE2
Else
c2Unidade := ""
Endif

//oFWMsExcel:AddRow(cNomePlan,cNomeTabela,{OP,PROD_OP,Posicione("SB1",1,xFilial("SB1")+PROD_OP,"B1_DESC"),LOCAL_OP,QTD_OP,Posicione("SB1",1,xFilial("SB1")+PROD_OP,"B1_UM"),DTOC(STOD(PREV_INI)),DTOC(STOD(ENTREGA)),PROD_EMP,DESCR_EMP,QTD_EMP,UNIDADE,LOCAL_EMP,"","","",0})
oFWMsExcel:AddRow(cNomePlan,cNomeTabela,{OP,DTOC(STOD(PREV_INI)),DTOC(STOD(ENTREGA)),PROD_EMP,DESCR_EMP,QTD_EMP,UNIDADE,QTDSEG,c2Unidade,LOCAL_EMP,"","","",0})

	(cAlias)->(dbSkip()) //PASSAR PARA O PRÓXIMO REGISTRO DA MINHA QUERY

Enddo

EndIf

    oFWMsExcel:Activate()
    oFWMsExcel:GetXMLFile(cArquivo)

    //Abrindo o excel e abrindo o arquivo xml
    oExcel:= MsExcel():New()            	//Abre uma nova conexão com Excel
    oExcel:WorkBooks:Open(cArquivo)     	//Abre uma planilha
    oExcel:SetVisible(.T.)              	//Visualiza a planilha
    oExcel:Destroy()                    	//Encerra o processo do gerenciador de tarefas

	(cAlias)->(dbClosearea()) 				//FECHO A TABELA APOS O USO

	RestArea(aArea)

Return

//Programa usado para criar perguntas na tabela SX1 (Tabela de perguntas)
Static Function ValidPerg(cPerg)

	Local aArea	:= GetArea()
	Local aRegs	:= {}
	Local i,j

	_sAlias := Alias()
	cPerg := PADR(cPerg,10)
	dbSelectArea("SX1")
	dbSetOrder(1)
	aRegs:={}

	AADD(aRegs,{cPerg,"01","Da OP:				","","","mv_ch01","C",12,0,0,"G","","mv_par01","","","","","","","","","","","","","","","","","","","","","","","","","SC2"})
	AADD(aRegs,{cPerg,"02","Até a OP:			","","","mv_ch02","C",12,0,0,"G","","mv_par02","","","","","","","","","","","","","","","","","","","","","","","","","SC2"})
	AADD(aRegs,{cPerg,"03","Da Data:			","","","mv_ch03","D",08,0,0,"G","","mv_par03","","","","","","","","","","","","","","","","","","","","","","","","",""})
	AADD(aRegs,{cPerg,"04","Até a Data:			","","","mv_ch04","D",08,0,0,"G","","mv_par04","","","","","","","","","","","","","","","","","","","","","","","","",""})
	AADD(aRegs,{cPerg,"05","Impr. Prod. PI:		","","","mv_ch05","N",01,0,0,"C","","mv_par05","Sim","","","","","Não","","","","","","","","","","","","","","","","","","",""})
	AADD(aRegs,{cPerg,"06","Tipo do Relatório:	","","","mv_ch06","N",01,0,0,"C","","mv_par06","Detalhado","","","","","Resumido","","","","","","","","","","","","","","","","","","",""})

	For i:=1 to Len(aRegs)
		If !dbSeek(cPerg+aRegs[i,2])
			RecLock("SX1",.T.)
			For j:=1 to FCount()
				If j <= Len(aRegs[i])
					FieldPut(j,aRegs[i,j])
				EndIf
			Next
			MsUnlock()
		EndIf
	Next

	RestArea(aArea)

Return
