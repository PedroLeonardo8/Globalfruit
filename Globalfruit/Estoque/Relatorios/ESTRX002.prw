#INCLUDE "TOPCONN.CH"
#INCLUDE "PROTHEUS.CH"

//Rafael França - 29/10/2021 - Relatório para saber o saldo do Kit de acordo com saldos dos componentes/partes
//SIGAEST -> Relatorios -> Especificos -> Saldo por Kit

User Function ESTRX002()

Private cAlias 		:= GetNextAlias() //Declarei meu ALIAS
Private aArea       := GetArea()
Private oFWMsExcel
Private oExcel
Private cArquivo    := GetTempPath()+'ESTRX002.xml'
Private cPerg  		:= "ESTRX002"
Private cNomePlan	:= UPPER("Saldos")
Private cNomeTabela	:= UPPER("Saldos por Kit")
Private cFiltro		:= ""
Private nSaldoSB2	:= 0
Private nSaldoTot	:= 0
Private cProduto	:= ""

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

If MV_PAR05 == 1
cNomeTabela	:= UPPER("Saldos por Kit - Detalhado")
Else
cNomeTabela	:= UPPER("Saldos por Kit - Resumido")
EndIf

FwMsgRun(Nil, { || ESTRX002A() }, "Processando", "Gerando planilha xml..." )

Return

Static Function ESTRX002A

//TRATO AS PERGUNTAS PARA USO NOS FILTROS
cFiltro		:= "% AND SB1P.B1_LOCPAD BETWEEN '" + (MV_PAR01) + "' AND '" + (MV_PAR02) + "' AND SB1P.B1_COD BETWEEN '" + (MV_PAR03) + "' AND '" + (MV_PAR04) + "' %"

//COMEÇO A MINHA CONSULTA SQL
BeginSql Alias cAlias

//QUERY
//"SELECT" -> Seleciono os dados/campos para retorno da query. Posso usar a expressão "AS" para dar apelido aos campos. Se usar apenas "*" retorna todos os campos da tabela e deixa a consulta mais lenta.
SELECT SB1P.B1_COD AS PRODUTO, SB1P.B1_DESC AS DESCRICAO, (B2_QATU / G1_QUANT) AS SALDO_ESTRUTURA, G1_COMP AS PARTE, SB1C.B1_DESC AS DESCRI_PARTE, G1_QUANT AS QTD_ESTRUTURA, SB1C.B1_LOCPAD AS ARMAZEM, SB1C.B1_UM AS UNIDADE, B2_QATU AS SALDO1, SB1C.B1_SEGUM AS UNIDADE2, B2_QTSEGUM AS SALDO2
//"FROM" -> Determino a tabela principal da minha consulta. Os primeiros três digitos são a tabela "SB1" (Produtos) os proximos dois são a empresa "01" e o ultimo é um campo de controle da Totvs que é sempre "0".
FROM %table:SB1% SB1P
//"INNER JOIN" - Faço a união com outras tabelas a partir de uma chave de relacionamento. No "ON" vc colocar a chave primaria da tabela de retorno, assim vc pode usar os campos dessa tabela no seu "SELECT"
INNER JOIN %table:SG1% ON SB1P.B1_COD = G1_COD AND %table:SG1%.D_E_L_E_T_ = ''
INNER JOIN %table:SB1% SB1C ON SB1C.B1_COD = G1_COMP AND SB1C.D_E_L_E_T_ = ''
INNER JOIN %table:SB2% ON G1_COMP = B2_COD AND SB1C.B1_LOCPAD = B2_LOCAL AND %table:SB2%.D_E_L_E_T_ = ''
//"WHERE" - Condição, crio o filtro para trazer apenas os registros desejados.
WHERE SB1P.D_E_L_E_T_ = '' AND SB1P.B1_UM = 'KT'
%exp:cFiltro%
ORDER BY PRODUTO, SALDO_ESTRUTURA

EndSql //FINALIZO A MINHA QUERY

    //Criando o objeto que irá gerar o conteúdo do Excel
    oFWMsExcel := FWMSExcel():New()

    //Aba 01 - Nome Guia
    oFWMsExcel:AddworkSheet(cNomePlan) //Não utilizar número junto com sinal de menos. Ex.: 1-
        //Criando a Tabela
        oFWMsExcel:AddTable(cNomePlan,cNomeTabela)
        //Criando Colunas
        oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"PRODUTO",1,1) //1,1 = Modo Texto  // 2,2 = Valor sem R$  //  3,3 = Valor com R$
        oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"DESCRICAO",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"SALDO_KIT",2,2)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"SALDO_ESTRUTURA",2,2)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"SALDO_TOTAL",2,2)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"PARTE",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"DESCRI_PARTE",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"QTD_ESTRUTURA",2,2)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"LOCAL",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"UNIDADE",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"SALDO_UM1",2,2)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"UNIDADE2",1,1)
		oFWMsExcel:AddColumn(cNomePlan,cNomeTabela,"SALDO_UM2",2,2)

While !(cAlias)->(Eof())

nSaldoSB2 := POSICIONE("SB2",1,XFilial("SB2")+PRODUTO,"B2_QATU")
nSaldoTot := nSaldoSB2 + SALDO_ESTRUTURA

If cProduto <> PRODUTO .OR. MV_PAR05 == 1
//Adciono linhas de acordo com o retorno da query, um campo por coluna
oFWMsExcel:AddRow(cNomePlan,cNomeTabela,{PRODUTO,DESCRICAO,nSaldoSB2,SALDO_ESTRUTURA,nSaldoTot,PARTE,DESCRI_PARTE,QTD_ESTRUTURA,ARMAZEM,UNIDADE,SALDO1,UNIDADE2,SALDO2})
EndIf

cProduto := PRODUTO

	(cAlias)->(dbSkip()) //PASSAR PARA O PRÓXIMO REGISTRO DA MINHA QUERY

Enddo

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

	AADD(aRegs,{cPerg,"01","Do Armazém:			","","","mv_ch01","C",02,0,0,"G","","mv_par01","","","","","","","","","","","","","","","","","","","","","","","","","NNR"})
	AADD(aRegs,{cPerg,"02","Até o Armazém:		","","","mv_ch02","C",02,0,0,"G","","mv_par02","","","","","","","","","","","","","","","","","","","","","","","","","NNR"})
	AADD(aRegs,{cPerg,"03","Do Produto:			","","","mv_ch03","C",15,0,0,"G","","mv_par03","","","","","","","","","","","","","","","","","","","","","","","","","SB1"})
	AADD(aRegs,{cPerg,"04","Até o Produto:		","","","mv_ch04","C",15,0,0,"G","","mv_par04","","","","","","","","","","","","","","","","","","","","","","","","","SB1"})
	AADD(aRegs,{cPerg,"05","Tipo do Relatório:	","","","mv_ch05","N",01,0,0,"C","","mv_par05","Detalhado","","","","","Resumido","","","","","","","","","","","","","","","","","","",""})

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
