#INCLUDE "rwmake.ch"
#INCLUDE "Protheus.ch"

//Rafael Fran�a - 15/10/2021 - Rotina de importa��o e valida��o de codigo de barras para conferencia e desmontagem dos produtos

User Function AXZA1

//Local cVldAlt := ".T." // Validacao para permitir a alteracao. Pode-se utilizar ExecBlock.
//Local cVldExc := ".T." // Validacao para permitir a exclusao. Pode-se utilizar ExecBlock.

Private cCadastro 	:= "Importa��o Codigo de Barras"
Private aRotina   	:= {	{"Pesquisar" 	,"AxPesqui"	,0,1} ,;
						 	{"Visualizar"	,"AxVisual"	,0,2} ,;
							{"Incluir"	 	,"AxInclui"	,0,3} ,;
							{"Alterar"	 	,"AxAltera"	,0,4} ,;
							{"Excluir"	 	,"AxDeleta"	,0,5} ,;
							{"Importar .csv","ImportZA1",0,4} ,;
							{"Desmontar"	,"MT242EXEC",0,4}}

Private cDelFunc  	:= ".T." // Validacao para a exclusao. Pode-se utilizar ExecBlock
Private cString   	:= "ZA1"

dbSelectArea(cString)
dbSetOrder(1)
mBrowse( 6,1,22,75,cString,,,,,,)

Return

/*
User Function AxAlt1()

Private cPedido 	:= SZL->ZL_PEDIDO
Private cSolic		:= SZL->ZL_SOLICIT
Private cLib := Posicione("SCR",1,xFilial("SCR") + "PC" + cPedido,"CR_DATALIB")

IF EMPTY(cLib) .OR. AllTrim(cUserName) $ "Administrador"

dbSelectArea("SZL")
dbSetOrder (2)
dbSeek(xFilial("SZL") + cPedido )
AxAltera("SZL",SZL->(Recno()),4,,,,,".T.",,,,,,,.F.)

ELSE

MsgInfo("Pedidos liberados n�o podem ser alterados","Aten��o!")
Return

ENDIF

Return
*/
