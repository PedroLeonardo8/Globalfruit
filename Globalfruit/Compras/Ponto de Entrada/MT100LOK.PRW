#INCLUDE "PROTHEUS.CH"

User Function MT100LOK()

&& Deixa sistema posicionado para usar informações nas validações
Local xArea		:= GetArea()
Local xAreaSF4	:= SF4->(GetArea())
Local xAreaSB1	:= SB1->(GetArea())
Local xAreaSED	:= SED->(GetArea())
Local xAreaCTT	:= CTT->(GetArea())

&& Variáveis para apoio de posicionamento / processamentos
Private lRet 		:= .t.
Private cCodProd	:= alltrim(aCols[ n, GDFieldPos("D1_COD")])
Private cCodTes		:= alltrim(aCols[ n, GDFieldPos("D1_TES")])
Private cConta		:= alltrim(aCols[ n, GDFieldPos("D1_CONTA")])
Private cCusto		:= alltrim(aCols[ n, GDFieldPos("D1_CC")])
Private cItemCta	:= alltrim(aCols[ n, GDFieldPos("D1_ITEMCTA")])
Private cClVl		:= alltrim(aCols[ n, GDFieldPos("D1_CLVL")])
Private lRateio		:= Iif(alltrim(aCols[ n, GDFieldPos("D1_RATEIO")])=="1",.t.,.f.)
Private cNumOP		:= alltrim(aCols[ n, GDFieldPos("D1_OP")])
Private cClasFis	:= alltrim(aCols[ n, GDFieldPos("D1_CLASFIS")])
Private cRN1CC		:= ""
Private cTipoEst	:= U_GCT5("P_TIPO-PROD-ESTOQUE")


&& Posiciona nas tabelas
fPosMt100()

&& --- Executa validações
lRet := U_FVDEI()

&& Restaura áreas de trabalho
RestArea(xAreaCTT)
RestArea(xAreaSED)
RestArea(xAreaSB1)
RestArea(xAreaSF4)
RestArea(xArea)

Return lRet

Static Function fPosMt100()

&& Posiciona no cadastro de Produtos
dbSelectArea("SB1")
dbSetOrder(1)
If !dbSeek(xFilial("SB1")+cCodProd)
	MsgStop("Erro de posicionamento no alias -> " + Alias())
EndIf

&& Posiciona no cadastro de TES
dbSelectArea("SF4")
dbSetOrder(1)
if !dbSeek(xFilial("SF4")+cCodTes)
	MsgStop("Erro de posicionamento no alias -> " + Alias())
EndIf

&& Posiciona no cadastro de C.Custo
cRN1CC := ""
If !Empty(cCusto)
	dbSelectArea("CTT")
	dbSetOrder(1)
	If !dbSeek(xFilial("CTT")+cCusto)
		MsgStop("Erro de posicionamento no alias -> " + Alias())
	Else
		cRN1CC	:= CTT->CTT_CRGNV1
	EndIf
EndIf

Return
