#INCLUDE "PROTHEUS.CH"

User Function FVDEI()

Local lVldOn	:= SUPERGETMV("ZZ_KITVLD", .f., .f.)

&& Verifica se executa valida��es.
If !lVldOn
	Return(.t.)
Endif

&& *** Importante: As tabelas devem estar posicionadas para as valida��es abaixo ***

lRet := Vld01() && Retorna conta cont�bil e valida a mesma.
lRet := Vld02() && Valida se o TES esta de acordo com o tipo de documento aceito.
lRet := Vld03() && Valida tipo de produto x TES que atualiza estoque
lRet := Vld04() && Valida entrada de NF de servi�o de beneficiamento x TES x OP
lRet := Vld05() && Valida Esp�cie x Formul�rio Pr�prio
lRet := Vld06() && Valida Quantidade de dias da emiss�o da NF
lRet := Vld07() && Valida Informa��es para DIRF
lRet := Vld08() && Valida Especie x CFOP de Frete
lRet := Vld09() && Valida Especie x Escritura��o ICMS / IPI
lRet := Vld10() && Valida NF Complemento ICMS x Tributa��o de PIS/Cofins
lRet := Vld11() && Valida Especie x CFOP de Energia
lRet := Vld12() && Valida Especie x CFOP de Serv.Comunica��o
lRet := Vld13() && Valida Especie Nota Servi�o x Escritura��o no Livro
//lRet := Vld14() && Valida Especie Nota Servi�o x C�d. ISS no produto
lRet := Vld15() && Valida o tamanho do n�mero da NF
lRet := Vld16() && Valida se a classifica��o fiscal esta correta
lRet := Vld17() && Valida se a especie foi preenchida
//lRet := Vld18() && Valida a situa��o tribut�ria do IPI
//lRet := Vld19() && Valida a situa��o tribut�ria do PIS/COFINS
//lRet := Vld20() && Valida a situa��o tribut�ria do ISS
lRet := Vld21() && Valida retorno de beneficiamento utilizado no processo de beneficiamento

Return lRet

&& -------------------------------------------------------------

Static Function Vld01

If !lRet
	Return lRet
EndIf

&& N�o valida caso TES n�o contabiliza
If SF4->F4_ZZCTB == "2"
	Return .t.
EndIf

&& N�o valida rateio
If lRateio
	Return .t.
EndIf

&& Obt�m a conta cont�bil no produto, j� que n�o foi informada no documento de entrada. Prioriza B1_CONTA para estoque e depois os campos customizados.
If Empty(cConta)
	If SF4->F4_ESTOQUE=="S"
		cConta := SB1->B1_CONTA
	Else
		If !Empty(cRN1CC)
			cConta := U_GCT5("P_SB1-CTA-"+cRN1CC)
		EndIf
	EndIf
Endif

&& Valida se foi informada a conta cont�bil.
If Empty(cConta) .And. (SF4->F4_ESTOQUE=="S" .Or. SF4->F4_DUPLIC=="S")
	lRet := .f.
	MsgStop("Conta cont�bil n�o informada. Se for uma opera��o de despesa/custos informe o centro de custo para retornar a conta." + ;
	" Se informou o c.custo e n�o veio conta, ent�o o produto esta com conta em branco.","#MT100LOK Inconsist�nia Cta. Cont�bil")
EndIf

&& Valida obrigatoriedade de entidades
If lRet
	lRet := CtbObrig(cConta,cCusto,cItemCta,cClVl,.t.)
EndIf

&& Valida regras de amarra��o
If lRet
	lRet := CtbAmarra(cConta,cCusto,cItemCta,cClVl,.t.)
EndIf

&& Retorna conta para o t�tulo
If lRet
	If Empty(aCols[ n, GDFieldPos("D1_CONTA")]) .And. !Empty(cConta)
		aCols[ n, GDFieldPos("D1_CONTA")] := padr(cConta,TamSx3( "CT1_CONTA" )[1])
	EndIf
EndIf

Return lRet

&& -------------------------------------------------------------

Static Function Vld02

Local cTpAceito := ""
Local cChvZT5	:= ""

If !lRet
	Return lRet
EndIf

cChvZT5 := "P_TIPO-CF-" + If(Left(SF4->F4_CODIGO,1)<="500","E","S") + SubStr(SF4->F4_CF,2,3)

cTpAceito := U_GCT5(cChvZT5)

If !Empty(cTpAceito) .And. !cTipo $ cTpAceito
	MsgStop("O tipo de movimento da NF [ "+cTipo+" ] n�o esta de acordo com os tipos de movimentos aceitos pelo CFOP [ "+cTpAceito + ;
	" ]. Verifique a movimenta��o e escolha o tipo de movimento correto ou TES correto. Para suporte procurar pela �rea Fiscal.","#MT100LOK Tipo Mov.")
	lRet := .f.
EndIf

Return lRet

&& -------------------------------------------------------------

Static Function Vld03

If !lRet
	Return lRet
EndIf

// Valida tipo de produto e se o TES alimenta estoques.
If !SB1->B1_TIPO $ cTipoEst .And. SF4->F4_ESTOQUE == "S"
	MsgStop("De acordo com as regras do SPED Fiscal este tipo de produto [ "+SB1->B1_TIPO+" ] n�o pode movimentar estoque."+;
	" Tipos permitidos [ "+cTipoEst+"] ","#MT100LOK Tipo produto x Estoques")
	lRet := .f.
EndIf

Return lRet

&& -------------------------------------------------------------

Static Function Vld04

If !lRet
	Return lRet
EndIf

// Se for servi�o de beneficiamento e o TES atualizar estoque, ver se a OP foi informada.
If SubStr(SF4->F4_CF,2,3) $ U_GCT5("P_CFOP-SERV-BENEF") .And. SF4->F4_ESTOQUE=="S" .And. Empty(cNumOP)
	MsgStop("Para servi�o de beneficiamento e TES que atualiza estoques, informar o n�mero da OP." ,"#MT100LOK Serv.Benef x OP")
	lRet := .f.
Endif

Return lRet

&& -------------------------------------------------------------

Static Function Vld05

If !lRet
	Return lRet
EndIf

If alltrim(cFormul) == "S" .And. alltrim(cEspecie) <> "SPED"
	MsgStop("O campo 'Esp�cie' deve ser preenchido com 'SPED' quando usar formul�rio pr�prio.","#MT100LOK Esp�cie x Formul�rio Pr�prio")
	lRet := .F.
EndIf

Return lRet

&& -------------------------------------------------------------

Static Function Vld06

If !lRet
	Return lRet
EndIf

If DDEMISSAO < DDATABASE-15
	lRet := MsgYesNo("A Data de Emiss�o da Nota � inferior h� 15 dias da DATABASE. Esta correto?", "#MT100LOK Idade da NF")
EndIf

Return lRet

&& -------------------------------------------------------------

Static Function Vld07

If !lRet
	Return lRet
EndIf

If ( MaFisRet(,"NF_VALIRR") + MaFisRet(,"NF_VALCOF") + MaFisRet(,"NF_VALPIS") + MaFisRet(,"NF_VALCSL") ) > 0
	If cDirf=="2"
		MsgStop("Para opera��es com IR, PIS, COFINS, CSLL o campo 'Gera Dirf' deve ser preenchido com 'Sim'.","#MT100LOK Gera Dirf")
		lRet := .f.
	EndIf
	If Empty(cCodRet) .And. lRet
		MsgStop("Para opera��es com IR, PIS, COFINS, CSLL o campo 'Cd.Retencao' deve ser preenchido com o codigo do DARF.","#MT100LOK c�digo do DARF")
		lRet := .f.
	EndIf
EndIf

Return lRet

&& -------------------------------------------------------------

Static Function Vld08

If !lRet
	Return lRet
EndIf

If SubStr(SF4->F4_CF,2,3) $ U_GCT5("P_CFOP-FRETE") .And. !Alltrim(CESPECIE) $ U_GCT5("P_ESPECIE-FRETE")
	MsgStop("Para frete a esp�cie de documento deve ser uma destas "+U_GCT5("P_ESPECIE-FRETE"),"#MT100LOK Esp�cie Documento")
	lRet := .F.
Endif

Return lRet

&& -------------------------------------------------------------

Static Function Vld09

If !lRet
	Return lRet
EndIf

If (SF4->F4_LFICM <> "N" .OR. SF4->F4_LFIPI <> "N") .And. Alltrim(cEspecie) $ U_GCT5("P_ESPECIE-SERVICOS")
	MsgStop("Quando se escritura ICMS ou IPI, a especie da NF n�o pode "+U_GCT5("P_ESPECIE-SERVICOS"),"#MT100LOK Esp�cie Documento")
	lRet := .F.
Endif

Return lRet

&& -------------------------------------------------------------

Static Function Vld10

If !lRet
	Return lRet
EndIf

If cTipo == "I"
	If SF4->F4_PISCRED=="1" .Or. AllTrim(SF4->F4_CSTPIS) $ "01/02/03/04/05/06" .Or. AllTrim(SF4->F4_CSTCOF) $ "01/02/03/04/05/06" 
		MsgStop("Para NF de Complemento de ICMS n�o deve ser usado TES com tributa��o de PIS/COFINS.","#MT100LOK Compl.ICMS x Sit.Trib. PIS/COFINS")
		lRet := .F.
	EndIf
Endif

Return lRet

&& -------------------------------------------------------------

Static Function Vld11

If !lRet
	Return lRet
EndIf

If SubStr(SF4->F4_CF,2,3) $ U_GCT5("P_CFOP-ENERGIA") .And. !Alltrim(CESPECIE) $ U_GCT5("P_ESPECIE-ENERGIA")
	MsgStop("Para energia a esp�cie de documento deve ser uma destas "+U_GCT5("P_ESPECIE-ENERGIA"),"#MT100LOK Esp�cie Documento")
	lRet := .F.
Endif

Return lRet


&& -------------------------------------------------------------

Static Function Vld12

If !lRet
	Return lRet
EndIf

If SubStr(SF4->F4_CF,2,3) $ U_GCT5("P_CFOP-SERV-COMUN") .And. !Alltrim(CESPECIE) $ U_GCT5("P_ESPECIE-SERV-COMUN")
	MsgStop("Para servi�o de comunica��o a esp�cie de documento deve ser uma destas "+U_GCT5("P_ESPECIE-SERV-COMUN"),"#MT100LOK Esp�cie Documento")
	lRet := .F.
Endif

Return lRet

&& -------------------------------------------------------------

Static Function Vld13

If !lRet
	Return lRet
EndIf

If Alltrim(cEspecie) $ U_GCT5("P_ESPECIE-SERVICOS") .And. SF4->F4_LFISS=="N"
	MsgStop("Para NF cuja esp�cie � servi�o o TES deve ser configurado para escriturar o Livro de ISS." , "#MT100LOK Livro ISS")
	lRet := .F.
Endif

Return lRet

&& -------------------------------------------------------------

Static Function Vld14

If !lRet
	Return lRet
EndIf

If Alltrim(cEspecie) $ U_GCT5("P_ESPECIE-SERVICOS") .And. Empty(SB1->B1_CODISS)
	MsgStop("Para NF de servi�o o c�digo do servi�o ISS deve ser informado no cadastro de produtos.","#MT100LOK Cod.ISS cadastro Produto" )
	lRet := .F.
Endif

&& -------------------------------------------------------------

Static Function Vld15

If !lRet
	Return lRet
EndIf

If  Len(Alltrim(CNFiscal)) < 9 .And. cFormul<>"S"
	MsgStop("N�mero de Documento menor que 9 d�gitos. Preencha usando zeros � esquerda at� atingir o tamanho de 9 d�gitos.","#MT100LOK n. NF")
	lRet := .F.
Endif

Return lRet

&& -------------------------------------------------------------

Static Function Vld16

If !lRet
	Return lRet
EndIf

If SF4->F4_LFICM <> "N" .And. Len(Alltrim(cClasFis)) < 3
	MsgStop("Classifica��o fiscal -> " + cClasFis +" inv�lida, verifique os cadastro TES (campo Sit.Trib.ICM) e PRODUTO (campo Origem).","#MT100LOK Class.Fiscal")
	lRet := .F.
Endif

Return lRet

&& -------------------------------------------------------------

Static Function Vld17

If !lRet
	Return lRet
EndIf

If Empty(cEspecie)
	MsgStop("O campo esp�cie do Documento deve ser preenchido." , "#MT100LOK Esp�cie em Branco")
	lRet := .F.
EndIf

Return lRet

&& -------------------------------------------------------------

Static Function Vld18

If !lRet
	Return lRet
EndIf

If SF4->F4_LFIPI <> "N" .And. Empty(SF4->F4_CTIPI)
	MsgStop("C�digo de Tributa��o IPI em branco, preencha o campo no cadastro de TES.","#MT100LOK Cod.Trib.IPI")
	lRet := .F.
Endif

Return lRet

&& -------------------------------------------------------------

Static Function Vld19

If !lRet
	Return lRet
EndIf

If SF4->F4_PISCOF <> "4" .And. ( Empty(SF4->F4_CSTPIS) .Or. Empty(SF4->F4_CSTCOF) )
	MsgStop("Cod.Sit.Trib. PIS ou Cod.Sit.Trib. COFINS em branco, preencha o campo no cadastro de TES.","#MT100LOK Cod.Trib.PIS/COFINS")
	lRet := .F.
Endif

Return lRet

&& -------------------------------------------------------------

Static Function Vld20

If !lRet
	Return lRet
EndIf

If SF4->F4_LFISS <> "N" .And. Empty(SF4->F4_CSTISS)
	MsgStop("Situacao Trib. do ISS em branco, preencha o campo no cadastro de TES.","#MT100LOK Sit.Trib.ISS")
	lRet := .F.
Endif

Return lRet

&& -------------------------------------------------------------

Static Function Vld21

If !lRet
	Return lRet
EndIf

// Se for servi�o de beneficiamento e o TES atualizar estoque, ver se a OP foi informada.
If U_GCT5("P_OBR-OP-RET-INDUSTR") .And. U_GCT5("LCFOP-RET-IND2") .And. SF4->F4_ESTOQUE=="S" .And. Empty(cNumOP)
	MsgStop("Para retorno de produto usado no beneficiamento e TES que atualiza estoques, informar o n�mero da OP." ,"#MT100LOK Ret.Benef x OP")
	lRet := .f.
Endif

Return lRet