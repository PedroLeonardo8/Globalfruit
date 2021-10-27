#INCLUDE "PROTHEUS.CH"

User Function FVDEI()

Local lVldOn	:= SUPERGETMV("ZZ_KITVLD", .f., .f.)

&& Verifica se executa validações.
If !lVldOn
	Return(.t.)
Endif

&& *** Importante: As tabelas devem estar posicionadas para as validações abaixo ***

lRet := Vld01() && Retorna conta contábil e valida a mesma.
lRet := Vld02() && Valida se o TES esta de acordo com o tipo de documento aceito.
lRet := Vld03() && Valida tipo de produto x TES que atualiza estoque
lRet := Vld04() && Valida entrada de NF de serviço de beneficiamento x TES x OP
lRet := Vld05() && Valida Espécie x Formulário Próprio
lRet := Vld06() && Valida Quantidade de dias da emissão da NF
lRet := Vld07() && Valida Informações para DIRF
lRet := Vld08() && Valida Especie x CFOP de Frete
lRet := Vld09() && Valida Especie x Escrituração ICMS / IPI
lRet := Vld10() && Valida NF Complemento ICMS x Tributação de PIS/Cofins
lRet := Vld11() && Valida Especie x CFOP de Energia
lRet := Vld12() && Valida Especie x CFOP de Serv.Comunicação
lRet := Vld13() && Valida Especie Nota Serviço x Escrituração no Livro
//lRet := Vld14() && Valida Especie Nota Serviço x Cód. ISS no produto
lRet := Vld15() && Valida o tamanho do número da NF
lRet := Vld16() && Valida se a classificação fiscal esta correta
lRet := Vld17() && Valida se a especie foi preenchida
//lRet := Vld18() && Valida a situação tributária do IPI
//lRet := Vld19() && Valida a situação tributária do PIS/COFINS
//lRet := Vld20() && Valida a situação tributária do ISS
lRet := Vld21() && Valida retorno de beneficiamento utilizado no processo de beneficiamento

Return lRet

&& -------------------------------------------------------------

Static Function Vld01

If !lRet
	Return lRet
EndIf

&& Não valida caso TES não contabiliza
If SF4->F4_ZZCTB == "2"
	Return .t.
EndIf

&& Não valida rateio
If lRateio
	Return .t.
EndIf

&& Obtém a conta contábil no produto, já que não foi informada no documento de entrada. Prioriza B1_CONTA para estoque e depois os campos customizados.
If Empty(cConta)
	If SF4->F4_ESTOQUE=="S"
		cConta := SB1->B1_CONTA
	Else
		If !Empty(cRN1CC)
			cConta := U_GCT5("P_SB1-CTA-"+cRN1CC)
		EndIf
	EndIf
Endif

&& Valida se foi informada a conta contábil.
If Empty(cConta) .And. (SF4->F4_ESTOQUE=="S" .Or. SF4->F4_DUPLIC=="S")
	lRet := .f.
	MsgStop("Conta contábil não informada. Se for uma operação de despesa/custos informe o centro de custo para retornar a conta." + ;
	" Se informou o c.custo e não veio conta, então o produto esta com conta em branco.","#MT100LOK Inconsistênia Cta. Contábil")
EndIf

&& Valida obrigatoriedade de entidades
If lRet
	lRet := CtbObrig(cConta,cCusto,cItemCta,cClVl,.t.)
EndIf

&& Valida regras de amarração
If lRet
	lRet := CtbAmarra(cConta,cCusto,cItemCta,cClVl,.t.)
EndIf

&& Retorna conta para o título
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
	MsgStop("O tipo de movimento da NF [ "+cTipo+" ] não esta de acordo com os tipos de movimentos aceitos pelo CFOP [ "+cTpAceito + ;
	" ]. Verifique a movimentação e escolha o tipo de movimento correto ou TES correto. Para suporte procurar pela área Fiscal.","#MT100LOK Tipo Mov.")
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
	MsgStop("De acordo com as regras do SPED Fiscal este tipo de produto [ "+SB1->B1_TIPO+" ] não pode movimentar estoque."+;
	" Tipos permitidos [ "+cTipoEst+"] ","#MT100LOK Tipo produto x Estoques")
	lRet := .f.
EndIf

Return lRet

&& -------------------------------------------------------------

Static Function Vld04

If !lRet
	Return lRet
EndIf

// Se for serviço de beneficiamento e o TES atualizar estoque, ver se a OP foi informada.
If SubStr(SF4->F4_CF,2,3) $ U_GCT5("P_CFOP-SERV-BENEF") .And. SF4->F4_ESTOQUE=="S" .And. Empty(cNumOP)
	MsgStop("Para serviço de beneficiamento e TES que atualiza estoques, informar o número da OP." ,"#MT100LOK Serv.Benef x OP")
	lRet := .f.
Endif

Return lRet

&& -------------------------------------------------------------

Static Function Vld05

If !lRet
	Return lRet
EndIf

If alltrim(cFormul) == "S" .And. alltrim(cEspecie) <> "SPED"
	MsgStop("O campo 'Espécie' deve ser preenchido com 'SPED' quando usar formulário próprio.","#MT100LOK Espécie x Formulário Próprio")
	lRet := .F.
EndIf

Return lRet

&& -------------------------------------------------------------

Static Function Vld06

If !lRet
	Return lRet
EndIf

If DDEMISSAO < DDATABASE-15
	lRet := MsgYesNo("A Data de Emissão da Nota é inferior há 15 dias da DATABASE. Esta correto?", "#MT100LOK Idade da NF")
EndIf

Return lRet

&& -------------------------------------------------------------

Static Function Vld07

If !lRet
	Return lRet
EndIf

If ( MaFisRet(,"NF_VALIRR") + MaFisRet(,"NF_VALCOF") + MaFisRet(,"NF_VALPIS") + MaFisRet(,"NF_VALCSL") ) > 0
	If cDirf=="2"
		MsgStop("Para operações com IR, PIS, COFINS, CSLL o campo 'Gera Dirf' deve ser preenchido com 'Sim'.","#MT100LOK Gera Dirf")
		lRet := .f.
	EndIf
	If Empty(cCodRet) .And. lRet
		MsgStop("Para operações com IR, PIS, COFINS, CSLL o campo 'Cd.Retencao' deve ser preenchido com o codigo do DARF.","#MT100LOK código do DARF")
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
	MsgStop("Para frete a espécie de documento deve ser uma destas "+U_GCT5("P_ESPECIE-FRETE"),"#MT100LOK Espécie Documento")
	lRet := .F.
Endif

Return lRet

&& -------------------------------------------------------------

Static Function Vld09

If !lRet
	Return lRet
EndIf

If (SF4->F4_LFICM <> "N" .OR. SF4->F4_LFIPI <> "N") .And. Alltrim(cEspecie) $ U_GCT5("P_ESPECIE-SERVICOS")
	MsgStop("Quando se escritura ICMS ou IPI, a especie da NF não pode "+U_GCT5("P_ESPECIE-SERVICOS"),"#MT100LOK Espécie Documento")
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
		MsgStop("Para NF de Complemento de ICMS não deve ser usado TES com tributação de PIS/COFINS.","#MT100LOK Compl.ICMS x Sit.Trib. PIS/COFINS")
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
	MsgStop("Para energia a espécie de documento deve ser uma destas "+U_GCT5("P_ESPECIE-ENERGIA"),"#MT100LOK Espécie Documento")
	lRet := .F.
Endif

Return lRet


&& -------------------------------------------------------------

Static Function Vld12

If !lRet
	Return lRet
EndIf

If SubStr(SF4->F4_CF,2,3) $ U_GCT5("P_CFOP-SERV-COMUN") .And. !Alltrim(CESPECIE) $ U_GCT5("P_ESPECIE-SERV-COMUN")
	MsgStop("Para serviço de comunicação a espécie de documento deve ser uma destas "+U_GCT5("P_ESPECIE-SERV-COMUN"),"#MT100LOK Espécie Documento")
	lRet := .F.
Endif

Return lRet

&& -------------------------------------------------------------

Static Function Vld13

If !lRet
	Return lRet
EndIf

If Alltrim(cEspecie) $ U_GCT5("P_ESPECIE-SERVICOS") .And. SF4->F4_LFISS=="N"
	MsgStop("Para NF cuja espécie é serviço o TES deve ser configurado para escriturar o Livro de ISS." , "#MT100LOK Livro ISS")
	lRet := .F.
Endif

Return lRet

&& -------------------------------------------------------------

Static Function Vld14

If !lRet
	Return lRet
EndIf

If Alltrim(cEspecie) $ U_GCT5("P_ESPECIE-SERVICOS") .And. Empty(SB1->B1_CODISS)
	MsgStop("Para NF de serviço o código do serviço ISS deve ser informado no cadastro de produtos.","#MT100LOK Cod.ISS cadastro Produto" )
	lRet := .F.
Endif

&& -------------------------------------------------------------

Static Function Vld15

If !lRet
	Return lRet
EndIf

If  Len(Alltrim(CNFiscal)) < 9 .And. cFormul<>"S"
	MsgStop("Número de Documento menor que 9 dígitos. Preencha usando zeros à esquerda até atingir o tamanho de 9 dígitos.","#MT100LOK n. NF")
	lRet := .F.
Endif

Return lRet

&& -------------------------------------------------------------

Static Function Vld16

If !lRet
	Return lRet
EndIf

If SF4->F4_LFICM <> "N" .And. Len(Alltrim(cClasFis)) < 3
	MsgStop("Classificação fiscal -> " + cClasFis +" inválida, verifique os cadastro TES (campo Sit.Trib.ICM) e PRODUTO (campo Origem).","#MT100LOK Class.Fiscal")
	lRet := .F.
Endif

Return lRet

&& -------------------------------------------------------------

Static Function Vld17

If !lRet
	Return lRet
EndIf

If Empty(cEspecie)
	MsgStop("O campo espécie do Documento deve ser preenchido." , "#MT100LOK Espécie em Branco")
	lRet := .F.
EndIf

Return lRet

&& -------------------------------------------------------------

Static Function Vld18

If !lRet
	Return lRet
EndIf

If SF4->F4_LFIPI <> "N" .And. Empty(SF4->F4_CTIPI)
	MsgStop("Código de Tributação IPI em branco, preencha o campo no cadastro de TES.","#MT100LOK Cod.Trib.IPI")
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

// Se for serviço de beneficiamento e o TES atualizar estoque, ver se a OP foi informada.
If U_GCT5("P_OBR-OP-RET-INDUSTR") .And. U_GCT5("LCFOP-RET-IND2") .And. SF4->F4_ESTOQUE=="S" .And. Empty(cNumOP)
	MsgStop("Para retorno de produto usado no beneficiamento e TES que atualiza estoques, informar o número da OP." ,"#MT100LOK Ret.Benef x OP")
	lRet := .f.
Endif

Return lRet