#Include 'Protheus.ch'
//Rafael Fran�a - 19/08/2021 - Exce��o para documento de entrada sem o pedido de compras

User Function MT140PC()

Local lRet := PARAMIXB[1]

//Valida��es do Usu�rio
    If (Upper(FunName()) == "MONNF001" .OR. Upper(FunName()) == "MATA116" .OR. Upper(FunName()) == "MATA140") .OR. SD1->D1_TES = "" //Quando for pre nota ou nota fiscal de conhecimento de frete sistema n�o ira validar o pedido de compras
        lRet := .F. // Retorno False para n�o validar o par�metro MV_PCNFE
    EndIf


Return lRet
