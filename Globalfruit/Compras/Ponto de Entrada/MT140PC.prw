#Include 'Protheus.ch'
//Rafael França - 19/08/2021 - Exceção para documento de entrada sem o pedido de compras

User Function MT140PC()

Local lRet := PARAMIXB[1]

//Validações do Usuário
    If (Upper(FunName()) == "MONNF001" .OR. Upper(FunName()) == "MATA116" .OR. Upper(FunName()) == "MATA140") .OR. SD1->D1_TES = "" //Quando for pre nota ou nota fiscal de conhecimento de frete sistema não ira validar o pedido de compras
        lRet := .F. // Retorno False para não validar o parâmetro MV_PCNFE
    EndIf


Return lRet
