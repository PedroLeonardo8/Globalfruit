User Function RETBSE5()

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �RETBSE5   �Autor  � Jo�o Barbosa    � Data �  20/09/2018    ���
�������������������������������������������������������������������������͹��
���Desc.     �Exeblock que retorna Numero do Banco no Browse do           ���
���          �Contas a Pagar (E2_ZZBSE5)                                  ���
�������������������������������������������������������������������������͹��
���Uso       �B2FINANCE - Protheus 12                                     ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

_cRet := " "
              
_cRet := RETFIELD("SE5",23,XFILIAL("SE5")+"P"+E2_FORNECE+E2_LOJA+E2_PREFIXO+E2_NUM,"SE5->E5_BANCO")

Return(_cRet)