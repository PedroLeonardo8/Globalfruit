User Function RETFCLI()

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �RETFCLI   �Autor  � Jo�o Barbosa    � Data �  03/01/2018    ���
�������������������������������������������������������������������������͹��
���Desc.     �Exeblock que retorna Nome Fantasia do Cliente no Browse do  ���
���          �Pedido de Vendas (C5_ZZFCLI)                                ���
�������������������������������������������������������������������������͹��
���Uso       �B2FINANCE - Protheus 12                                     ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

_cRet := " "
           
if SC5->C5_TIPO$"B|D"
_cRet := RETFIELD("SA2",1,XFILIAL("SA2")+C5_CLIENTE+C5_LOJACLI,"SA2->A2_NREDUZ")

else
_cRet := RETFIELD("SA1",1,XFILIAL("SA1")+C5_CLIENTE+C5_LOJACLI,"SA1->A1_NREDUZ")

endif 

Return(_cRet)