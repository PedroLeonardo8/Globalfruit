User Function RETNBCO()

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �RETNBCO   �Autor  � Jo�o Barbosa    � Data �  20/09/2018    ���
�������������������������������������������������������������������������͹��
���Desc.     �Exeblock que retorna Nome do Banco no Browse do Contas a    ���
���          �Receber (C5_ZZNBCO)                                         ���
�������������������������������������������������������������������������͹��
���Uso       �B2FINANCE - Protheus 12                                     ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

_cRet := " "
              
_cRet := RETFIELD("SA6",1,XFILIAL("SA6")+E1_PORTADO+E1_AGEDEP+E1_CONTA,"SA6->A6_NOME")

Return(_cRet)