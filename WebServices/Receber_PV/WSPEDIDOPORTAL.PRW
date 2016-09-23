#INCLUDE "APWEBSRV.CH"
#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOTVS.CH"

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³PEDPORTAL	ºAutor  ³Renato Nogueira     º Data ³  03/06/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³Webservice de integração dos pedidos de vendas			   º±±
±±º          ³	    							 	 		        			º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºParametro ³ Nenhum                                                     º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±ºRetorno   ³ Nenhum                                                     º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

WSSTRUCT Itens
	
	WSDATA cItem			AS STRING
	WSDATA cProduto		AS STRING
	WSDATA nQuant			AS FLOAT
	WSDATA nPreco 		AS FLOAT
	WSDATA cItemPed		AS STRING
	
ENDWSSTRUCT

WSSTRUCT Estrutura
	
	WSDATA cCnpjSte	AS STRING
	WSDATA cCnpjEnt	AS STRING
	WSDATA cCnpjCli	AS STRING
	WSDATA cCondPag	AS STRING
	WSDATA cTipoFre	AS STRING
	WSDATA cTipoEnt	AS STRING
	WSDATA cTipoFat	AS STRING
	WSDATA cPedCli	AS STRING
	WSDATA cObserv	AS STRING
	WSDATA dDtEnt		AS DATE
	
	WSDATA aProds   AS ARRAY OF Itens
	
ENDWSSTRUCT

WSSERVICE WSPEDIDOPORTAL DESCRIPTION "Serviço para receber pedidos do portal - produção"
	
	WSDATA Pedido    AS Estrutura
	WSDATA lRetorno  AS BOOLEAN
	
	WSMETHOD RetPedido DESCRIPTION "Método que retorna se o pedido foi inserido com sucesso! - produção"
	
ENDWSSERVICE

WSMETHOD RetPedido WSRECEIVE Pedido WSSEND lRetorno WSSERVICE WSPEDIDOPORTAL
	
	Conout("["+DTOC(Date())+" - "+Time()+"] - Gravando pedido via webservice - OC: "+::Pedido:cPedCli)
	
	//Verificar se todos os campos obrigatórios foram preenchidos
	
	If	Empty(::Pedido:cCnpjCli) .Or.;
			Empty(::Pedido:cCondPag) .Or.;
			Empty(::Pedido:cPedCli) .Or.;
			Empty(::Pedido:cTipoFre)
		Conout("Os campos obrigatórios do cabeçalho não foram preenchidos - OC: "+::Pedido:cPedCli)
		::lRetorno	:= .F.
		Return .T.
	EndIf
	
	For i := 1 to len(::Pedido:aProds)
		If 	Empty(::Pedido:aProds[i]:cItemPed) .Or.;
				Empty(::Pedido:aProds[i]:nPreco) .Or.;
				Empty(::Pedido:aProds[i]:nQuant) .Or.;
				Empty(::Pedido:aProds[i]:cProduto)
			Conout("Os campos obrigatórios do item não foram preenchidos - OC: "+::Pedido:cPedCli)
			::lRetorno	:= .F.
			Return .T.
		EndIf
	Next
	
	//Verifica se cliente existe na base
	DbSelectArea("SA1")
	SA1->(DbSetOrder(3)) //A1_FILIAL+A1_CGC
	SA1->(DbGoTop())
	If !SA1->(DbSeek(xFilial("SA1")+::Pedido:cCnpjCli))
		Conout("CNPJ: "+AllTrim(::Pedido:cCnpjCli)+" não encontrado!"+" - OC: "+::Pedido:cPedCli)
		::lRetorno	:= .F.
		Return .T.
	EndIf
	
	//Verifica se OC já existe na base
	DbSelectArea("ZZ6")
	ZZ6->(DbSetOrder(1)) //ZZ6_FILIAL+ZZ6_NUMOC+ZZ6_CLIENT+ZZ6_LOJA
	ZZ6->(DbGoTop())
	If ZZ6->(DbSeek(xFilial("ZZ6")+PADR(AllTrim(::Pedido:cPedCli),15)+SA1->(A1_COD+A1_LOJA)))
		Conout("OC: "+AllTrim(::Pedido:cPedCli)+" já existe na base!")
		::lRetorno	:= .T.
		Return .T.
	EndIf
	
	//Iniciando a importação dos dados para as tabelas
	
	Begin Transaction
		
		//Cabeçalho
		ZZ6->(RecLock("ZZ6",.T.))
		ZZ6->ZZ6_FILIAL	:= xFilial("ZZ6")
		ZZ6->ZZ6_NUMOC	:= AllTrim(::Pedido:cPedCli)
		ZZ6->ZZ6_CLIENT	:= SA1->A1_COD
		ZZ6->ZZ6_LOJA	:= SA1->A1_LOJA
		ZZ6->ZZ6_CNPJ	:= AllTrim(::Pedido:cCnpjCli)
		ZZ6->ZZ6_CONDPG	:= ::Pedido:cCondPag
		ZZ6->ZZ6_OBSERV	:= ::Pedido:cObserv
		ZZ6->ZZ6_TPENT	:= ::Pedido:cTipoEnt
		ZZ6->ZZ6_TPFAT	:= ::Pedido:cTipoFat	//::Pedido:cTipoFat - Tipo de faturamento - sempre 2 - regra definida dia 13/11/2014 com a Tatiane Neves
		ZZ6->ZZ6_TPFRET	:= ::Pedido:cTipoFre 	//::Pedido:cTipoFre - Tipo de entrega - sempre 2 - regra definida dia 13/11/2014 com a Tatiane Neves
		ZZ6->ZZ6_DTENT	:= ::Pedido:dDtEnt
		ZZ6->ZZ6_STATUS	:= "1"
		ZZ6->ZZ6_DATA  	:= Date()
		ZZ6->ZZ6_HORA	:= Time()
		ZZ6->(MsUnLock())
		
		//Itens
		
		For i := 1 to len(::Pedido:aProds)
			ZZ7->(RecLock("ZZ7",.T.))
			ZZ7->ZZ7_FILIAL	:=	xFilial("ZZ7")
			ZZ7->ZZ7_NUMOC 	:= AllTrim(::Pedido:cPedCli)
			ZZ7->ZZ7_CLIENT	:= SA1->A1_COD
			ZZ7->ZZ7_LOJA	:= SA1->A1_LOJA
			ZZ7->ZZ7_PRODUT	:= ::Pedido:aProds[i]:cProduto
			ZZ7->ZZ7_ITEMP 	:= ::Pedido:aProds[i]:cItemPed
			ZZ7->ZZ7_QUANT 	:= ::Pedido:aProds[i]:nQuant
			ZZ7->ZZ7_PRECO 	:= ::Pedido:aProds[i]:nPreco
			ZZ7->(MsUnLock())
		Next
		
	End Transaction
	
	Conout("["+DTOC(Date())+" - "+Time()+"] - Pedido via webservice - OC: "+::Pedido:cPedCli+" gravado ")
	::lRetorno	:= .T.
	
Return .T.