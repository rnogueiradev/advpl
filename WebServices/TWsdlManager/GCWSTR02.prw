#INCLUDE "TOTVS.CH"
#INCLUDE "TBICONN.CH"
#INCLUDE "RWMAKE.CH"
#INCLUDE "aarray.ch"
#INCLUDE "json.ch"

/*====================================================================================\
|Programa  | GCWSTR02         | Autor | Renato Nogueira            | Data | 09/06/2016|
|=====================================================================================|
|Descrição | Fonte utilizado obter o tracking da TES                                  |
|          |                                                                          |
|          |                                                                          |
|=====================================================================================|
|Sintaxe   | 			                                                                 |
|=====================================================================================|
|Uso       | Especifico GrandCru                                                      |
|=====================================================================================|
|........................................Histórico....................................|
\====================================================================================*/

User Function GCWSTR02(_cEmp,_cFil)

Local x
Local cQuery1		:= ""
Local cAlias1	 	:= GetNextAlias()
Local xRet
Local aOps 		:= {}, aSimple := {}
Local nRegs   	:= 0
Local _n			:= 0
Local cAviso		:= ""
Local _aEmail		:= {}
Local cError  := ""
Local oLastError := ErrorBlock({|e| cError := e:Description + e:ErrorStack})
Local aHeadOut        	:= {}
Local cHeadRet     		:= ""
Local cUrl              := "http://app.shipfy.com/api/tracking/Create"
Local nTimeOut 			:= 60
Local cJson 			:= ""
Local cRet	:= ""
Local _z	:= 0
Default _cEmp	:= "02"
Default _cFil	:= "01"

PREPARE ENVIRONMENT EMPRESA '02' FILIAL '01' TABLES 'SE1,SA1,SE2' MODULO 'FAT'

If !LockByName("GCWSTR02",.F.,.F.,.T.)
	ConOut("[GCWSTR02]["+ FWTimeStamp(2) +"] - Já existe uma sessão em processamento.")
	Return()
EndIf

ConOut("[GCWSTR02]["+ FWTimeStamp(2) +"] Inicio do processamento.")

For _nX:=1 To 2
	
	oWsdl := TWsdlManager():New()
	
	If _nX==1
		oWsdl:SetAuthentication("grandcru-prod","*****") //Producao
	Else
		oWsdl:SetAuthentication("vinogc-prod","*****") //Producao
	EndIf
	
	xRet := oWsdl:ParseURL( "http://edi.totalexpress.com.br/webservice24.php?wsdl" )
	If xRet == .F.
		Conout("Erro: " + oWsdl:cError )
		Return
	EndIf
	aOps := oWsdl:ListOperations()
	If Len( aOps ) == 0
		Conout( "Erro: " + oWsdl:cError )
		Return
	EndIf
	xRet := oWsdl:SetOperation( "ObterTracking" )
	If xRet == .F.
		Conout( "Erro: " + oWsdl:cError )
		Return
	EndIf
	
	aSimple := oWsdl:SimpleInput()
	_nDataConsulta	:= aScan(aSimple, {|aVet| aVet[2] == "DataConsulta" .AND. aVet[5] == "ObterTrackingRequest#1" } )
	xRet := oWsdl:SetValue( aSimple[_nDataConsulta][1],  SUBSTR(DTOS(DATE()),1,4)+"-"+SUBSTR(DTOS(DATE()),5,2)+"-"+SUBSTR(DTOS(DATE()),7,2) )
	
	//oWsdl:lRemEmptyTags := .T.
	
	/*
	cMsg	:= '<soapenv:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:ObterTracking">'
	cMsg	+= '<soapenv:Header/>'
	cMsg	+= '<soapenv:Body>'
	cMsg	+= '<urn:ObterTracking soapenv:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">'
	cMsg	+= '<ObterTrackingRequest xsi:type="web:ObterTrackingRequest" xmlns:web="http://edi.totalexpress.com.br/soap/webservice_v24.total">'
	cMsg	+= '<DataConsulta xsi:type="xsd:date"></DataConsulta>'
	cMsg	+= '</ObterTrackingRequest>'
	cMsg	+= '</urn:ObterTracking>'
	cMsg	+= '</soapenv:Body>'
	cMsg	+= '</soapenv:Envelope>'
	
	// Envia a mensagem SOAP ao servidor
	oWsdl:SendSoapMsg(cMsg)
	*/
	
	Conout( oWsdl:GetSoapMsg())
	
	cRet	:= oWsdl:GetSoapMsg()
	If _nX==1
		cUser	:= Encode64("grandcru-prod:r3X8wfqU")
	Else
		cUser	:= Encode64("vinogc-prod:cZ2CVuCa")
	EndIf
	           
	aHeadOut	:= {}
	
	aAdd( aHeadOut , "POST http://edi.totalexpress.com.br/webservice24.php HTTP/1.1" )
	//aAdd( aHeadOut , "Accept-Encoding: gzip,deflate" )
	aAdd( aHeadOut , "Content-Type: text/xml;charset=UTF-8" )
	aAdd( aHeadOut , 'SOAPAction: "urn:ObterTracking#ObterTracking"' )
	//aAdd( aHeadOut , "Content-Length: 608" )
	aAdd( aHeadOut , "Host: edi.totalexpress.com.br" )
	aAdd( aHeadOut , "Connection: Keep-Alive")
	aAdd( aHeadOut , "User-Agent: Apache-HttpClient/4.1.1 (java 1.5)")
	
	If _nX==1
		//Gerado pelo SOAPUI
		aAdd( aHeadOut , "Authorization: Basic Z3JhbmRjcnUtcHJvZDpyM1g4d2ZxVQ==") //GrandCru
	Else
		//Gerado pelo SOAPUI
		aAdd( aHeadOut , "Authorization: Basic dmlub2djLXByb2Q6Y1oyQ1Z1Q2E=") //VinoClub
	EndIf
	
	cUrl	:= "http://edi.totalexpress.com.br/webservice24.php"
	
	sPostRet := HttpPost(cUrl,"",cRet,nTimeOut,aHeadOut,@cHeadRet)
	
	DbSelectArea("SZC")
	SZC->(DbSetOrder(2))
	
	//oWsdl:SendSoapMsg()
	
	//cResp	:=  oWsdl:GetSoapResponse()
	If Type("sPostRet")=="C"
		
		oResp 	:= XmlParser(sPostRet,"_",@cAviso,@cErro)
		
		If Type("oResp:_SOAP_ENV_ENVELOPE:_SOAP_ENV_BODY:_NS1_OBTERTRACKINGRESPONSE:_OBTERTRACKINGRESPONSE:_ARRAYLOTERETORNO:_ITEM")=="A"
			
			_aItens	:= oResp:_SOAP_ENV_ENVELOPE:_SOAP_ENV_BODY:_NS1_OBTERTRACKINGRESPONSE:_OBTERTRACKINGRESPONSE:_ARRAYLOTERETORNO:_ITEM
			
			For _nY:=1 To Len(_aItens)
				
				If Type("_aItens[_nY]:_ARRAYENCOMENDARETORNO:_ITEM")=="A"
					
					_aItens2	:= _aItens[_nY]:_ARRAYENCOMENDARETORNO:_ITEM
					
					For _nW:=1 To Len(_aItens2)
						
						If Type("_aItens2[_nW]:_ARRAYSTATUSTOTAL:_ITEM")=="A"
							
							For _nZ:=1 To Len(_aItens2[_nW]:_ARRAYSTATUSTOTAL:_ITEM)
								
								Begin Transaction
								
								SZC->(DbGoTop())
								If !SZC->(DbSeek(xFilial("SZC")+PADL(_aItens2[_nW]:_NOTAFISCAL:TEXT,9,"0")+PADR(_aItens2[_nW]:_SERIENOTAFISCAL:TEXT,3)+;
									PADR(_aItens2[_nW]:_PEDIDO:TEXT,6)+PADR(_aItens2[_nW]:_ARRAYSTATUSTOTAL:_ITEM[_nZ]:_CODSTATUS:TEXT,4)+;
									PADR(_aItens2[_nW]:_ARRAYSTATUSTOTAL:_ITEM[_nZ]:_DATASTATUS:TEXT,20)))
									SZC->(RecLock("SZC",.T.))
									SZC->ZC_FILIAL	:= xFilial("SZC")
									SZC->ZC_NOTA		:= PADL(_aItens2[_nW]:_NOTAFISCAL:TEXT,9,"0")
									SZC->ZC_SERIE		:= _aItens2[_nW]:_SERIENOTAFISCAL:TEXT
									SZC->ZC_PEDIDO	:= _aItens2[_nW]:_PEDIDO:TEXT
									SZC->ZC_AWB		:= _aItens2[_nW]:_AWB:TEXT
									SZC->ZC_STATUS	:= _aItens2[_nW]:_ARRAYSTATUSTOTAL:_ITEM[_nZ]:_CODSTATUS:TEXT
									SZC->ZC_DTORIG	:= _aItens2[_nW]:_ARRAYSTATUSTOTAL:_ITEM[_nZ]:_DATASTATUS:TEXT
									SZC->ZC_DATA		:= STOD(SubStr(StrTran(_aItens2[_nW]:_ARRAYSTATUSTOTAL:_ITEM[_nZ]:_DATASTATUS:TEXT,"-",""),1,8))
									SZC->ZC_HORA		:= SubStr(StrTran(_aItens2[_nW]:_ARRAYSTATUSTOTAL:_ITEM[_nZ]:_DATASTATUS:TEXT,"-",""),10,8)
									SZC->ZC_DESCRI	:= _aItens2[_nW]:_ARRAYSTATUSTOTAL:_ITEM[_nZ]:_DESCSTATUS:TEXT
									SZC->ZC_TRANSP	:= "TR0129"
									SZC->(MsUnLock())
									
									DbSelectArea("SC5")
									SC5->(DbSetOrder(1))
									SC5->(DbGoTop())
									If SC5->(DbSeek(SZC->(ZC_FILIAL+ZC_PEDIDO)))
										SC5->(RecLock("SC5",.F.))
										If !Empty(GETSTATUS(SZC->ZC_STATUS))
											SC5->C5_XSTATEX	:= GETSTATUS(SZC->ZC_STATUS)
										EndIf
										SC5->(MsUnLock())
									EndIf
									
									GRVSZD()
									
								EndIf
								
								End Transaction
								
							Next
							
						ElseIf Type("_aItens2[_nW]:_ARRAYSTATUSTOTAL:_ITEM")=="O"
							
							Begin Transaction
							//Conout(_z++)
							//If _z==402
							//msgalert("")
							//EndIf
							SZC->(DbGoTop())
							If !SZC->(DbSeek(xFilial("SZC")+PADL(_aItens2[_nW]:_NOTAFISCAL:TEXT,9,"0")+PADR(_aItens2[_nW]:_SERIENOTAFISCAL:TEXT,3)+;
								PADR(_aItens2[_nW]:_PEDIDO:TEXT,6)+PADR(_aItens2[_nW]:_ARRAYSTATUSTOTAL:_ITEM:_CODSTATUS:TEXT,4)+;
								PADR(_aItens2[_nW]:_ARRAYSTATUSTOTAL:_ITEM:_DATASTATUS:TEXT,20)))
								SZC->(RecLock("SZC",.T.))
								SZC->ZC_FILIAL	:= xFilial("SZC")
								SZC->ZC_NOTA		:= PADL(_aItens2[_nW]:_NOTAFISCAL:TEXT,9,"0")
								SZC->ZC_SERIE		:= _aItens2[_nW]:_SERIENOTAFISCAL:TEXT
								SZC->ZC_PEDIDO	:= _aItens2[_nW]:_PEDIDO:TEXT
								SZC->ZC_AWB		:= _aItens2[_nW]:_AWB:TEXT
								SZC->ZC_STATUS	:= _aItens2[_nW]:_ARRAYSTATUSTOTAL:_ITEM:_CODSTATUS:TEXT
								SZC->ZC_DTORIG	:= _aItens2[_nW]:_ARRAYSTATUSTOTAL:_ITEM:_DATASTATUS:TEXT
								SZC->ZC_DATA		:= STOD(SubStr(StrTran(_aItens2[_nW]:_ARRAYSTATUSTOTAL:_ITEM:_DATASTATUS:TEXT,"-",""),1,8))
								SZC->ZC_HORA		:= SubStr(StrTran(_aItens2[_nW]:_ARRAYSTATUSTOTAL:_ITEM:_DATASTATUS:TEXT,"-",""),10,8)
								SZC->ZC_DESCRI	:= _aItens2[_nW]:_ARRAYSTATUSTOTAL:_ITEM:_DESCSTATUS:TEXT
								SZC->ZC_TRANSP	:= "TR0129"
								SZC->(MsUnLock())
								
								DbSelectArea("SC5")
								SC5->(DbSetOrder(1))
								SC5->(DbGoTop())
								If SC5->(DbSeek(SZC->(ZC_FILIAL+ZC_PEDIDO)))
									SC5->(RecLock("SC5",.F.))
									If !Empty(GETSTATUS(SZC->ZC_STATUS))
										SC5->C5_XSTATEX	:= GETSTATUS(SZC->ZC_STATUS)
									EndIf
									SC5->(MsUnLock())
								EndIf
								
								GRVSZD()
								
							EndIf
							
							End Transaction
							
						EndIf
						
					Next
					
				EndIf
				
			Next
			
		EndIf
	EndIf
	
Next

UnLockByName("GCWSTR02",.F.,.F.,.T.)

ConOut("[GCWSTR02]["+ FWTimeStamp(2) +"] Fim do processamento.")

Reset Environment

Return()

Static Function GRVSZD()

DbSelectArea("SC5")
SC5->(DbSetOrder(1))
SC5->(DbGoTop())

DbSelectArea("SZD")
SZD->(DbSetOrder(2))
SZD->(DbGoTop())

Do Case
	Case AllTrim(SZC->ZC_STATUS)=="103" //Recebido pela transportadora
		
		If SC5->(DbSeek(SZC->(ZC_FILIAL+ZC_PEDIDO)))
			If !SZD->(DbSeek(SC5->(C5_FILIAL+C5_NUM+C5_CLIENTE+C5_LOJACLI)+"3"))
				
				SZD->(RecLock("SZD",.T.))
				SZD->ZD_FILIAL	:= SC5->C5_FILIAL
				SZD->ZD_PEDIDO	:= SC5->C5_NUM
				SZD->ZD_CLIENTE	:= SC5->C5_CLIENTE
				SZD->ZD_LOJA		:= SC5->C5_LOJACLI
				SZD->ZD_NOTA		:= SC5->C5_NOTA
				SZD->ZD_SERIE		:= SC5->C5_SERIE
				SZD->ZD_EMAIL		:= Posicione("SA1",1,xFilial("SA1")+SC5->(C5_CLIENTE+C5_LOJACLI),"A1_EMAIL")
				SZD->ZD_STATUS	:= "3"
				SZD->ZD_DTGRAV	:= Date()
				SZD->ZD_HRGRAV	:= Time()
				SZD->(MsUnLock())
				
			EndIf
		EndIf
		
	Case AllTrim(SZC->ZC_STATUS)=="1" //Entrega realizada
		
		If SC5->(DbSeek(SZC->(ZC_FILIAL+ZC_PEDIDO)))
			If !SZD->(DbSeek(SC5->(C5_FILIAL+C5_NUM+C5_CLIENTE+C5_LOJACLI)+"4"))
				
				SZD->(RecLock("SZD",.T.))
				SZD->ZD_FILIAL	:= SC5->C5_FILIAL
				SZD->ZD_PEDIDO	:= SC5->C5_NUM
				SZD->ZD_CLIENTE	:= SC5->C5_CLIENTE
				SZD->ZD_LOJA		:= SC5->C5_LOJACLI
				SZD->ZD_NOTA		:= SC5->C5_NOTA
				SZD->ZD_SERIE		:= SC5->C5_SERIE
				SZD->ZD_EMAIL		:= Posicione("SA1",1,xFilial("SA1")+SC5->(C5_CLIENTE+C5_LOJACLI),"A1_EMAIL")
				SZD->ZD_STATUS	:= "4"
				SZD->ZD_DTGRAV	:= Date()
				SZD->ZD_HRGRAV	:= Time()
				SZD->(MsUnLock())
				
			EndIf
		EndIf
		
EndCase

Return()

Static Function GETSTATUS(_cStatus)

Local _cRetorno	:= ""
Default _cStatus:= ""

Do Case
	Case AllTrim(_cStatus)=="0"
		_cRetorno	:= "1"
	Case AllTrim(_cStatus)=="69"
		_cRetorno	:= "2"
	Case AllTrim(_cStatus)=="102"
		_cRetorno	:= "3"
	Case AllTrim(_cStatus)=="104"
		_cRetorno	:= "4"
	Case AllTrim(_cStatus)=="1"
		_cRetorno	:= "7"
EndCase

Return(_cRetorno)
