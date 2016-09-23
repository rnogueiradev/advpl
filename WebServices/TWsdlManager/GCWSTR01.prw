#INCLUDE "TOTVS.CH"
#INCLUDE "TBICONN.CH"
#INCLUDE "RWMAKE.CH"
#INCLUDE "aarray.ch"
#INCLUDE "json.ch"

/*====================================================================================\
|Programa  | GCWSTR01         | Autor | Renato Nogueira            | Data | 09/06/2016|
|=====================================================================================|
|Descrição | Fonte utilizado para enviar as notas pendentes para a Total Express      |
|          |                                                                          |
|          |                                                                          |
|=====================================================================================|
|Sintaxe   | 			                                                                 |
|=====================================================================================|
|Uso       | Especifico GrandCru                                                      |
|=====================================================================================|
|........................................Histórico....................................|
\====================================================================================*/

User Function GCWSTR01(_cEmp,_cFil)

Local cQuery1		:= ""
Local cAlias1	 	:= GetNextAlias()
Local oWsdl
Local xRet
Local aOps 		:= {}, aSimple := {}
Local nRegs   	:= 0
Local _n			:= 0
Local cAviso		:= ""
Local cErro		:= ""
Local _cErro		:= ""
Local _aEmail		:= {}
Local _cCfop	:= ""
Local _lOk		:= .F.
Default _cEmp	:= "02"
Default _cFil	:= "01"

//RpcSetType(3)
//RpcSetEnv(_cEmp,_cFil,,,"FAT")

PREPARE ENVIRONMENT EMPRESA '02' FILIAL '01' TABLES 'SE1,SA1,SE2' MODULO 'FAT'

If !LockByName("GCWSTR01",.F.,.F.,.T.)
	ConOut("[GCWSTR01]["+ FWTimeStamp(2) +"] - Já existe uma sessão em processamento.")
	Return()
EndIf

ConOut("[GCWSTR01]["+ FWTimeStamp(2) +"] Inicio do processamento.")

//1=Magento;2=Oracle;3=Confr.XTECH;4=Confr.BRADESCO;5=Televendas;6=Vinoclub;7=Confr.Multiplus;8=Email;9=Geosales;A=Outros

For _nX:=1 To 2
	
	cQuery1  := " SELECT F2_FILIAL, F2_DOC, F2_SERIE, F2_CHVNFE, D2_PEDIDO, F2_CLIENTE, F2_LOJA, F2_VOLUME1, REGISTRO, PEDIDO, A1_NOME, "
	cQuery1  += " A1_CGC, A1_END, A1_BAIRRO, A1_MUN, A1_EST, A1_EMAIL, A1_CEP, A1_DDD, A1_TEL, EMISSAO, F2_VALBRUT, F2_VALMERC, F2_CHVNFE, C5_XORIGWS "
	cQuery1  += " FROM (
	cQuery1  += " SELECT DISTINCT F2_FILIAL, F2_DOC, F2_SERIE, F2_CHVNFE, D2_PEDIDO, F2_CLIENTE, F2_LOJA, F2_VOLUME1, F2.R_E_C_N_O_ REGISTRO,"
	cQuery1  += " C5_NUM PEDIDO, "
	cQuery1  += " A1_NOME COLLATE sql_latin1_general_cp1251_ci_as A1_NOME, A1_CGC, "
	cQuery1  += " A1_END COLLATE sql_latin1_general_cp1251_ci_as A1_END, "
	cQuery1  += " A1_BAIRRO COLLATE sql_latin1_general_cp1251_ci_as A1_BAIRRO, "
	cQuery1  += " A1_MUN COLLATE sql_latin1_general_cp1251_ci_as A1_MUN, "
	cQuery1  += " A1_EST COLLATE sql_latin1_general_cp1251_ci_as A1_EST, "
	cQuery1  += " A1_EMAIL COLLATE sql_latin1_general_cp1251_ci_as A1_EMAIL, A1_CEP, A1_DDD, A1_TEL, SUBSTRING(F2_EMISSAO,1,4)+'-'+SUBSTRING(F2_EMISSAO,5,2)+'-'+SUBSTRING(F2_EMISSAO,7,2) EMISSAO, F2_VALBRUT, F2_VALMERC, C5_XORIGWS "
	cQuery1  += " FROM "+RetSqlName("SC5")+" (NOLOCK) C5 "
	cQuery1  += " LEFT JOIN "+RetSqlName("SD2")+" (NOLOCK) D2 "
	cQuery1  += " ON D2_FILIAL=C5_FILIAL AND D2_PEDIDO=C5_NUM AND D2_CLIENTE=C5_CLIENTE AND D2_LOJA=C5_LOJACLI "
	cQuery1  += " LEFT JOIN "+RetSqlName("SF2")+" (NOLOCK) F2 "
	cQuery1  += " ON F2_FILIAL=D2_FILIAL AND D2_DOC=F2_DOC AND D2_SERIE=F2_SERIE AND D2_CLIENTE=D2_CLIENTE AND D2_LOJA=F2_LOJA "
	cQuery1  += " LEFT JOIN "+RetSqlName("SA1")+" (NOLOCK) A1 "
	cQuery1  += " ON A1_COD=F2_CLIENTE AND A1_LOJA=F2_LOJA "
	cQuery1  += " WHERE C5.D_E_L_E_T_=' ' AND D2.D_E_L_E_T_=' ' AND F2.D_E_L_E_T_=' ' AND A1.D_E_L_E_T_=' ' AND F2_CHVNFE<>' '  "
	cQuery1  += " AND C5_TRANSP='TR0129' AND F2_XTOTEXP=' ' "
	//cQuery1  += " AND F2_EMISSAO<cast(convert(varchar(8),GETDATE(),112) as int) " //NÃO ENVIAR NOTAS DO MESMO DIA
	
	If _nX==1 //Todas menos vinoclub
		cQuery1  += " AND C5_XORIGWS<>' ' AND C5_XORIGWS<>'6' "
	Else //Vinoclub
		cQuery1  += " AND C5_XORIGWS IN ('6')
	EndIf
	
	cQuery1  += " ) XXX "
	
	If !Empty(Select(cAlias1))
		DbSelectArea(cAlias1)
		(cAlias1)->(dbCloseArea())
	Endif
	
	dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery1),cAlias1,.T.,.T.)
	
	Count To nRegs
	
	DbSelectArea(cAlias1)
	(cAlias1)->(DbGoTop())
	
	While (cAlias1)->(!Eof())
		
		oWsdl := TWsdlManager():New()
		
		If _nX==1
			//oWsdl:SetAuthentication("grandcru-qa","bz79g25X") //Teste
			oWsdl:SetAuthentication("grandcru-prod","r3X8wfqU") //Producao
		Else
			//oWsdl:SetAuthentication("vinogc-qa","trGFc3qb") //Teste
			oWsdl:SetAuthentication("vinogc-prod","cZ2CVuCa") //Producao
		EndIf
		
		xRet := oWsdl:ParseURL( SuperGetMv("GC_WSDLTEX",.F.,"http://edi.totalexpress.com.br/webservice24.php?wsdl") )
		If xRet == .F.
			Conout("Erro: " + oWsdl:cError )
			Return
		EndIf
		aOps := oWsdl:ListOperations()
		If Len( aOps ) == 0
			Conout( "Erro: " + oWsdl:cError )
			Return
		EndIf
		xRet := oWsdl:SetOperation( "RegistraColeta" )
		If xRet == .F.
			Conout( "Erro: " + oWsdl:cError )
			Return
		EndIf
		
		oWsdl:lRemEmptyTags := .T.
		
		aComplex := oWsdl:NextComplex()
		While Type("aComplex")=="A"
			varinfo( "aComplex", aComplex )
			
			If (aComplex[2]=="item") .And. (aComplex[5]=="RegistraColetaRequest#1.Encomendas#1")
				nOccurs := 1//nRegs
			ElseIf (aComplex[2]=="DocFiscalNFe")
				nOccurs := 1
			ElseIf (aComplex[2]=="item")
				nOccurs := 1
			Else
				nOccurs := 0
			Endif
			
			xRet := oWsdl:SetComplexOccurs( aComplex[1], nOccurs )
			If xRet == .F.
				Conout( "Erro ao definir elemento " + aComplex[2] + ", ID " + cValToChar( aComplex[1] ) + ", com " + cValToChar( nOccurs ) + " ocorrencias" )
				Return
			EndIf
			
			aComplex := oWsdl:NextComplex()
		EndDo
		
		If xRet == .F.
			Return
		EndIf
		
		aSimple := oWsdl:SimpleInput()
		varinfo( "aSimple",aSimple)
		
		_n	:= 1
		
		_nCodRem	:= aScan(aSimple, {|aVet| aVet[2] == "CodRemessa" .AND. aVet[5] == "RegistraColetaRequest#1" } )
		xRet := oWsdl:SetValue( aSimple[_nCodRem][1], 	"1" )
		
		//Dados da encomenda
		_nTpServ	:= aScan(aSimple, {|aVet| aVet[2] == "TipoServico" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1") } )
		_nTpEnt	:= aScan(aSimple, {|aVet| aVet[2] == "TipoEntrega" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1") } )
		_nVolumes	:= aScan(aSimple, {|aVet| aVet[2] == "Volumes" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1") } )
		_nCondFre	:= aScan(aSimple, {|aVet| aVet[2] == "CondFrete" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1") } )
		_nPedido	:= aScan(aSimple, {|aVet| aVet[2] == "Pedido" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1") } )
		_nNaturez	:= aScan(aSimple, {|aVet| aVet[2] == "Natureza" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1") } )
		_nIIcms	:= aScan(aSimple, {|aVet| aVet[2] == "IsencaoIcms" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1") } )
		
		//Dados do destinatário
		_nNome		:= aScan(aSimple, {|aVet| aVet[2] == "DestNome" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1") } )
		_nCgc 		:= aScan(aSimple, {|aVet| aVet[2] == "DestCpfCnpj" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1") } )
		_nEnd 		:= aScan(aSimple, {|aVet| aVet[2] == "DestEnd" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1") } )
		_nNum 		:= aScan(aSimple, {|aVet| aVet[2] == "DestEndNum" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1") } )
		_nBairro	:= aScan(aSimple, {|aVet| aVet[2] == "DestBairro" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1") } )
		_nCidade	:= aScan(aSimple, {|aVet| aVet[2] == "DestCidade" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1") } )
		_nEstado	:= aScan(aSimple, {|aVet| aVet[2] == "DestEstado" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1") } )
		_nCep		:= aScan(aSimple, {|aVet| aVet[2] == "DestCep" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1") } )
		_nEmail	:= aScan(aSimple, {|aVet| aVet[2] == "DestEmail" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1") } )
		_nDDD		:= aScan(aSimple, {|aVet| aVet[2] == "DestDdd" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1") } )
		_nTel		:= aScan(aSimple, {|aVet| aVet[2] == "DestTelefone1" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1") } )
		
		//Dados da nota fiscal
		_nNF		:= aScan(aSimple, {|aVet| aVet[2] == "NfeNumero" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1")+".DocFiscalNFe#1.*#1" } )
		_nSerie	:= aScan(aSimple, {|aVet| aVet[2] == "NfeSerie" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1")+".DocFiscalNFe#1.*#1" } )
		_nData		:= aScan(aSimple, {|aVet| aVet[2] == "NfeData" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1")+".DocFiscalNFe#1.*#1" } )
		_nTotal	:= aScan(aSimple, {|aVet| aVet[2] == "NfeValTotal" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1")+".DocFiscalNFe#1.*#1" } )
		_nLiquido	:= aScan(aSimple, {|aVet| aVet[2] == "NfeValProd" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1")+".DocFiscalNFe#1.*#1" } )
		_nCfops	:= aScan(aSimple, {|aVet| aVet[2] == "NfeCfop" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1")+".DocFiscalNFe#1.*#1" } )
		_nChave	:= aScan(aSimple, {|aVet| aVet[2] == "NfeChave" .AND. aVet[5] == "RegistraColetaRequest#1.Encomendas#1."+IIf(_n>1,"item#"+CVALTOCHAR(_n),"*#1")+".DocFiscalNFe#1.*#1" } )
		
		//Faz o set dos valores
		
		_cCfop	:= ""
		DbSelectArea("SD2")
		SD2->(DbSetOrder(3)) //D2_FILIAL+D2_DOC+D2_SERIE+D2_CLIENTE+D2_LOJA+D2_COD+D2_ITEM
		SD2->(DbGoTop())
		If SD2->(DbSeek((cAlias1)->(F2_FILIAL+F2_DOC+F2_SERIE+F2_CLIENTE+F2_LOJA)))
			_cCfop	:= SD2->D2_CF
		EndIf
		
		xRet := oWsdl:SetValue( aSimple[_nTpServ][1], 	"1" )
		xRet := oWsdl:SetValue( aSimple[_nTpEnt][1], 		"0" )
		xRet := oWsdl:SetValue( aSimple[_nVolumes][1], 	AllTrim(NOACENTO(CVALTOCHAR(IIf((cAlias1)->F2_VOLUME1>0,(cAlias1)->F2_VOLUME1,1)))) )
		xRet := oWsdl:SetValue( aSimple[_nCondFre][1], 	"CIF" )
		xRet := oWsdl:SetValue( aSimple[_nPedido][1], 	AllTrim(NOACENTO((cAlias1)->PEDIDO)) )
		xRet := oWsdl:SetValue( aSimple[_nNaturez][1], 	AllTrim(NOACENTO(_cCfop)) )
		xRet := oWsdl:SetValue( aSimple[_nIIcms][1], 		"0" )
		
		xRet := oWsdl:SetValue( aSimple[_nNome][1], 		AllTrim(NOACENTO((cAlias1)->A1_NOME)) )
		xRet := oWsdl:SetValue( aSimple[_nCgc][1], 		AllTrim(NOACENTO((cAlias1)->A1_CGC)) )
		xRet := oWsdl:SetValue( aSimple[_nEnd][1], 		AllTrim(NOACENTO(StrTran((cAlias1)->A1_END,","))) )
		xRet := oWsdl:SetValue( aSimple[_nNum][1], 		Iif(Empty(AllTrim(U_GCCTOVAL(NOACENTO((cAlias1)->A1_END)))),"SN",AllTrim(U_GCCTOVAL(NOACENTO((cAlias1)->A1_END)))) )
		xRet := oWsdl:SetValue( aSimple[_nBairro][1], 	AllTrim(NOACENTO((cAlias1)->A1_BAIRRO)) )
		xRet := oWsdl:SetValue( aSimple[_nCidade][1], 	AllTrim(NOACENTO((cAlias1)->A1_MUN) ) )
		xRet := oWsdl:SetValue( aSimple[_nEstado][1], 	AllTrim(NOACENTO((cAlias1)->A1_EST)) )
		xRet := oWsdl:SetValue( aSimple[_nCep][1], 		AllTrim(NOACENTO((cAlias1)->A1_CEP)) )
		xRet := oWsdl:SetValue( aSimple[_nEmail][1], 	AllTrim(NOACENTO((cAlias1)->A1_EMAIL)) )
		xRet := oWsdl:SetValue( aSimple[_nDDD][1], 		AllTrim(NOACENTO((cAlias1)->A1_DDD)) )
		xRet := oWsdl:SetValue( aSimple[_nTel][1], 		AllTrim(NOACENTO((cAlias1)->A1_TEL)) )
		
		xRet := oWsdl:SetValue( aSimple[_nNF][1], 		AllTrim(NOACENTO((cAlias1)->F2_DOC)) )
		xRet := oWsdl:SetValue( aSimple[_nSerie][1], 		AllTrim(NOACENTO((cAlias1)->F2_SERIE)) )
		xRet := oWsdl:SetValue( aSimple[_nData][1], 		(cAlias1)->EMISSAO )
		xRet := oWsdl:SetValue( aSimple[_nTotal][1], 		AllTrim(CVALTOCHAR((cAlias1)->F2_VALBRUT)) )
		xRet := oWsdl:SetValue( aSimple[_nLiquido][1], 	AllTrim(NOACENTO(CVALTOCHAR((cAlias1)->F2_VALMERC))) )
		xRet := oWsdl:SetValue( aSimple[_nCfops][1], 		AllTrim(NOACENTO(_cCfop)) )
		xRet := oWsdl:SetValue( aSimple[_nChave][1], 		AllTrim(NOACENTO((cAlias1)->F2_CHVNFE)) )
		
		Conout( oWsdl:GetSoapMsg() )
		
		// Envia a mensagem SOAP ao servidor
		xRet := oWsdl:SendSoapMsg()
		/*
		If xRet == .F.
		Conout( "Erro: " + oWsdl:cError )
		Return
		EndIf
		*/
		// Pega a mensagem de resposta
		cResp	:=  oWsdl:GetSoapResponse()
		
		_lOk		:= .F.
		
		DbSelectArea("SF2")
		SF2->(DbGoTop())
		SF2->(DbGoTo((cAlias1)->REGISTRO))
		If SF2->(!Eof())
			If Type("cResp")=="C"
				oResp 	:= XmlParser(cResp,"_",@cAviso,@cErro)
				If Type("oResp:_SOAP_ENV_ENVELOPE:_SOAP_ENV_BODY:_NS1_REGISTRACOLETARESPONSE:_REGISTRACOLETARESPONSE:_CODIGOPROC:TEXT")=="C"
					_cProc	:= AllTrim(oResp:_SOAP_ENV_ENVELOPE:_SOAP_ENV_BODY:_NS1_REGISTRACOLETARESPONSE:_REGISTRACOLETARESPONSE:_CODIGOPROC:TEXT)
					Do Case
						Case _cProc=="1" //Processado
							If Type("oResp:_SOAP_ENV_ENVELOPE:_SOAP_ENV_BODY:_NS1_REGISTRACOLETARESPONSE:_REGISTRACOLETARESPONSE:_ITENSPROCESSADOS:TEXT")=="C" //Processado
								_cQtdProc	:= oResp:_SOAP_ENV_ENVELOPE:_SOAP_ENV_BODY:_NS1_REGISTRACOLETARESPONSE:_REGISTRACOLETARESPONSE:_ITENSPROCESSADOS:TEXT
								_nQtdProc	:= Val(_cQtdProc)
								If _nQtdProc==1 //Processado
									SF2->(RecLock("SF2",.F.))
									SF2->F2_XTOTEXP	:= "S"
									SF2->F2_XDTTEX	:= Date()
									SF2->F2_XHRTEX	:= Time()
									//SF2->F2_XROMTEX	:= ""
									SF2->(MsUnLock())
									_lOk	:= .T.
								EndIf
							EndIf
						Case _cProc=="5"
							If Type("oResp:_SOAP_ENV_ENVELOPE:_SOAP_ENV_BODY:_NS1_REGISTRACOLETARESPONSE:_REGISTRACOLETARESPONSE:_ERROSINDIVIDUAIS:_ITEM")=="A"
								_aErros	:= oResp:_SOAP_ENV_ENVELOPE:_SOAP_ENV_BODY:_NS1_REGISTRACOLETARESPONSE:_REGISTRACOLETARESPONSE:_ERROSINDIVIDUAIS:_ITEM
								For _nY:=1 To Len(_aErros)
									_cErro	+= _aErros[_nY]:_DESCRICAOERRO:TEXT+" / "
								Next
							ElseIf Type("oResp:_SOAP_ENV_ENVELOPE:_SOAP_ENV_BODY:_NS1_REGISTRACOLETARESPONSE:_REGISTRACOLETARESPONSE:_ERROSINDIVIDUAIS:_ITEM:_DESCRICAOERRO:TEXT")=="C"
									_cErro	+= oResp:_SOAP_ENV_ENVELOPE:_SOAP_ENV_BODY:_NS1_REGISTRACOLETARESPONSE:_REGISTRACOLETARESPONSE:_ERROSINDIVIDUAIS:_ITEM:_DESCRICAOERRO:TEXT 
							EndIf
					EndCase
				EndIf
			EndIf
			
			If _lOk
				AADD(_aEmail,{GETORIG((cAlias1)->C5_XORIGWS),SF2->F2_FILIAL,SF2->F2_DOC,SF2->F2_SERIE,SF2->F2_CLIENTE,SF2->F2_LOJA,SF2->F2_VALBRUT,"OK",""})
			Else
				AADD(_aEmail,{GETORIG((cAlias1)->C5_XORIGWS),SF2->F2_FILIAL,SF2->F2_DOC,SF2->F2_SERIE,SF2->F2_CLIENTE,SF2->F2_LOJA,SF2->F2_VALBRUT,"ERRO",_cErro})
			EndIf
			
		EndIf
		
		(cAlias1)->(DbSkip())
		
	EndDo
	
Next

If Len(_aEmail)>0
	SENDMAIL(_aEmail)
EndIf

UnLockByName("GCWSTR01",.F.,.F.,.T.)

ConOut("[GCWSTR01]["+ FWTimeStamp(2) +"] Fim do processamento.")

Reset Environment

Return()

Static Function SENDMAIL(_aEmail)

Local cMsg	    := ""
Local cAttach   := ''
Local _aAttach  := {}
Local _cCaminho := ''
Local _cCo	:= ""

Local cAssunto		:= "[WFPROTHEUS] - Processamento de notas total express"
Local cAccount		:= GetMV("MV_EMCONTA")
Local cPassword 	:= GetMV("MV_EMSENHA")
Local cServer   	:= GetMV("MV_RELSERV")
Local cEmailde  	:= GetMV("MV_RELFROM")
Local cTo 			:= "renato.nogueira@rvgsolucoes.com.br;reinaldo.dantas@rvgsolucoes.com.br"//SuperGetMv("GC_EMLTEX",.F.,"renato.nogueira@rvgsolucoes.com.br")

Local oButton1
Local oGet1
Local oSay1
Local oDlg
Local lOk := .F.

cMsg := ""
cMsg += '<html>'
cMsg += '<head>'
cMsg += '<title>' +SM0->M0_NOME+"/"+SM0->M0_FILIAL+'</title>'
cMsg += '</head>'
cMsg += '<body>'
cMsg += '<HR Width=85% Size=3 Align=Centered Color=Red> <P>'
cMsg += '<Table Border=1 Align=Center BorderColor=#000000 CELLPADDING=4 CELLSPACING=0 Width=60%>'
cMsg += '<Caption> <FONT COLOR=#000000 FACE= "ARIAL" SIZE=5>GRAND CRU</FONT> </Caption>'
cMsg += '<TR><B><TD>ORIGEM</TD><TD>FILIAL</TD><TD>NOTA</TD><TD>SERIE</TD><TD>CLIENTE</TD><TD>LOJA</TD><TD>VALOR</TD><TD>STATUS</TD><TD>MENSAGEM</TD></B></TR>'
For _nLin := 1 to Len(_aEmail)
	If _aEmail[_nLin,8]=="OK"
		cMsg += '<TR BgColor=#9AFF9A>'
	ElseIf _aEmail[_nLin,8]=="ERRO"
		cMsg += '<TR BgColor=#FFA07A>'
	EndIf
	cMsg += '<TD><Font Color=#000000 Size="2" Face="Arial">' + _aEmail[_nLin,1] + ' </Font></TD>'
	cMsg += '<TD><Font Color=#000000 Size="2" Face="Arial">' + _aEmail[_nLin,2] + ' </Font></TD>'
	cMsg += '<TD><Font Color=#000000 Size="2" Face="Arial">' + _aEmail[_nLin,3] + ' </Font></TD>'
	cMsg += '<TD><Font Color=#000000 Size="2" Face="Arial">' + _aEmail[_nLin,4] + ' </Font></TD>'
	cMsg += '<TD><Font Color=#000000 Size="2" Face="Arial">' + _aEmail[_nLin,5] + ' </Font></TD>'
	cMsg += '<TD><Font Color=#000000 Size="2" Face="Arial">' + _aEmail[_nLin,6] + ' </Font></TD>'
	cMsg += '<TD><Font Color=#000000 Size="2" Face="Arial">' + CVALTOCHAR(_aEmail[_nLin,7]) + ' </Font></TD>'
	cMsg += '<TD><Font Color=#000000 Size="2" Face="Arial">' + _aEmail[_nLin,8] + ' </Font></TD>'
	cMsg += '<TD><Font Color=#000000 Size="2" Face="Arial">' + _aEmail[_nLin,9] + ' </Font></TD>'
	cMsg += '</TR>'
Next
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Definicao do rodape do email                                                ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
cMsg += '</Table>'
cMsg += '<P>'
cMsg += '<Table align="center">'
cMsg += '<tr>'
cMsg += '<td colspan="10" align="center"><font color="red" size="3">Mensagem enviada automaticamente pelo sistema PROTHEUS - <font color="red" size="1">(GCWSTR01.PRW)</td>'
cMsg += '</tr>'
cMsg += '</Table>'
cMsg += '<HR Width=85% Size=3 Align=Centered Color=Red> <P>'
//cMsg += '<B><Font Color=#000000 Size="2" Face="Arial"> Atenciosamente </Font></B><BR>'
//cMsg += '<B><Font Color=#000000 Size="2" Face="Arial">' + SM0->M0_NOMECOM + '</Font></B><BR>'
//cMsg += '<Img Src="C:/AP5/SIGAADV/LGRL01.BMP">'
cMsg += '</body>'
cMsg += '</html>'

STMAILTES(cTo, "", cAssunto, cMsg,_aAttach,_cCaminho,_cCo)
//TkSendMail( cAccount,cPassword,cServer,cEmailde,cTo,cAssunto,cMsg,"")

Return()

/*====================================================================================\
|Programa  | STMAILTES        | Autor | GIOVANI.ZAGO             | Data | 27/03/2013  |
|=====================================================================================|
|Descrição | STMAILTES                                                                |
|          |                                                                          |
|          |                                                                          |
|=====================================================================================|
|Sintaxe   | STMAILTES                                                                |
|=====================================================================================|
|Uso       | Especifico STECK                                                         |
|=====================================================================================|
|........................................Histórico....................................|
\====================================================================================*/
*------------------------------------------------------------------------*
Static Function STMAILTES(cPara, cCopia, cAssunto, cMsg,_aAttach,_cCaminho,_cCo)
*------------------------------------------------------------------------*
Local oMail , oMessage , nErro
Local lRet := .T.
Local i    := 0
Local cSMTPServer	:= GetMV("GC_RELSERV",,"smtp.servername.com.br")
Local cSMTPUser		:= GetMV("GC_RELAUSR",,"minhaconta@servername.com.br")
Local cSMTPPass		:= GetMV("GC_RELAPSW",,"minhasenha")
Local cMailFrom		:= GetMV("GC_RELFROM",,"minhaconta@servername.com.br")
Local nPort	   		:= GetMV("GC_GCPPORT",,25)
Local lUseAuth		:= GetMV("GC_RELAUTH",,.T.)

Default cMsg     := "<hr>Envio de e-mail via Protheus<hr>"
Default cCopia   := ""
Default cPara    := 'giovani.zago@steck.com.br'
Default cAssunto := 'Envio de Teste'
Default _aAttach := {}
Default _cCo     := ""
conout('Conectando com SMTP ['+cSMTPServer+'] ')

oMail := TMailManager():New()
conout('Inicializando SMTP')

If GetMv("GC_RELTLS",,.T.)
	oMail:SetUseTLS(.t.)
EndIf

oMail:Init( '', cSMTPServer , cSMTPUser, cSMTPPass, 0, nPort  )

conout('Setando Time-Out')
oMail:SetSmtpTimeOut( 30 )

conout('Conectando com servidor...')
nErro := oMail:SmtpConnect()

conout('Status de Retorno = '+str(nErro,6))

If lUseAuth
	Conout("Autenticando Usuario ["+cSMTPUser+"] senha ["+cSMTPPass+"]")
	nErro := oMail:SmtpAuth(cSMTPUser ,cSMTPPass)
	conout('Status de Retorno = '+str(nErro,6))
	If nErro <> 0
		// Recupera erro ...
		cMAilError := oMail:GetErrorString(nErro)
		DEFAULT cMailError := '***UNKNOW***'
		Conout("Erro de Autenticacao "+str(nErro,4)+' ('+cMAilError+')')
		lRet := .F.
	Endif
Endif

if nErro <> 0
	
	// Recupera erro
	cMAilError := oMail:GetErrorString(nErro)
	DEFAULT cMailError := '***UNKNOW***'
	conout(cMAilError)
	
	Conout("Erro de Conexão SMTP "+str(nErro,4))
	
	conout('Desconectando do SMTP')
	oMail:SMTPDisconnect()
	
	lRet := .F.
	
Endif

If lRet
	conout('Compondo mensagem em memória')
	oMessage := TMailMessage():New()
	oMessage:Clear()
	oMessage:cFrom	:= cMailFrom
	oMessage:cTo	:= cPara
	oMessage:cBcc   := _cCo+";"+GetMv("ST_MAILBAC",,' ')
	If !Empty(cCopia)
		oMessage:cCc	:= cCopia
	Endif
	oMessage:cSubject	:= cAssunto
	oMessage:cBody		:= cMsg
	
	
	For i:= 1 To Len(_aAttach)
		
		//Adiciona um attach
		If oMessage:AttachFile( _cCaminho+_aAttach[i] ) < 0
			Conout( "Erro ao atachar o arquivo" )
			//Return .F.
		Else
			//adiciona uma tag informando que é um attach e o nome do arq
			oMessage:AddAtthTag( 'Content-Disposition: attachment; filename='+_aAttach[i])
		EndIf
	Next i
	
	
	conout('Enviando Mensagem para ['+cPara+'] ')
	nErro := oMessage:Send( oMail )
	
	if nErro <> 0
		xError := oMail:GetErrorString(nErro)
		Conout("Erro de Envio SMTP "+str(nErro,4)+" ("+xError+")")
		lRet := .F.
	Endif
	
	conout('Desconectando do SMTP')
	oMail:SMTPDisconnect()
Endif

Return(lRet)

Static Function GETORIG(_cOrigem)

Local _cDesc	:= ""

_cOrigem	:= AllTrim(_cOrigem)

//1=Magento;2=Oracle;3=Confr.XTECH;4=Confr.BRADESCO;5=Televendas;6=Vinoclub;7=Confr.Multiplus;8=Email;9=Geosales;A=Outros;B=Citi

Do case
	Case _cOrigem=="1"
		_cDesc	:= "Magento"
	Case _cOrigem=="2"
		_cDesc	:= "Oracle"
	Case _cOrigem=="3"
		_cDesc	:= "XTECH"
	Case _cOrigem=="4"
		_cDesc	:= "Bradesco"
	Case _cOrigem=="5"
		_cDesc	:= "Televendas"
	Case _cOrigem=="6"
		_cDesc	:= "VinoClub"
	Case _cOrigem=="7"
		_cDesc	:= "Multiplus"
	Case _cOrigem=="8"
		_cDesc	:= "Email"
	Case _cOrigem=="9"
		_cDesc	:= "GeoSales"
	Case _cOrigem=="A"
		_cDesc	:= "Outros"
	Case _cOrigem=="B"
		_cDesc	:= "Citibank"
EndCase

Return(_cDesc)

Static function TiraGraf(_sOrig)

Local _sRet := _sOrig

_sRet = strtran (_sRet, "á", "a")
_sRet = strtran (_sRet, "é", "e")
_sRet = strtran (_sRet, "í", "i")
_sRet = strtran (_sRet, "ó", "o")
_sRet = strtran (_sRet, "ú", "u")
_SRET = STRTRAN (_SRET, "Á", "A")
_SRET = STRTRAN (_SRET, "É", "E")
_SRET = STRTRAN (_SRET, "Í", "I")
_SRET = STRTRAN (_SRET, "Ó", "O")
_SRET = STRTRAN (_SRET, "Ú", "U")
_sRet = strtran (_sRet, "ã", "a")
_sRet = strtran (_sRet, "õ", "o")
_SRET = STRTRAN (_SRET, "Ã", "A")
_SRET = STRTRAN (_SRET, "Â ", "")
_SRET = STRTRAN (_SRET, "Õ", "O")
_sRet = strtran (_sRet, "â", "a")
_sRet = strtran (_sRet, "ê", "e")
_sRet = strtran (_sRet, "î", "i")
_sRet = strtran (_sRet, "ô", "o")
_sRet = strtran (_sRet, "û", "u")
_SRET = STRTRAN (_SRET, "Â", "A")
_SRET = STRTRAN (_SRET, "Ê", "E")
_SRET = STRTRAN (_SRET, "Î", "I")
_SRET = STRTRAN (_SRET, "Ô", "O")
_SRET = STRTRAN (_SRET, "Û", "U")
_sRet = strtran (_sRet, "ç", "c")
_sRet = strtran (_sRet, "Ç", "C")
_sRet = strtran (_sRet, "à", "a")
_sRet = strtran (_sRet, "À", "A")
_sRet = strtran (_sRet, "º", ".")
_sRet = strtran (_sRet, "ª", ".")
_sRet = strtran (_sRet, "&", "")

return _sRet

User Function GCCTOVAL(_cString)

Local	_cStr2		:= ""
Default _cString	:= ""

_cString	:= AllTrim(_cString)

For _n:=1 To Len(_cString)
	If SubStr(_cString,_n,1) $ "0123456789"
		_cStr2	+= SubStr(_cString,_n,1)
	EndIf
Next

Return(_cStr2)
