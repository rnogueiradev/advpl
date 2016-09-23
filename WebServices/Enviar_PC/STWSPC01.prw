#INCLUDE "PROTHEUS.CH"
#INCLUDE "APWEBSRV.CH"

/* ===============================================================================
WSDL Location    http://www.unicasolucoes.com.br/ws_uniconnect/ws_uniconnect.asmx?WSDL
Gerado em        09/11/15 09:47:54
Observações      Código-Fonte gerado por ADVPL WSDL Client 1.120703
                 Alterações neste arquivo podem causar funcionamento incorreto
                 e serão perdidas caso o código-fonte seja gerado novamente.
=============================================================================== */

User Function _OMLUKDP ; Return  // "dummy" function - Internal Use 

/* -------------------------------------------------------------------------------
WSDL Service WSws_uniconnect
------------------------------------------------------------------------------- */

WSCLIENT WSws_uniconnect

	WSMETHOD NEW
	WSMETHOD INIT
	WSMETHOD RESET
	WSMETHOD CLONE
	WSMETHOD ConfiguracaoAutomatica
	WSMETHOD RecebeArquivoConfiguracao
	WSMETHOD ValidaLogin
	WSMETHOD testHTTPConnection
	WSMETHOD recebeArquivo
	WSMETHOD enviaArquivo
	WSMETHOD listarArquivos
	WSMETHOD moverArquivo
	WSMETHOD logFTP

	WSDATA   _URL                      AS String
	WSDATA   _HEADOUT                  AS Array of String
	WSDATA   _COOKIES                  AS Array of String
	WSDATA   cstrIdConf                AS string
	WSDATA   oWSConfiguracaoAutomaticaResult AS ws_uniconnect_UsuariosComunic
	WSDATA   oWSoUsuario               AS ws_uniconnect_UsuariosComunic
	WSDATA   cstrArquivoConfiguracao   AS string
	WSDATA   lRecebeArquivoConfiguracaoResult AS boolean
	WSDATA   lValidaLoginResult        AS boolean
	WSDATA   ctestHTTPConnectionResult AS string
	WSDATA   oWSoArquivo               AS ws_uniconnect_Arquivo
	WSDATA   lrecebeArquivoResult      AS boolean
	WSDATA   oWSenviaArquivoResult     AS ws_uniconnect_Arquivo
	WSDATA   oWSlistarArquivosResult   AS ws_uniconnect_ArrayOfArquivo
	WSDATA   lmoverArquivoResult       AS boolean
	WSDATA   oWSoLog                   AS ws_uniconnect_Log_InOut
	WSDATA   llogFTPResult             AS boolean

ENDWSCLIENT

WSMETHOD NEW WSCLIENT WSws_uniconnect
::Init()
If !FindFunction("XMLCHILDEX")
	UserException("O Código-Fonte Client atual requer os executáveis do Protheus Build [7.00.131227A-20150625] ou superior. Atualize o Protheus ou gere o Código-Fonte novamente utilizando o Build atual.")
EndIf
If val(right(GetWSCVer(),8)) < 1.040504
	UserException("O Código-Fonte Client atual requer a versão de Lib para WebServices igual ou superior a ADVPL WSDL Client 1.040504. Atualize o repositório ou gere o Código-Fonte novamente utilizando o repositório atual.")
EndIf
Return Self

WSMETHOD INIT WSCLIENT WSws_uniconnect
	::oWSConfiguracaoAutomaticaResult := ws_uniconnect_USUARIOSCOMUNIC():New()
	::oWSoUsuario        := ws_uniconnect_USUARIOSCOMUNIC():New()
	::oWSoArquivo        := ws_uniconnect_ARQUIVO():New()
	::oWSenviaArquivoResult := ws_uniconnect_ARQUIVO():New()
	::oWSlistarArquivosResult := ws_uniconnect_ARRAYOFARQUIVO():New()
	::oWSoLog            := ws_uniconnect_LOG_INOUT():New()
Return

WSMETHOD RESET WSCLIENT WSws_uniconnect
	::cstrIdConf         := NIL 
	::oWSConfiguracaoAutomaticaResult := NIL 
	::oWSoUsuario        := NIL 
	::cstrArquivoConfiguracao := NIL 
	::lRecebeArquivoConfiguracaoResult := NIL 
	::lValidaLoginResult := NIL 
	::ctestHTTPConnectionResult := NIL 
	::oWSoArquivo        := NIL 
	::lrecebeArquivoResult := NIL 
	::oWSenviaArquivoResult := NIL 
	::oWSlistarArquivosResult := NIL 
	::lmoverArquivoResult := NIL 
	::oWSoLog            := NIL 
	::llogFTPResult      := NIL 
	::Init()
Return

WSMETHOD CLONE WSCLIENT WSws_uniconnect
Local oClone := WSws_uniconnect():New()
	oClone:_URL          := ::_URL 
	oClone:cstrIdConf    := ::cstrIdConf
	oClone:oWSConfiguracaoAutomaticaResult :=  IIF(::oWSConfiguracaoAutomaticaResult = NIL , NIL ,::oWSConfiguracaoAutomaticaResult:Clone() )
	oClone:oWSoUsuario   :=  IIF(::oWSoUsuario = NIL , NIL ,::oWSoUsuario:Clone() )
	oClone:cstrArquivoConfiguracao := ::cstrArquivoConfiguracao
	oClone:lRecebeArquivoConfiguracaoResult := ::lRecebeArquivoConfiguracaoResult
	oClone:lValidaLoginResult := ::lValidaLoginResult
	oClone:ctestHTTPConnectionResult := ::ctestHTTPConnectionResult
	oClone:oWSoArquivo   :=  IIF(::oWSoArquivo = NIL , NIL ,::oWSoArquivo:Clone() )
	oClone:lrecebeArquivoResult := ::lrecebeArquivoResult
	oClone:oWSenviaArquivoResult :=  IIF(::oWSenviaArquivoResult = NIL , NIL ,::oWSenviaArquivoResult:Clone() )
	oClone:oWSlistarArquivosResult :=  IIF(::oWSlistarArquivosResult = NIL , NIL ,::oWSlistarArquivosResult:Clone() )
	oClone:lmoverArquivoResult := ::lmoverArquivoResult
	oClone:oWSoLog       :=  IIF(::oWSoLog = NIL , NIL ,::oWSoLog:Clone() )
	oClone:llogFTPResult := ::llogFTPResult
Return oClone

// WSDL Method ConfiguracaoAutomatica of Service WSws_uniconnect

WSMETHOD ConfiguracaoAutomatica WSSEND cstrIdConf WSRECEIVE oWSConfiguracaoAutomaticaResult WSCLIENT WSws_uniconnect
Local cSoap := "" , oXmlRet

BEGIN WSMETHOD

cSoap += '<ConfiguracaoAutomatica xmlns="http://unicasolucoes.com.br/">'
cSoap += WSSoapValue("strIdConf", ::cstrIdConf, cstrIdConf , "string", .F. , .F., 0 , NIL, .F.) 
cSoap += "</ConfiguracaoAutomatica>"

oXmlRet := SvcSoapCall(	Self,cSoap,; 
	"http://unicasolucoes.com.br/ConfiguracaoAutomatica",; 
	"DOCUMENT","http://unicasolucoes.com.br/",,,; 
	"http://www.unicasolucoes.com.br/ws_uniconnect/ws_uniconnect.asmx")

::Init()
::oWSConfiguracaoAutomaticaResult:SoapRecv( WSAdvValue( oXmlRet,"_CONFIGURACAOAUTOMATICARESPONSE:_CONFIGURACAOAUTOMATICARESULT","UsuariosComunic",NIL,NIL,NIL,NIL,NIL,NIL) )

END WSMETHOD

oXmlRet := NIL
Return .T.

// WSDL Method RecebeArquivoConfiguracao of Service WSws_uniconnect

WSMETHOD RecebeArquivoConfiguracao WSSEND oWSoUsuario,cstrArquivoConfiguracao WSRECEIVE lRecebeArquivoConfiguracaoResult WSCLIENT WSws_uniconnect
Local cSoap := "" , oXmlRet

BEGIN WSMETHOD

cSoap += '<RecebeArquivoConfiguracao xmlns="http://unicasolucoes.com.br/">'
cSoap += WSSoapValue("oUsuario", ::oWSoUsuario, oWSoUsuario , "UsuariosComunic", .F. , .F., 0 , NIL, .F.) 
cSoap += WSSoapValue("strArquivoConfiguracao", ::cstrArquivoConfiguracao, cstrArquivoConfiguracao , "string", .F. , .F., 0 , NIL, .F.) 
cSoap += "</RecebeArquivoConfiguracao>"

oXmlRet := SvcSoapCall(	Self,cSoap,; 
	"http://unicasolucoes.com.br/RecebeArquivoConfiguracao",; 
	"DOCUMENT","http://unicasolucoes.com.br/",,,; 
	"http://www.unicasolucoes.com.br/ws_uniconnect/ws_uniconnect.asmx")

::Init()
::lRecebeArquivoConfiguracaoResult :=  WSAdvValue( oXmlRet,"_RECEBEARQUIVOCONFIGURACAORESPONSE:_RECEBEARQUIVOCONFIGURACAORESULT:TEXT","boolean",NIL,NIL,NIL,NIL,NIL,NIL) 

END WSMETHOD

oXmlRet := NIL
Return .T.

// WSDL Method ValidaLogin of Service WSws_uniconnect

WSMETHOD ValidaLogin WSSEND oWSoUsuario WSRECEIVE lValidaLoginResult WSCLIENT WSws_uniconnect
Local cSoap := "" , oXmlRet

BEGIN WSMETHOD

cSoap += '<ValidaLogin xmlns="http://unicasolucoes.com.br/">'
cSoap += WSSoapValue("oUsuario", ::oWSoUsuario, oWSoUsuario , "UsuariosComunic", .F. , .F., 0 , NIL, .F.) 
cSoap += "</ValidaLogin>"

oXmlRet := SvcSoapCall(	Self,cSoap,; 
	"http://unicasolucoes.com.br/ValidaLogin",; 
	"DOCUMENT","http://unicasolucoes.com.br/",,,; 
	"http://www.unicasolucoes.com.br/ws_uniconnect/ws_uniconnect.asmx")

::Init()
::lValidaLoginResult :=  WSAdvValue( oXmlRet,"_VALIDALOGINRESPONSE:_VALIDALOGINRESULT:TEXT","boolean",NIL,NIL,NIL,NIL,NIL,NIL) 

END WSMETHOD

oXmlRet := NIL
Return .T.

// WSDL Method testHTTPConnection of Service WSws_uniconnect

WSMETHOD testHTTPConnection WSSEND oWSoUsuario WSRECEIVE ctestHTTPConnectionResult WSCLIENT WSws_uniconnect
Local cSoap := "" , oXmlRet

BEGIN WSMETHOD

cSoap += '<testHTTPConnection xmlns="http://unicasolucoes.com.br/">'
cSoap += WSSoapValue("oUsuario", ::oWSoUsuario, oWSoUsuario , "UsuariosComunic", .F. , .F., 0 , NIL, .F.) 
cSoap += "</testHTTPConnection>"

oXmlRet := SvcSoapCall(	Self,cSoap,; 
	"http://unicasolucoes.com.br/testHTTPConnection",; 
	"DOCUMENT","http://unicasolucoes.com.br/",,,; 
	"http://www.unicasolucoes.com.br/ws_uniconnect/ws_uniconnect.asmx")

::Init()
::ctestHTTPConnectionResult :=  WSAdvValue( oXmlRet,"_TESTHTTPCONNECTIONRESPONSE:_TESTHTTPCONNECTIONRESULT:TEXT","string",NIL,NIL,NIL,NIL,NIL,NIL) 

END WSMETHOD

oXmlRet := NIL
Return .T.

// WSDL Method recebeArquivo of Service WSws_uniconnect

WSMETHOD recebeArquivo WSSEND oWSoUsuario,oWSoArquivo WSRECEIVE lrecebeArquivoResult WSCLIENT WSws_uniconnect
Local cSoap := "" , oXmlRet

BEGIN WSMETHOD

cSoap += '<recebeArquivo xmlns="http://unicasolucoes.com.br/">'
cSoap += WSSoapValue("oUsuario", ::oWSoUsuario, oWSoUsuario , "UsuariosComunic", .F. , .F., 0 , NIL, .F.) 
cSoap += WSSoapValue("oArquivo", ::oWSoArquivo, oWSoArquivo , "Arquivo", .F. , .F., 0 , NIL, .F.) 
cSoap += "</recebeArquivo>"

oXmlRet := SvcSoapCall(	Self,cSoap,; 
	"http://unicasolucoes.com.br/recebeArquivo",; 
	"DOCUMENT","http://unicasolucoes.com.br/",,,; 
	"http://www.unicasolucoes.com.br/ws_uniconnect/ws_uniconnect.asmx")

::Init()
::lrecebeArquivoResult :=  WSAdvValue( oXmlRet,"_RECEBEARQUIVORESPONSE:_RECEBEARQUIVORESULT:TEXT","boolean",NIL,NIL,NIL,NIL,NIL,NIL) 

END WSMETHOD

oXmlRet := NIL
Return .T.

// WSDL Method enviaArquivo of Service WSws_uniconnect

WSMETHOD enviaArquivo WSSEND oWSoUsuario,oWSoArquivo WSRECEIVE oWSenviaArquivoResult WSCLIENT WSws_uniconnect
Local cSoap := "" , oXmlRet

BEGIN WSMETHOD

cSoap += '<enviaArquivo xmlns="http://unicasolucoes.com.br/">'
cSoap += WSSoapValue("oUsuario", ::oWSoUsuario, oWSoUsuario , "UsuariosComunic", .F. , .F., 0 , NIL, .F.) 
cSoap += WSSoapValue("oArquivo", ::oWSoArquivo, oWSoArquivo , "Arquivo", .F. , .F., 0 , NIL, .F.) 
cSoap += "</enviaArquivo>"

oXmlRet := SvcSoapCall(	Self,cSoap,; 
	"http://unicasolucoes.com.br/enviaArquivo",; 
	"DOCUMENT","http://unicasolucoes.com.br/",,,; 
	"http://www.unicasolucoes.com.br/ws_uniconnect/ws_uniconnect.asmx")

::Init()
::oWSenviaArquivoResult:SoapRecv( WSAdvValue( oXmlRet,"_ENVIAARQUIVORESPONSE:_ENVIAARQUIVORESULT","Arquivo",NIL,NIL,NIL,NIL,NIL,NIL) )

END WSMETHOD

oXmlRet := NIL
Return .T.

// WSDL Method listarArquivos of Service WSws_uniconnect

WSMETHOD listarArquivos WSSEND oWSoUsuario WSRECEIVE oWSlistarArquivosResult WSCLIENT WSws_uniconnect
Local cSoap := "" , oXmlRet

BEGIN WSMETHOD

cSoap += '<listarArquivos xmlns="http://unicasolucoes.com.br/">'
cSoap += WSSoapValue("oUsuario", ::oWSoUsuario, oWSoUsuario , "UsuariosComunic", .F. , .F., 0 , NIL, .F.) 
cSoap += "</listarArquivos>"

oXmlRet := SvcSoapCall(	Self,cSoap,; 
	"http://unicasolucoes.com.br/listarArquivos",; 
	"DOCUMENT","http://unicasolucoes.com.br/",,,; 
	"http://www.unicasolucoes.com.br/ws_uniconnect/ws_uniconnect.asmx")

::Init()
::oWSlistarArquivosResult:SoapRecv( WSAdvValue( oXmlRet,"_LISTARARQUIVOSRESPONSE:_LISTARARQUIVOSRESULT","ArrayOfArquivo",NIL,NIL,NIL,NIL,NIL,NIL) )

END WSMETHOD

oXmlRet := NIL
Return .T.

// WSDL Method moverArquivo of Service WSws_uniconnect

WSMETHOD moverArquivo WSSEND oWSoUsuario,oWSoArquivo WSRECEIVE lmoverArquivoResult WSCLIENT WSws_uniconnect
Local cSoap := "" , oXmlRet

BEGIN WSMETHOD

cSoap += '<moverArquivo xmlns="http://unicasolucoes.com.br/">'
cSoap += WSSoapValue("oUsuario", ::oWSoUsuario, oWSoUsuario , "UsuariosComunic", .F. , .F., 0 , NIL, .F.) 
cSoap += WSSoapValue("oArquivo", ::oWSoArquivo, oWSoArquivo , "Arquivo", .F. , .F., 0 , NIL, .F.) 
cSoap += "</moverArquivo>"

oXmlRet := SvcSoapCall(	Self,cSoap,; 
	"http://unicasolucoes.com.br/moverArquivo",; 
	"DOCUMENT","http://unicasolucoes.com.br/",,,; 
	"http://www.unicasolucoes.com.br/ws_uniconnect/ws_uniconnect.asmx")

::Init()
::lmoverArquivoResult :=  WSAdvValue( oXmlRet,"_MOVERARQUIVORESPONSE:_MOVERARQUIVORESULT:TEXT","boolean",NIL,NIL,NIL,NIL,NIL,NIL) 

END WSMETHOD

oXmlRet := NIL
Return .T.

// WSDL Method logFTP of Service WSws_uniconnect

WSMETHOD logFTP WSSEND oWSoLog WSRECEIVE llogFTPResult WSCLIENT WSws_uniconnect
Local cSoap := "" , oXmlRet

BEGIN WSMETHOD

cSoap += '<logFTP xmlns="http://unicasolucoes.com.br/">'
cSoap += WSSoapValue("oLog", ::oWSoLog, oWSoLog , "Log_InOut", .F. , .F., 0 , NIL, .F.) 
cSoap += "</logFTP>"

oXmlRet := SvcSoapCall(	Self,cSoap,; 
	"http://unicasolucoes.com.br/logFTP",; 
	"DOCUMENT","http://unicasolucoes.com.br/",,,; 
	"http://www.unicasolucoes.com.br/ws_uniconnect/ws_uniconnect.asmx")

::Init()
::llogFTPResult      :=  WSAdvValue( oXmlRet,"_LOGFTPRESPONSE:_LOGFTPRESULT:TEXT","boolean",NIL,NIL,NIL,NIL,NIL,NIL) 

END WSMETHOD

oXmlRet := NIL
Return .T.


// WSDL Data Structure UsuariosComunic

WSSTRUCT ws_uniconnect_UsuariosComunic
	WSDATA   nintIdUsuarioComunic      AS int
	WSDATA   cstrLogin                 AS string OPTIONAL
	WSDATA   ctxtSenha                 AS string OPTIONAL
	WSDATA   cstrCNPJUsuario           AS string OPTIONAL
	WSDATA   cstrCNPJVendedor          AS string OPTIONAL
	WSDATA   cstrCNPJComprador         AS string OPTIONAL
	WSDATA   cstrPastaEntrada          AS string OPTIONAL
	WSDATA   cstrPastaSaida            AS string OPTIONAL
	WSDATA   cintIdStatus              AS string OPTIONAL
	WSDATA   cbitIPFixo                AS string OPTIONAL
	WSDATA   cstrIdConf                AS string OPTIONAL
	WSDATA   cstrConfFile              AS string OPTIONAL
	WSMETHOD NEW
	WSMETHOD INIT
	WSMETHOD CLONE
	WSMETHOD SOAPSEND
	WSMETHOD SOAPRECV
ENDWSSTRUCT

WSMETHOD NEW WSCLIENT ws_uniconnect_UsuariosComunic
	::Init()
Return Self

WSMETHOD INIT WSCLIENT ws_uniconnect_UsuariosComunic
Return

WSMETHOD CLONE WSCLIENT ws_uniconnect_UsuariosComunic
	Local oClone := ws_uniconnect_UsuariosComunic():NEW()
	oClone:nintIdUsuarioComunic := ::nintIdUsuarioComunic
	oClone:cstrLogin            := ::cstrLogin
	oClone:ctxtSenha            := ::ctxtSenha
	oClone:cstrCNPJUsuario      := ::cstrCNPJUsuario
	oClone:cstrCNPJVendedor     := ::cstrCNPJVendedor
	oClone:cstrCNPJComprador    := ::cstrCNPJComprador
	oClone:cstrPastaEntrada     := ::cstrPastaEntrada
	oClone:cstrPastaSaida       := ::cstrPastaSaida
	oClone:cintIdStatus         := ::cintIdStatus
	oClone:cbitIPFixo           := ::cbitIPFixo
	oClone:cstrIdConf           := ::cstrIdConf
	oClone:cstrConfFile         := ::cstrConfFile
Return oClone

WSMETHOD SOAPSEND WSCLIENT ws_uniconnect_UsuariosComunic
	Local cSoap := ""
	cSoap += WSSoapValue("intIdUsuarioComunic", ::nintIdUsuarioComunic, ::nintIdUsuarioComunic , "int", .T. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("strLogin", ::cstrLogin, ::cstrLogin , "string", .F. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("txtSenha", ::ctxtSenha, ::ctxtSenha , "string", .F. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("strCNPJUsuario", ::cstrCNPJUsuario, ::cstrCNPJUsuario , "string", .F. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("strCNPJVendedor", ::cstrCNPJVendedor, ::cstrCNPJVendedor , "string", .F. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("strCNPJComprador", ::cstrCNPJComprador, ::cstrCNPJComprador , "string", .F. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("strPastaEntrada", ::cstrPastaEntrada, ::cstrPastaEntrada , "string", .F. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("strPastaSaida", ::cstrPastaSaida, ::cstrPastaSaida , "string", .F. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("intIdStatus", ::cintIdStatus, ::cintIdStatus , "string", .F. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("bitIPFixo", ::cbitIPFixo, ::cbitIPFixo , "string", .F. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("strIdConf", ::cstrIdConf, ::cstrIdConf , "string", .F. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("strConfFile", ::cstrConfFile, ::cstrConfFile , "string", .F. , .F., 0 , NIL, .F.) 
Return cSoap

WSMETHOD SOAPRECV WSSEND oResponse WSCLIENT ws_uniconnect_UsuariosComunic
	::Init()
	If oResponse = NIL ; Return ; Endif 
	::nintIdUsuarioComunic :=  WSAdvValue( oResponse,"_INTIDUSUARIOCOMUNIC","int",NIL,"Property nintIdUsuarioComunic as s:int on SOAP Response not found.",NIL,"N",NIL,NIL) 
	::cstrLogin          :=  WSAdvValue( oResponse,"_STRLOGIN","string",NIL,NIL,NIL,"S",NIL,NIL) 
	::ctxtSenha          :=  WSAdvValue( oResponse,"_TXTSENHA","string",NIL,NIL,NIL,"S",NIL,NIL) 
	::cstrCNPJUsuario    :=  WSAdvValue( oResponse,"_STRCNPJUSUARIO","string",NIL,NIL,NIL,"S",NIL,NIL) 
	::cstrCNPJVendedor   :=  WSAdvValue( oResponse,"_STRCNPJVENDEDOR","string",NIL,NIL,NIL,"S",NIL,NIL) 
	::cstrCNPJComprador  :=  WSAdvValue( oResponse,"_STRCNPJCOMPRADOR","string",NIL,NIL,NIL,"S",NIL,NIL) 
	::cstrPastaEntrada   :=  WSAdvValue( oResponse,"_STRPASTAENTRADA","string",NIL,NIL,NIL,"S",NIL,NIL) 
	::cstrPastaSaida     :=  WSAdvValue( oResponse,"_STRPASTASAIDA","string",NIL,NIL,NIL,"S",NIL,NIL) 
	::cintIdStatus       :=  WSAdvValue( oResponse,"_INTIDSTATUS","string",NIL,NIL,NIL,"S",NIL,NIL) 
	::cbitIPFixo         :=  WSAdvValue( oResponse,"_BITIPFIXO","string",NIL,NIL,NIL,"S",NIL,NIL) 
	::cstrIdConf         :=  WSAdvValue( oResponse,"_STRIDCONF","string",NIL,NIL,NIL,"S",NIL,NIL) 
	::cstrConfFile       :=  WSAdvValue( oResponse,"_STRCONFFILE","string",NIL,NIL,NIL,"S",NIL,NIL) 
Return

// WSDL Data Structure Arquivo

WSSTRUCT ws_uniconnect_Arquivo
	WSDATA   cstrNmArquivo             AS string OPTIONAL
	WSDATA   cbConteudoArquivo         AS base64Binary OPTIONAL
	WSMETHOD NEW
	WSMETHOD INIT
	WSMETHOD CLONE
	WSMETHOD SOAPSEND
	WSMETHOD SOAPRECV
ENDWSSTRUCT

WSMETHOD NEW WSCLIENT ws_uniconnect_Arquivo
	::Init()
Return Self

WSMETHOD INIT WSCLIENT ws_uniconnect_Arquivo
Return

WSMETHOD CLONE WSCLIENT ws_uniconnect_Arquivo
	Local oClone := ws_uniconnect_Arquivo():NEW()
	oClone:cstrNmArquivo        := ::cstrNmArquivo
	oClone:cbConteudoArquivo    := ::cbConteudoArquivo
Return oClone

WSMETHOD SOAPSEND WSCLIENT ws_uniconnect_Arquivo
	Local cSoap := ""
	cSoap += WSSoapValue("strNmArquivo", ::cstrNmArquivo, ::cstrNmArquivo , "string", .F. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("bConteudoArquivo", ::cbConteudoArquivo, ::cbConteudoArquivo , "base64Binary", .F. , .F., 0 , NIL, .F.) 
Return cSoap

WSMETHOD SOAPRECV WSSEND oResponse WSCLIENT ws_uniconnect_Arquivo
	::Init()
	If oResponse = NIL ; Return ; Endif 
	::cstrNmArquivo      :=  WSAdvValue( oResponse,"_STRNMARQUIVO","string",NIL,NIL,NIL,"S",NIL,NIL) 
	::cbConteudoArquivo  :=  WSAdvValue( oResponse,"_BCONTEUDOARQUIVO","base64Binary",NIL,NIL,NIL,"SB",NIL,NIL) 
Return

// WSDL Data Structure ArrayOfArquivo

WSSTRUCT ws_uniconnect_ArrayOfArquivo
	WSDATA   oWSArquivo                AS ws_uniconnect_Arquivo OPTIONAL
	WSMETHOD NEW
	WSMETHOD INIT
	WSMETHOD CLONE
	WSMETHOD SOAPRECV
ENDWSSTRUCT

WSMETHOD NEW WSCLIENT ws_uniconnect_ArrayOfArquivo
	::Init()
Return Self

WSMETHOD INIT WSCLIENT ws_uniconnect_ArrayOfArquivo
	::oWSArquivo           := {} // Array Of  ws_uniconnect_ARQUIVO():New()
Return

WSMETHOD CLONE WSCLIENT ws_uniconnect_ArrayOfArquivo
	Local oClone := ws_uniconnect_ArrayOfArquivo():NEW()
	oClone:oWSArquivo := NIL
	If ::oWSArquivo <> NIL 
		oClone:oWSArquivo := {}
		aEval( ::oWSArquivo , { |x| aadd( oClone:oWSArquivo , x:Clone() ) } )
	Endif 
Return oClone

WSMETHOD SOAPRECV WSSEND oResponse WSCLIENT ws_uniconnect_ArrayOfArquivo
	Local nRElem1, oNodes1, nTElem1
	::Init()
	If oResponse = NIL ; Return ; Endif 
	oNodes1 :=  WSAdvValue( oResponse,"_ARQUIVO","Arquivo",{},NIL,.T.,"O",NIL,NIL) 
	nTElem1 := len(oNodes1)
	For nRElem1 := 1 to nTElem1 
		If !WSIsNilNode( oNodes1[nRElem1] )
			aadd(::oWSArquivo , ws_uniconnect_Arquivo():New() )
			::oWSArquivo[len(::oWSArquivo)]:SoapRecv(oNodes1[nRElem1])
		Endif
	Next
Return

// WSDL Data Structure Log_InOut

WSSTRUCT ws_uniconnect_Log_InOut
	WSDATA   nintIdLogUcomunic         AS int
	WSDATA   cstrNmUsuario             AS string OPTIONAL
	WSDATA   oWSintIdDirecao           AS ws_uniconnect_Direcao
	WSDATA   cstrNmArquivo             AS string OPTIONAL
	WSDATA   cstrNmMaquinaCliente      AS string OPTIONAL
	WSDATA   cstrPathRemoto            AS string OPTIONAL
	WSDATA   cdtTimeStamp              AS dateTime
	WSDATA   cdtDataCriacao            AS dateTime
	WSDATA   nfTamanho                 AS long
	WSMETHOD NEW
	WSMETHOD INIT
	WSMETHOD CLONE
	WSMETHOD SOAPSEND
ENDWSSTRUCT

WSMETHOD NEW WSCLIENT ws_uniconnect_Log_InOut
	::Init()
Return Self

WSMETHOD INIT WSCLIENT ws_uniconnect_Log_InOut
Return

WSMETHOD CLONE WSCLIENT ws_uniconnect_Log_InOut
	Local oClone := ws_uniconnect_Log_InOut():NEW()
	oClone:nintIdLogUcomunic    := ::nintIdLogUcomunic
	oClone:cstrNmUsuario        := ::cstrNmUsuario
	oClone:oWSintIdDirecao      := IIF(::oWSintIdDirecao = NIL , NIL , ::oWSintIdDirecao:Clone() )
	oClone:cstrNmArquivo        := ::cstrNmArquivo
	oClone:cstrNmMaquinaCliente := ::cstrNmMaquinaCliente
	oClone:cstrPathRemoto       := ::cstrPathRemoto
	oClone:cdtTimeStamp         := ::cdtTimeStamp
	oClone:cdtDataCriacao       := ::cdtDataCriacao
	oClone:nfTamanho            := ::nfTamanho
Return oClone

WSMETHOD SOAPSEND WSCLIENT ws_uniconnect_Log_InOut
	Local cSoap := ""
	cSoap += WSSoapValue("intIdLogUcomunic", ::nintIdLogUcomunic, ::nintIdLogUcomunic , "int", .T. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("strNmUsuario", ::cstrNmUsuario, ::cstrNmUsuario , "string", .F. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("intIdDirecao", ::oWSintIdDirecao, ::oWSintIdDirecao , "Direcao", .T. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("strNmArquivo", ::cstrNmArquivo, ::cstrNmArquivo , "string", .F. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("strNmMaquinaCliente", ::cstrNmMaquinaCliente, ::cstrNmMaquinaCliente , "string", .F. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("strPathRemoto", ::cstrPathRemoto, ::cstrPathRemoto , "string", .F. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("dtTimeStamp", ::cdtTimeStamp, ::cdtTimeStamp , "dateTime", .T. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("dtDataCriacao", ::cdtDataCriacao, ::cdtDataCriacao , "dateTime", .T. , .F., 0 , NIL, .F.) 
	cSoap += WSSoapValue("fTamanho", ::nfTamanho, ::nfTamanho , "long", .T. , .F., 0 , NIL, .F.) 
Return cSoap

// WSDL Data Enumeration Direcao

WSSTRUCT ws_uniconnect_Direcao
	WSDATA   Value                     AS string
	WSDATA   cValueType                AS string
	WSDATA   aValueList                AS Array Of string
	WSMETHOD NEW
	WSMETHOD CLONE
	WSMETHOD SOAPSEND
	WSMETHOD SOAPRECV
ENDWSSTRUCT

WSMETHOD NEW WSCLIENT ws_uniconnect_Direcao
	::Value := NIL
	::cValueType := "string"
	::aValueList := {}
	aadd(::aValueList , "Entrada" )
	aadd(::aValueList , "Saida" )
Return Self

WSMETHOD SOAPSEND WSCLIENT ws_uniconnect_Direcao
	Local cSoap := "" 
	cSoap += WSSoapValue("Value", ::Value, NIL , "string", .F. , .F., 3 , NIL, .F.) 
Return cSoap

WSMETHOD SOAPRECV WSSEND oResponse WSCLIENT ws_uniconnect_Direcao
	::Value := NIL
	If oResponse = NIL ; Return ; Endif 
	::Value :=  oResponse:TEXT
Return 

WSMETHOD CLONE WSCLIENT ws_uniconnect_Direcao
Local oClone := ws_uniconnect_Direcao():New()
	oClone:Value := ::Value
Return oClone


