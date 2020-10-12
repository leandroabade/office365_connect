Option Compare Database
Option Explicit

Sub subTest()

    GetO365SPO_Token <usuario>, <senha>, "https://fedaa.sharepoint.com"

End Sub

'Pega o token
Private Function GetO365SPO_Token(ByVal strUsuario As String, ByVal strSenha As String, ByVal strSharepoint As String) As String
    
    'Declaranção de variáveis
    Dim strSiteLogin As String
    Dim strMensagem As String
    Dim Request As Object
    Dim intNum As Integer
    Dim strResposta As String
    Dim lngLen As Long
    Dim strToken As String

    'Site API Microsoft para pegar token
    strSiteLogin = "https://login.microsoftonline.com/extSTS.srf"
    
    'XML padrão
    strMensagem = "<s:Envelope xmlns:s=""http://www.w3.org/2003/05/soap-envelope"" xmlns:a=""http://www.w3.org/2005/08/addressing"" xmlns:u=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd"">" & _
                        "<s:Header><a:Action s:mustUnderstand=""1"">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action><a:ReplyTo><a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address></a:ReplyTo><a:To s:mustUnderstand=""1"">https://login.microsoftonline.com/extSTS.srf</a:To><o:Security s:mustUnderstand=""1"" xmlns:o=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"">" & _
                            "<o:UsernameToken>" & _
                                "<o:Username>" & strUsuario & "</o:Username>" & _
                                "<o:Password>" & strSenha & "</o:Password>" & _
                            "</o:UsernameToken></o:Security></s:Header>" & _
                        "<s:Body><t:RequestSecurityToken xmlns:t=""http://schemas.xmlsoap.org/ws/2005/02/trust""><wsp:AppliesTo xmlns:wsp=""http://schemas.xmlsoap.org/ws/2004/09/policy"">" & _
                            "<a:EndpointReference>" & _
                                "<a:Address>" & strSharepoint & "</a:Address>" & _
                            "</a:EndpointReference>" & _
                            "</wsp:AppliesTo><t:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</t:KeyType><t:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</t:RequestType><t:TokenType>urn:oasis:names:tc:SAML:1.0:assertion</t:TokenType></t:RequestSecurityToken></s:Body>" & _
                    "</s:Envelope>"
    
    'Recupera o tamanho da mensagem
    lngLen = Len(strMensagem)
    
    'Instancia objeto XMLHTTP
    Set Request = CreateObject("MSXML2.XMLHTTP")
    
    'Abre a conexão para execução
    Request.Open "POST", strSiteLogin, False
    
    'Prepara o reader
    Request.setRequestHeader "Content-Type", "application/xml"
    Request.setRequestHeader "User-Agent", "Mozilla/5.0 (compatible; MSIE 9.0; Windows Phone OS 7.5; Trident/5.0; IEMobile/9.0)"
    Request.setRequestHeader "host", "login.microsoftonline.com"
    Request.setRequestHeader "Content-Length", lngLen
    
    'Envia o XML
    Request.Send (strMensagem)

    'Se a resposta for 200
    If Request.Status = 200 Then
    
        'Recupera a resposta
        strResposta = Request.responseText
        
        'Extrai o Token
        strToken = GetO365SPO_RETToken(strResposta)
        
        'Exibe o token
        Debug.Print strToken
        
        'Retorna a informação para a função
        GetO365SPO_Token = strToken
    End If

End Function

'Extrai o tokem do retorno
Private Function GetO365SPO_RETToken(ByVal strResposta As String) As String
    
    'Declaração de variáveis
    Dim strNomeAvaliacao As String
    Dim intP01 As Integer
    Dim intP02 As Integer
    Dim strToken As String
    
    'Limpa espaços
    strResposta = RTrim(LTrim(strResposta))
    
    'Identfica as posições de início e fim da leitura
    intP01 = InStr(1, strResposta, ">t=") + 1
    intP02 = InStr(1, strResposta, "&amp;p=</") - 1
    intP02 = intP02 - intP01
    
    'Faz a leitura
    strToken = Mid(strResposta, intP01, intP02)
    
    'Retorna a informação para a função
    GetO365SPO_RETToken = strToken

End Function
