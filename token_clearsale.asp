<%

        Login_clearSale = ""    
        Senha_clearSale = ""    
        '=======================Teste=============================
        ''url_producao_authenticate = "https://homologacao.clearsale.com.br/api/v1/authenticate"
        '=======================Teste=============================

        url_producao_authenticate = "https://api.clearsale.com.br/v1/authenticate"

        var_jsonClearSale_authenticate = " {""name"": """&Login_clearSale&""",""password"": """&Senha_clearSale&"""}"

        Set objhttpPag3 = Server.CreateObject ("MSXML2.XMLHTTP.6.0")
        objhttpPag3.Open "POST", url_producao_authenticate , False
        objhttpPag3.SetRequestHeader "Content-Type", "application/json"
        objhttpPag3.Send var_jsonClearSale_authenticate
        strResponseStatusw3 = objhttpPag3.Status & " " & objhttpPag3.StatusText
        strResponseTextw3 = objhttpPag3.ResponseText

        response.write(strResponseTextw3)
%>
