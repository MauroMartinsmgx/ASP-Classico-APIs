<%
        Function GeraTokenClearSale(Login_clearSale,Senha_clearSale) 
                '=======================Teste=============================
                ''url_producao_authenticate = "https://homologacao.clearsale.com.br/api/v1/authenticate"
                '=======================Teste=============================

                url_producao_authenticate = "https://api.clearsale.com.br/v1/authenticate"

                var_jsonClearSale_authenticate = " {""name"": """&Login_clearSale&""",""password"": """&Senha_clearSale&"""}"

                Set objhttpPag = Server.CreateObject ("MSXML2.XMLHTTP.6.0")
                objhttpPag.Open "POST", url_producao_authenticate , False
                objhttpPag.SetRequestHeader "Content-Type", "application/json"
                objhttpPag.Send var_jsonClearSale_authenticate
                strResponseStatusw = objhttpPag.Status & " " & objhttpPag.StatusText
                strResponseTextw = objhttpPag.ResponseText

                GeraTokenClearSale = strResponseTextw
        End Function
%>

