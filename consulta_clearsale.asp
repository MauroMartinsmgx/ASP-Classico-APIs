<%
function ConsultaClearSale(pedido,token)

        'url_consulta = "https://homologacao.clearsale.com.br/api/v1/orders/"&pedido&"/status"

        url_consulta = "https://api.clearsale.com.br/v1/orders/"&pedido&"/status"

        Set objhttpPag = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
        objhttpPag.Open "GET", url_consulta , False
        objhttpPag.SetRequestHeader "Accept", "application/json"
        objhttpPag.SetRequestHeader "Authorization", "Bearer "&token&""
        objhttpPag.Send
        strResponseStatusw = objhttpPag.Status & " " & objhttpPag.StatusText
        strResponseTextw = objhttpPag.ResponseText

        ConsultaClearSale = strResponseTextw

End Function
%>
