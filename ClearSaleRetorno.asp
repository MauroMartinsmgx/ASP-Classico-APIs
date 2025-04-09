<!--#INCLUDE FILE="../../../com_estilo.asp"-->
<script language="javascript" runat="server" src="../../../cielo30/json2.asp"></script>
<!--#include FILE="aspJSON1.17.asp" -->
<%

' ============================================
' Configuração da Conexão com o Banco de Dados
' ============================================
Conn_SQL = "DRIVER={MySQL ODBC 5.1 Driver};SERVER={SEU_SERVIDOR};PORT=3306;DATABASE={SEU_BANCO};USER={SEU_USUARIO};PASSWORD={SUA_SENHA};OPTION=3;"
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Conn_SQL

' ============================================
' Função: BytesToStr
' Descrição: Converte um array de bytes em uma string.
' Parâmetros:
'   - bytes: Array de bytes a ser convertido.
' Retorno:
'   - String convertida.
' ============================================
Function BytesToStr(bytes)
    Dim Stream
    Set Stream = Server.CreateObject("Adodb.Stream")
    Stream.Type = 1 'adTypeBinary
    Stream.Open
    Stream.Write bytes
    Stream.Position = 0
    Stream.Type = 2 'adTypeText
    Stream.Charset = "iso-8859-1"
    BytesToStr = Stream.ReadText
    Stream.Close
    Set Stream = Nothing
End Function

' ============================================
' Função: ClearSale
' Descrição: Consulta o status de um pedido na API ClearSale.
' Parâmetros:
'   - codigo_pedido: Código do pedido a ser consultado.
'   - jsonRetorno: JSON recebido da ClearSale.
' Retorno:
'   - Retorno da API ClearSale.
' ============================================
Function ClearSale(codigo_pedido, jsonRetorno)
    ' Configuração de Login
    Login_clearSale = "SEU_LOGIN"
    Senha_clearSale = "SUA_SENHA"

    ' Recupera o Token de Autenticação
    Set RS_Token = CreateObject("ADODB.Recordset")
    RS_Token.Open "SELECT * FROM ClearSale_Token_new ORDER BY codigo_chave DESC LIMIT 1", Conexao
    If Not RS_Token.EOF Then
        token = RS_Token("token")
    End If

    If token <> "" Then
        ' URL de consulta
        url_consulta = "https://api.clearsale.com.br/v1/orders/" & codigo_pedido & "/status"

        ' Faz a requisição para a API ClearSale
        Set objhttpPag4 = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
        objhttpPag4.Open "GET", url_consulta, False
        objhttpPag4.SetRequestHeader "Accept", "application/json"
        objhttpPag4.SetRequestHeader "Authorization", "Bearer " & token
        objhttpPag4.Send
        strResponseTextw4 = objhttpPag4.ResponseText

        ' Processa o JSON de resposta
        Dim variavel_json2
        variavel_json2 = strResponseTextw4

        If variavel_json2 <> "" Then
            Set oJSON = New aspJSON
            oJSON.loadJSON(variavel_json2)

            ' Extrai informações do JSON
            codigo_pedido = oJSON.data("code")
            status = oJSON.data("status")
            score = oJSON.data("score")

            ' Define a descrição do status
            Select Case status
                Case "APA": status_pedido_clearsale_desc = "APA - Pedido aprovado automaticamente"
                Case "APM": status_pedido_clearsale_desc = "APM - Pedido aprovado manualmente"
                Case "RPM": status_pedido_clearsale_desc = "RPM - Pedido Reprovado sem Suspeita"
                Case "AMA": status_pedido_clearsale_desc = "AMA - Pedido está em fila"
                Case "NVO": status_pedido_clearsale_desc = "NVO - Pedido importado classificado"
                Case "SUS": status_pedido_clearsale_desc = "SUS - Pedido suspeita de fraude"
                Case "CAN": status_pedido_clearsale_desc = "CAN - Cancelado pelo cliente"
                Case "FRD": status_pedido_clearsale_desc = "FRD - Pedido Fraude Confirmada"
                Case "RPA": status_pedido_clearsale_desc = "RPA - Pedido Reprovado Automaticamente"
                Case "RPP": status_pedido_clearsale_desc = "RPP - Pedido reprovado automaticamente"
                Case "APP": status_pedido_clearsale_desc = "APP - Pedido aprovado automaticamente"
                Case Else: status_pedido_clearsale_desc = "Status desconhecido"
            End Select

            ' Atualiza o status do pedido no banco de dados
            If codigo_pedido <> "" Then
                Conexao.Execute "UPDATE tb_clearsale_pedidos SET json_retorno_notificacao = '" & jsonRetorno & "', json_retorno_consulta = '" & variavel_json2 & "', status_interno_pedido = '3', desc_status_interno_pedido = 'Retorno da ClearSale', status_pedido_clearsale = '" & status & "', status_pedido_clearsale_desc = '" & status_pedido_clearsale_desc & "' WHERE codigo_pedido = '" & codigo_pedido & "'"
            End If
        End If
    End If

    ClearSale = "Processamento concluído"
End Function

' ============================================
' Processamento do JSON de Retorno
' ============================================
If Request.TotalBytes > 0 Then
    Dim lngBytesCount, json
    lngBytesCount = Request.TotalBytes
    json = BytesToStr(Request.BinaryRead(lngBytesCount))

    ' Salva o JSON recebido no banco de dados
    Conexao.Execute "INSERT INTO tb_retorno_clearsale (json_retorno) VALUES ('" & json & "')"

    If json <> "" Then
        Dim variavel_json
        variavel_json = json

        Set oJSON = New aspJSON
        oJSON.loadJSON(variavel_json)

        ' Extrai informações do JSON
        codigo_pedido = oJSON.data("code")
        status = oJSON.data("type")

        ' Chama a função ClearSale para processar o pedido
        If status = "status" Then
            Call ClearSale(codigo_pedido, variavel_json)
        End If
    End If
End If

' ============================================
' Finaliza a Conexão com o Banco de Dados
' ============================================
Conexao.Close
Set Conexao = Nothing

%>