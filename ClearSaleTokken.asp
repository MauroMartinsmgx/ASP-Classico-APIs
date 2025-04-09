<%
' ============================================
' Configuração da Conexão com o Banco de Dados
' ============================================
Conn_SQL = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=;PORT=;DATABASE=;USER=;PASSWORD=;OPTION=3;"
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Conn_SQL

' ============================================
' Função Principal para Processar Pedidos
' ============================================
Set RS_PedidosClearSale = CreateObject("ADODB.Recordset")
RS_PedidosClearSale.Open "SELECT * FROM tb_clearsale_pedidos WHERE status_interno_pedido = '1' ORDER BY id_clearsale DESC LIMIT 6", Conexao

If Not RS_PedidosClearSale.EOF Then
    While Not RS_PedidosClearSale.EOF
        retorno = ClearSale(RS_PedidosClearSale("codigo_pedido"))

        str_returno_ClearSale = ""
        If retorno <> "" Then
            If InStr(retorno, "The request is invalid.") = 0 Then
                arry_retorno = Split(retorno, ";")
                str_returno_ClearSale = ", code_retorno_clearsale = '" & arry_retorno(0) & "', status_retorno_clearsale = '" & arry_retorno(1) & "', score_retorno_clearsale = '" & arry_retorno(2) & "', json_retorno_clearsale='" & arry_retorno(4) & "' "
            Else
                str_returno_ClearSale = ", json_retorno_clearsale='" & retorno & "' "
            End If
        End If

        Conexao.Execute "UPDATE tb_clearsale_pedidos SET status_interno_pedido = '2', desc_status_interno_pedido = 'Enviado para ClearSale' " & str_returno_ClearSale & " WHERE codigo_pedido = '" & RS_PedidosClearSale("codigo_pedido") & "'"
        RS_PedidosClearSale.MoveNext
    Wend
End If

' ============================================
' Função ClearSale
' Envia os dados do pedido para a API ClearSale
' ============================================
Function ClearSale(codigo_pedido)
    ' Configuração de Login
    Login_clearSale = ""
    Senha_clearSale = ""

    ' Recupera o Token de Autenticação
    Set RS_Token = CreateObject("ADODB.Recordset")
    RS_Token.Open "SELECT * FROM ClearSale_Token_new ORDER BY codigo_chave DESC LIMIT 1", Conexao
    If Not RS_Token.EOF Then
        token = RS_Token("token")
    End If

    If token <> "" Then
        ' Recupera os dados do pedido
        Set RS_v = CreateObject("ADODB.Recordset")
        RS_v.Open "SELECT * FROM Pedidos WHERE codigo_pedido = '" & codigo_pedido & "'", Conexao

        If Not RS_v.EOF Then
            ' Processa os dados do pedido
            ' (Adicione aqui o processamento necessário, como formatação de valores e criação do JSON)

            ' Envia os dados para a API ClearSale
            Set objhttpPag4 = Server.CreateObject("MSXML2.XMLHTTP.6.0")
            objhttpPag4.Open "POST", "https://api.clearsale.com.br/v1/orders/", False
            objhttpPag4.SetRequestHeader "Content-Type", "application/json"
            objhttpPag4.SetRequestHeader "Authorization", "Bearer " & token
            objhttpPag4.Send var_jsonClearSale

            ' Processa a resposta da API
            strResponseTextw4 = objhttpPag4.ResponseText
            If InStr(strResponseTextw4, "code") Then
                ' Processa o JSON de resposta
                ' (Adicione aqui o processamento do JSON de resposta)
            Else
                Retorno = "The request is invalid.;" & codigo_pedido & ";" & strResponseTextw4
            End If
        End If
    End If

    ClearSale = Retorno
End Function

' ============================================
' Função Encriptor
' Realiza a criptografia e descriptografia de dados
' ============================================
Function Encriptor(FctDadoEncr, FctAcao)
    Set oEncryptor = Server.CreateObject("Dynu.Encrypt")
    ChaveCripto = ""

    If FctAcao = "encriptar" Then
        Encriptor = oEncryptor.Encrypt(FctDadoEncr, ChaveCripto)
    ElseIf FctAcao = "decriptar" Then
        Encriptor = oEncryptor.Decrypt(FctDadoEncr, ChaveCripto)
    End If

    Set oEncryptor = Nothing
End Function

' ============================================
' Função LimpaLixo_new
' Remove caracteres indesejados de uma string
' ============================================
Function LimpaLixo_new(input)
    lixo = Array("'")
    textoOK = input

    For i = 0 To UBound(lixo)
        textoOK = Replace(textoOK, lixo(i), "")
    Next

    LimpaLixo_new = textoOK
End Function

' ============================================
' Função ConverterDataHoraOFIP
' Converte uma data para o formato ISO 8601
' ============================================
Function ConverterDataHoraOFIP(ConDataHora)
    ConverterDataHoraOFIP = Year(ConDataHora) & "-" & Right("0" & Month(ConDataHora), 2) & "-" & Right("0" & Day(ConDataHora), 2) & " " & Right("0" & Hour(ConDataHora), 2) & ":" & Right("0" & Minute(ConDataHora), 2) & ":" & Right("0" & Second(ConDataHora), 2)
End Function

' ============================================
' Finaliza a Conexão com o Banco de Dados
' ============================================
Conexao.Close
Set Conexao = Nothing
%>