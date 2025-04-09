<%
' ============================================
' Configuração de Segurança
' ============================================
Response.AddHeader "X-XSS-Protection", "1; mode=block"

' ============================================
' Função: sqlInjectionCom
' Descrição: Verifica e bloqueia tentativas de SQL Injection.
' ============================================
Function sqlInjectionCom()
    palavrasDoMal = Array("<script>", "select", "drop", "insert", "delete", "xp_", "1=1", "update", "truncate", "union", "--", "/*", "*/", "alert", "onmouseover")

    ' Verifica QueryString
    For Each item In Request.QueryString
        For j = LBound(palavrasDoMal) To UBound(palavrasDoMal)
            If InStr(LCase(Request.QueryString(item)), LCase(palavrasDoMal(j))) > 0 Then
                Response.Redirect("/index.asp")
            End If
        Next
    Next

    ' Verifica Form
    For Each item In Request.Form
        For j = LBound(palavrasDoMal) To UBound(palavrasDoMal)
            If InStr(LCase(Request.Form(item)), LCase(palavrasDoMal(j))) > 0 Then
                Response.Redirect("/index.asp")
            End If
        Next
    Next

    ' Verifica Cookies
    For Each item In Request.Cookies
        For j = LBound(palavrasDoMal) To UBound(palavrasDoMal)
            If InStr(LCase(Request.Cookies(item)), LCase(palavrasDoMal(j))) > 0 Then
                Response.Redirect("/index.asp")
            End If
        Next
    Next
End Function

sqlInjectionCom()

' ============================================
' Função: LimpaLixo_Prod
' Descrição: Remove caracteres indesejados de uma string.
' Parâmetros:
'   - input: String a ser limpa.
' Retorno:
'   - String limpa.
' ============================================
Function LimpaLixo_Prod(input)
    lixo = Array("<script>", ";", "--", "/*", "*/", "'", "<", ">", "$", "#")
    textoOK = input

    For i = 0 To UBound(lixo)
        textoOK = Replace(textoOK, lixo(i), "")
    Next

    LimpaLixo_Prod = Replace(textoOK, "'", "")
End Function

' ============================================
' Configuração da Conexão com o Banco de Dados
' ============================================
Conn_SQL = "DRIVER={MySQL ODBC 5.1 Driver};SERVER={SEU_SERVIDOR};PORT=3306;DATABASE={SEU_BANCO};USER={SEU_USUARIO};PASSWORD={SUA_SENHA};OPTION=3;"
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Conn_SQL

' ============================================
' Função: Encriptor
' Descrição: Realiza a criptografia e descriptografia de dados.
' Parâmetros:
'   - FctDadoEncr: Dado a ser criptografado ou descriptografado.
'   - FctAcao: Ação a ser realizada ("encriptar" ou "decriptar").
' Retorno:
'   - Dado criptografado ou descriptografado.
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
' Função: alteracartao
' Descrição: Atualiza os dados do cartão de crédito do usuário.
' ============================================
Function alteracartao()
    Set RS_v = CreateObject("ADODB.Recordset")
    RS_v.Open "SELECT user_id FROM Usuarios WHERE user_id = '" & Session("logado") & "'", Conexao

    If Not RS_v.EOF Then
        numero_cartao = LimpaLixo_Prod(Request("numero_cartao"))
        nome_cartao = LimpaLixo_Prod(Request("nome_cartao"))
        cod_seguranca = LimpaLixo_Prod(Request("cod_seguranca"))
        bandeira = LimpaLixo_Prod(Request("bandeira"))
        validade_mes = LimpaLixo_Prod(Request("validade_mes"))
        validade_ano = LimpaLixo_Prod(Request("validade_ano"))

        numero_cartao_cript = Encriptor(numero_cartao, "encriptar")
        mes_val_cript = Encriptor(validade_mes, "encriptar")
        ano_val_cript = Encriptor(validade_ano, "encriptar")

        Conexao.Execute "UPDATE Pedidos SET cartao_portador_card = '" & numero_cartao_cript & "', mes_portador_card = '" & mes_val_cript & "', ano_portador_card = '" & ano_val_cript & "' WHERE user_ID = '" & Session("logado") & "'"
        Response.Write("OK")
    Else
        Response.Write("falha")
    End If
End Function

' ============================================
' Função: cobramensal
' Descrição: Realiza a cobrança mensal do pedido.
' ============================================
Function cobramensal()
    Set RS_v = CreateObject("ADODB.Recordset")
    RS_v.Open "SELECT subtotal, taxa_envio, total_desc_cupom, codigo_pedido  FROM Pedidos WHERE codigo_pedido = '" & LimpaLixo_Prod(Request("pedido")) & "'", Conexao

    If Not RS_v.EOF Then
        valor = Replace(FormatNumber(RS_v("subtotal") + RS_v("taxa_envio") - RS_v("total_desc_cupom"), 2), ",", "")
        var_jsonCielo = "{""MerchantOrderId"":""" & RS_v("codigo_pedido") & """,""Payment"":{""Amount"":" & valor & "}}"

        Set objhttpPag3 = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
        objhttpPag3.Open "POST", "https://api.mundipagg.com/core/v1/charges", False
        objhttpPag3.SetRequestHeader "Content-Type", "application/json"
        objhttpPag3.Send var_jsonCielo

        Response.Write(objhttpPag3.ResponseText)
    Else
        Response.Write("falha")
    End If
End Function

' ============================================
' Finaliza a Conexão com o Banco de Dados
' ============================================
Conexao.Close
Set Conexao = Nothing
%>