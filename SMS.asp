<%
' ============================================
' Configuração de Cache
' ============================================
Response.CacheControl = "no-cache"
Response.addHeader "pragma", "no-cache"
Response.CacheControl = "Private"
Response.Expires = -1

' ============================================
' Configuração da Conexão com o Banco de Dados
' ============================================
Conn_SQL = "DRIVER={MySQL ODBC 5.1 Driver};SERVER={SEU_SERVIDOR};PORT=3306;DATABASE={SEU_BANCO};USER={SEU_USUARIO};PASSWORD={SUA_SENHA};OPTION=3;"
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open Conn_SQL

' ============================================
' Recupera Configurações Gerais
' ============================================
Set RS_Layout1 = Server.CreateObject("ADODB.Recordset")
RS_Layout1.CursorLocation = 3
RS_Layout1.CursorType = 0
RS_Layout1.LockType = 3
RS_Layout1.Open "SELECT nome_site FROM Geral WHERE codigo_chave = 1", Conexao

' Inclui funções auxiliares
%>
<!--#INCLUDE FILE="funcoes-cobranca.asp"-->
<script language="javascript" runat="server" src="../cielo30/json2.asp"></script>
<%
' ============================================
' Função: EnviarSMS
' Descrição: Envia um SMS para o cliente.
' Parâmetros:
'   - celular: Número do celular do cliente.
'   - mensagem: Mensagem a ser enviada.
' ============================================
Function EnviarSMS(celular, mensagem)
    celular = Replace(celular, " ", "")
    celular = Replace(celular, "(", "")
    celular = Replace(celular, ")", "")
    celular = Replace(celular, "-", "")

    Set objSMS = Server.CreateObject("MSXML2.XMLHTTP.6.0")
    objSMS.Open "GET", "https://mex10.com/api/shortcode.aspx?token={SEU_TOKEN}&t=send&u={SEU_USUARIO}&p={SUA_SENHA}&n=" & celular & "&m=" & mensagem & "&i=FEIRANABOX", False
    objSMS.Send
    EnviarSMS = objSMS.ResponseText
    Set objSMS = Nothing
End Function

' ============================================
' Processamento de Pedidos
' ============================================
data_inicio_b = ConverterDataHoraOF(FormataDia(Now() + 2) & "/" & FormataMes(Now() + 2) & "/" & Year(Now() + 2) & " " & "00:00:01")
data_fim_b = ConverterDataHoraOF(FormataDia(Now() + 2) & "/" & FormataMes(Now() + 2) & "/" & Year(Now() + 2) & " " & "23:59:59")

Set RS_pedidos = Server.CreateObject("ADODB.Recordset")
RS_pedidos.CursorLocation = 3
RS_pedidos.CursorType = 0
RS_pedidos.LockType = 3
RS_pedidos.Open "SELECT codigo_pedido, telefone_cobranca, nome_cobranca, user_ID FROM Pedidos WHERE forma_pagamento = 'Cielo' AND finalizado = 'sim' AND Pedidos.data_pedido > '" & data_inicio_b & "' AND Pedidos.data_pedido < '" & data_fim_b & "' AND status_interno IN ('6','11') AND tentativas = 1 AND ((tipoEntrega <> '2' AND tipoEntrega <> '0') OR tipoEntrega IS NULL) AND ocorrencia <> '1' ORDER BY data_pedido ASC LIMIT 3", Conexao

If Not RS_pedidos.EOF Then
    While Not RS_pedidos.EOF
        ' Atualiza tentativa no banco de dados
        Conexao.Execute "UPDATE Pedidos SET tentativas = '2' WHERE codigo_pedido = '" & RS_pedidos("codigo_pedido") & "'"

        ' Envia SMS
        mensagem = " "
        celular = RS_pedidos("telefone_cobranca")
        EnviarSMS celular, mensagem

        RS_pedidos.MoveNext
    Wend
End If

' ============================================
' Finaliza a Conexão com o Banco de Dados
' ============================================
Conexao.Close
Set Conexao = Nothing
%>