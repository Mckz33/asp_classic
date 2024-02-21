<%
' Verifica se a requisição é um POST (ou seja, se o formulário foi submetido)
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' Recupera os dados do formulário
    Dim userEmail, newEmail, newPassword
    userEmail = Request.Form("userEmail")
    newEmail = Request.Form("newEmail")
    newPassword = Request.Form("newPassword")
    
    ' Validação básica (você pode adicionar validações adicionais conforme necessário)
    If userEmail <> "" And newEmail <> "" And newPassword <> "" Then
        ' Conecta ao banco de dados MariaDB
        Dim conn
        Set conn = Server.CreateObject("ADODB.Connection")
        conn.Open "DRIVER={MariaDB ODBC 3.1 Driver}; SERVER=localhost; DATABASE=novo; USER=root; PASSWORD=root; OPTION=3;"
        
        ' Query SQL para atualizar os dados do usuário com o email especificado
        Dim sql
        sql = "UPDATE usuario SET email = '" & newEmail & "', PASSWORD = '" & newPassword & "' WHERE email = '" & userEmail & "'"
        
        ' Executa a query SQL
        conn.Execute(sql)
        
        ' Fecha a conexão
        conn.Close
        Set conn = Nothing
        
        Response.Write "Dados do usuário atualizados com sucesso!"
    Else
        Response.Write "Por favor, preencha todos os campos do formulário."
    End If
End If
%>

<!DOCTYPE html>
<html>
<head>
    <title>Formulário de Atualização de Dados</title>
    <meta http-equiv="Content-Type" content="text/html;charset=UTF-8">
</head>
<body>
    <h1>Formulário de Atualização de Dados</h1>
    
    <form method="post">
        <label for="userEmail">Email do Usuário:</label>
        <input type="email" id="userEmail" name="userEmail" required><br><br>
        
        <label for="newEmail">Novo Email:</label>
        <input type="email" id="newEmail" name="newEmail" required><br><br>
        
        <label for="newPassword">Nova Senha:</label>
        <input type="password" id="newPassword" name="newPassword" required><br><br>
        
        <input type="submit" value="Atualizar Dados">
    </form>
</body>
</html>
