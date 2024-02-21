<%
' Verifica se a requisição é um POST (ou seja, se o formulário foi submetido)
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' Recupera os dados do formulário
    Dim email, senha
    email = Request.Form("email")
    senha = Request.Form("senha")
    
    ' Validação básica (você pode adicionar validações mais robustas conforme necessário)
    If email <> "" And senha <> "" Then
        ' Conecta ao banco de dados MariaDB
        Dim conn
        Set conn = Server.CreateObject("ADODB.Connection")
        conn.Open "DRIVER={MariaDB ODBC 3.1 Driver}; SERVER=localhost; DATABASE=novo; USER=root; PASSWORD=root; OPTION=3;"
        
        ' Query SQL para inserir os dados na tabela
        Dim sql
        sql = "INSERT INTO usuario (email, PASSWORD) VALUES ('" & email & "', '" & senha & "')"
        
        ' Executa a query SQL
        conn.Execute(sql)
        
        ' Fecha a conexão
        conn.Close
        Set conn = Nothing
        
        Response.Write "Dados inseridos com sucesso!"
    Else
        Response.Write "Por favor, preencha todos os campos do formulário."
    End If
End If
%>

<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html;charset=UTF-8">
    <title>Formulário de Inserção de Dados</title>
</head>
<body>
    <h1>Formulário de Inserção de Dados</h1>
    
    <form method="post">
        <label for="email">Email:</label>
        <input type="email" id="email" name="email" required><br><br>
        
        <label for="senha">Senha:</label>
        <input type="password" id="senha" name="senha" required><br><br>
        
        <input type="submit" value="Enviar">
    </form>
</body>
</html>
