<%
' Verifica se a requisição é um POST (ou seja, se o formulário foi submetido)
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' Verifica se o parâmetro "userEmail" está presente na URL
    If Request.QueryString("userEmail") <> "" Then
        ' Recupera o email do usuário a ser excluído do parâmetro de consulta
        Dim userEmail
        userEmail = Request.QueryString("userEmail")
        
        ' Conecta ao banco de dados MariaDB
        Dim conn
        Set conn = Server.CreateObject("ADODB.Connection")
        conn.Open "DRIVER={MariaDB ODBC 3.1 Driver}; SERVER=localhost; DATABASE=novo; USER=root; PASSWORD=root; OPTION=3;"
        
        ' Query SQL para excluir o usuário com o email especificado
        Dim sql
        sql = "DELETE FROM usuario WHERE email = '" & userEmail & "'"
        
        ' Executa a query SQL
        conn.Execute(sql)
        
        ' Fecha a conexão
        conn.Close
        Set conn = Nothing
        
        Response.Write "Usuário excluído com sucesso!"
    Else
        Response.Write "Por favor, forneça o email do usuário que deseja excluir."
    End If
End If
%>


<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html;charset=UTF-8">
    <title>Exclusão de Email</title>
</head>
<body>
    <h1>Excluir Email</h1>
    
    <label for="userEmail">Email do Usuário:</label>
    <input type="email" id="userEmail" name="userEmail" required><br><br>
    
    <button onclick="deleteUser()">Excluir Email</button>

    <script>
    function deleteUser() {
        var userEmail = document.getElementById("userEmail").value;

        // Crie uma nova solicitação XMLHttpRequest
        var xhr = new XMLHttpRequest();

        // Configurar a solicitação para uma solicitação POST e adicionar cabeçalho personalizado para DELETE
        xhr.open("POST", "deletar.asp?userEmail=" + encodeURIComponent(userEmail), true);
        xhr.setRequestHeader("X-HTTP-Method-Override", "DELETE");

        // Definir o que fazer em caso de resposta
        xhr.onload = function() {
            if (xhr.status >= 200 && xhr.status < 300) {
                alert("Email excluído com sucesso!");
                // Faça qualquer outra coisa que você precise fazer após a exclusão bem-sucedida
            } else {
                alert("Ocorreu um erro ao excluir o email. Por favor, tente novamente mais tarde.");
                // Lidar com o erro de forma apropriada
            }
        };

        // Definir o que fazer em caso de erro
        xhr.onerror = function() {
            alert("Ocorreu um erro de rede ao excluir o email. Por favor, verifique sua conexão e tente novamente.");
            // Lidar com o erro de forma apropriada
        };

        // Enviar a solicitação POST com cabeçalho personalizado para DELETE
        xhr.send();
    }
</script>

</body>
</html>
