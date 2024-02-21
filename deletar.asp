<%
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Request.QueryString("userEmail") <> "" Then
        Dim userEmail
        userEmail = Request.QueryString("userEmail")
        
        Dim conn
        Set conn = Server.CreateObject("ADODB.Connection")
        conn.Open "DRIVER={MariaDB ODBC 3.1 Driver}; SERVER=localhost; DATABASE=novo; USER=root; PASSWORD=root; OPTION=3;"
        
        Dim sql
        sql = "DELETE FROM usuario WHERE email = '" & userEmail & "'"
        
        conn.Execute(sql)
        
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

        var xhr = new XMLHttpRequest();

        xhr.open("POST", "deletar.asp?userEmail=" + encodeURIComponent(userEmail), true);
        xhr.setRequestHeader("X-HTTP-Method-Override", "DELETE");

        xhr.onload = function() {
            if (xhr.status >= 200 && xhr.status < 300) {
                alert("Email excluído com sucesso!");

            } else {
                alert("Ocorreu um erro ao excluir o email. Por favor, tente novamente mais tarde.");

            }
        };

        xhr.onerror = function() {
            alert("Ocorreu um erro de rede ao excluir o email. Por favor, verifique sua conexão e tente novamente.");
        };

        xhr.send();
    }
</script>

</body>
</html>
