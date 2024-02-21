<%
If Request.ServerVariables("REQUEST_METHOD") = "GET" Then
    Dim conn, rs
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "DRIVER={MariaDB ODBC 3.1 Driver}; SERVER=localhost; DATABASE=novo; USER=root; PASSWORD=root; OPTION=3;"
    
    Dim sql
    sql = "SELECT email, PASSWORD FROM usuario"
    
    Set rs = conn.Execute(sql)
    
    Dim dados
    dados = "<h2>Dados do Banco de Dados:</h2><ul>"
    
    Do Until rs.EOF
        dados = dados & "<li>Email: " & rs("email") & ", Senha: " & rs("PASSWORD") & "</li>"
        rs.MoveNext
    Loop
    
    dados = dados & "</ul>"
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    Response.Write dados
End If
%>

<!DOCTYPE html>
<html>
<head>
    <title>Exemplo de Busca de Dados</title>
    <meta http-equiv="Content-Type" content="text/html;charset=UTF-8">
    <script>
        function buscarDados() {
            var xhr = new XMLHttpRequest();
            
            xhr.onreadystatechange = function() {
                if (xhr.readyState == 4 && xhr.status == 200) {
                    document.getElementById("dadosBanco").innerHTML = xhr.responseText;
                }
            };
            
            xhr.open("GET", "buscarDados.asp", true);
            
            xhr.send();
        }
    </script>
</head>
<body>
    <h1>Exemplo de Busca de Dados</h1>
    
    <form>
        <input type="button" value="Buscar Dados" onclick="buscarDados()">
    </form>
    
    <div id="dadosBanco"></div>
</body>
</html>
