<%
' Verifica se a requisição é um GET
If Request.ServerVariables("REQUEST_METHOD") = "GET" Then
    ' Conecta ao banco de dados MariaDB
    Dim conn, rs
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "DRIVER={MariaDB ODBC 3.1 Driver}; SERVER=localhost; DATABASE=novo; USER=root; PASSWORD=root; OPTION=3;"
    
    ' Query SQL para buscar dados
    Dim sql
    sql = "SELECT email, PASSWORD FROM usuario"
    
    ' Executa a query SQL
    Set rs = conn.Execute(sql)
    
    ' Cria uma string para armazenar os dados do banco de dados
    Dim dados
    dados = "<h2>Dados do Banco de Dados:</h2><ul>"
    
    ' Percorre os registros do banco de dados e adiciona-os à string
    Do Until rs.EOF
        dados = dados & "<li>Email: " & rs("email") & ", Senha: " & rs("PASSWORD") & "</li>"
        rs.MoveNext
    Loop
    
    dados = dados & "</ul>"
    
    ' Fecha a conexão e o recordset
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    ' Retorna os dados como uma resposta para a requisição AJAX
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
            // Cria uma instância do objeto XMLHttpRequest
            var xhr = new XMLHttpRequest();
            
            // Define a função a ser chamada quando a resposta for recebida
            xhr.onreadystatechange = function() {
                if (xhr.readyState == 4 && xhr.status == 200) {
                    // Se a requisição for bem-sucedida, atualiza a parte do documento com os dados recebidos
                    document.getElementById("dadosBanco").innerHTML = xhr.responseText;
                }
            };
            
            // Abre a requisição GET para o servidor ASP Classic
            xhr.open("GET", "buscarDados.asp", true);
            
            // Envia a requisição
            xhr.send();
        }
    </script>
</head>
<body>
    <h1>Exemplo de Busca de Dados</h1>
    
    <form>
        <!-- Chama a função buscarDados() quando o botão for clicado -->
        <input type="button" value="Buscar Dados" onclick="buscarDados()">
    </form>
    
    <!-- Aqui serão exibidos os dados do banco de dados -->
    <div id="dadosBanco"></div>
</body>
</html>
