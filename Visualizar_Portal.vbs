Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Obtém o diretório deste script
scriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
serverScript = scriptDir & "\local_server.py"

' 1. Verifica se o script do servidor existe
If objFSO.FileExists(serverScript) Then
    ' Executa o servidor local em Python de forma oculta (o número 0 esconde a tela preta)
    objShell.Run "python """ & serverScript & """", 0, False
    
    ' Aguarda 1 segundo para garantir que a porta do servidor foi aberta
    WScript.Sleep 1000
    
    ' 2. Abre o site local no navegador padrão do Windows usando o servidor
    objShell.Run "http://localhost:8000"
Else
    ' Fallback caso o python server não exista
    htmlFile = scriptDir & "\index.html"
    If objFSO.FileExists(htmlFile) Then
        objShell.Run """" & htmlFile & """"
        MsgBox "O servidor local.py não foi encontrado. Abrindo o arquivo HTML estático (recursos de integração estarão indisponíveis).", vbExclamation, "Aviso"
    Else
        MsgBox "Arquivo index.html ou local_server.py não foram encontrados na pasta do projeto!", vbCritical, "Erro"
    End If
End If
