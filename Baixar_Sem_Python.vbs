' VBScript para Baixar Relatorio do Interview sem precisar de Python
Option Explicit

Dim objShell, objFSO, xml, html, window
Dim usuario, senha, reportId, loginUrl, dataUrl, payload
Dim responseText, cookieHeader, cookie
Dim objExcel, objWorkbook, objWorksheet, recordCount
Dim defaultDestino, diretorioDestino, dataAtual, nomeArquivo, caminhoCompleto
Dim opcao, fileBaseName, i, totalSuccess
Dim inBatch, shellApp, folderObj

Dim reportListId, reportListFile, reportListName
reportListId = Array("132", "310", "316", "372")
reportListFile = Array("Colaborador_Senior", "Colaboradores_Grupo_Geral", "Colaboradores_Senior_MS", "Colaboradores_Posto_Trabalho")
reportListName = Array("Colaborador Senior", "Colaboradores Grupo - Geral", "Colaboradores Senior - MS", "Colaboradores com Posto de Trabalho")

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' 1. Criar HTMLFile temporario para usar recursos de JavaScript (JSON.parse e encodeURIComponent)
Set html = CreateObject("htmlfile")
html.write "<meta http-equiv=""x-ua-compatible"" content=""IE=9"">"
Set window = html.parentWindow

' Inserir a funcao JavaScript que vai popular o Excel diretamente (super rapido e sem bugs de conversao de array)
Dim jsCode
jsCode = "function populateExcel(sheet, jsonText) {" & _
         "    var data = JSON.parse(jsonText);" & _
         "    var colunas = data.colunas;" & _
         "    var valores = data.valores;" & _
         "    for (var c = 0; c < colunas.length; c++) {" & _
         "        sheet.Cells(1, c + 1).Value = colunas[c].title || colunas[c];" & _
         "    }" & _
         "    for (var r = 0; r < valores.length; r++) {" & _
         "        var row = valores[r];" & _
         "        for (var c = 0; c < row.length; c++) {" & _
         "            var val = row[c];" & _
         "            if (val !== null && val !== undefined) {" & _
         "                if (!isNaN(val) && val !== '' && val.indexOf('0') !== 0) {" & _
         "                    sheet.Cells(r + 2, c + 1).Value = Number(val);" & _
         "                } else {" & _
         "                    sheet.Cells(r + 2, c + 1).Value = val;" & _
         "                }" & _
         "            }" & _
         "        }" & _
         "    }" & _
         "    return valores.length;" & _
         "}"

window.execScript jsCode, "JScript"

' 2. Solicitar credenciais ao usuario de forma amigavel
usuario = InputBox("Digite seu usuario do Interview:", "Login Interview - Cocal", "maicon.bahls")
If usuario = "" Then WScript.Quit

senha = InputBox("Digite sua senha do Interview:", "Login Interview - Cocal")
If senha = "" Then WScript.Quit

' Menu de selecao de relatorios para evitar digitacao
opcao = InputBox("Escolha o relatorio que deseja baixar (digite de 1 a 6):" & vbCrLf & vbCrLf & _
                 "1 - Colaborador Senior (ID 132)" & vbCrLf & _
                 "2 - Colaboradores Grupo - Geral (ID 310)" & vbCrLf & _
                 "3 - Colaboradores Senior - MS (ID 316)" & vbCrLf & _
                 "4 - Colaboradores com Posto de Trabalho (ID 372)" & vbCrLf & _
                 "5 - Outro ID (Digitar ID personalizado)" & vbCrLf & _
                 "6 - Baixar TODOS os 4 relatorios acima em LOTE", "Selecionar Relatorio", "6")

If opcao = "" Then WScript.Quit

' Definir se e download em lote ou individual
inBatch = (opcao = "6")

If Not inBatch Then
    Select Case opcao
        Case "1"
            reportId = "132"
            fileBaseName = "Colaborador_Senior"
        Case "2"
            reportId = "310"
            fileBaseName = "Colaboradores_Grupo_Geral"
        Case "3"
            reportId = "316"
            fileBaseName = "Colaboradores_Senior_MS"
        Case "4"
            reportId = "372"
            fileBaseName = "Colaboradores_Posto_Trabalho"
        Case "5"
            reportId = InputBox("Digite o ID do relatorio personalizado:", "ID Personalizado")
            If reportId = "" Then WScript.Quit
            fileBaseName = "Relatorio_ID_" & reportId
        Case Else
            ' Se o usuario digitar o ID direto (ex: 132) em vez do numero do menu
            If IsNumeric(opcao) Then
                reportId = opcao
                fileBaseName = "Relatorio_ID_" & reportId
            Else
                MsgBox "Opcao invalida!", vbExclamation, "Erro"
                WScript.Quit
            End If
    End Select
End If

' 3. Definir pasta de destino interativamente
defaultDestino = "C:\Users\maicon.bahls\Cocal\Recursos Humanos - 09_Selects\AUTOMAÇÃO"

' Se a pasta padrao nao existir, tenta cria-la
If Not objFSO.FolderExists(defaultDestino) Then
    On Error Resume Next
    CreateFolderChain objFSO, defaultDestino
    If Err.Number <> 0 Then
        ' Caso falhe, usa a pasta onde o script esta
        defaultDestino = objFSO.GetParentFolderName(WScript.ScriptFullName)
    End If
    On Error GoTo 0
End If

' Abre a caixa de selecao de pasta nativa do Windows
On Error Resume Next
Set shellApp = CreateObject("Shell.Application")
' &H0010 = BIF_RETURNONLYFSDIRS (somente pastas reais), &H0040 = BIF_USENEWUI (visual moderno com botao de criar pasta)
Set folderObj = shellApp.BrowseForFolder(0, "Selecione a pasta onde os relatorios serão salvos:", &H0010 + &H0040, defaultDestino)

If Err.Number = 0 And Not folderObj Is Nothing Then
    diretorioDestino = folderObj.Self.Path
Else
    ' Se o usuario cancelar ou fechar a janela, usa a pasta padrao
    diretorioDestino = defaultDestino
End If
On Error GoTo 0

' 4. Criar objeto HTTP com suporte a cookies e Headers de Navegador Real (Chrome)
Set xml = CreateObject("MSXML2.ServerXMLHTTP.6.0")

' Configurar URL de Login
loginUrl = "http://reportview.cocal.com.br/login.php"

' Codificar dados para evitar problemas com caracteres especiais (como @, #, $, etc.)
payload = "usuario=" & window.encodeURIComponent(usuario) & "&senha=" & window.encodeURIComponent(senha)

' 5. Efetuar Login (com cabecalho disfarce para simular Chrome)
On Error Resume Next
xml.open "POST", loginUrl, False
xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
xml.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
xml.send payload

If Err.Number <> 0 Then
    MsgBox "Erro de conexao ao servidor do Interview: " & Err.Description, vbCritical, "Erro de Rede"
    WScript.Quit
End If
On Error GoTo 0

' Extrair Cookie de Sessao
cookie = ""
cookieHeader = xml.getResponseHeader("Set-Cookie")
If cookieHeader <> "" Then
    If InStr(cookieHeader, ";") > 0 Then
        cookie = Split(cookieHeader, ";")(0)
    Else
        cookie = cookieHeader
    End If
End If

' Testar se login falhou
responseText = xml.responseText
If InStr(responseText, "Login - Interview") > 0 Or InStr(responseText, "<!doctype html") > 0 Then
    MsgBox "Falha de login ou acesso negado. Verifique seu usuario e senha do Interview.", vbExclamation, "Falha de Autenticacao"
    WScript.Quit
End If

' 6. Iniciar Excel
On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Nao foi possivel abrir o Microsoft Excel. Verifique se ele esta instalado.", vbCritical, "Erro no Excel"
    WScript.Quit
End If
On Error GoTo 0

objExcel.Visible = False
dataAtual = Replace(FormatDateTime(Date, 2), "/", "-")

' 7. Executar Downloads e Salvar
If inBatch Then
    totalSuccess = 0
    For i = 0 To UBound(reportListId)
        reportId = reportListId(i)
        dataUrl = "http://reportview.cocal.com.br/relatorio/dados.php?id=" & reportId
        
        On Error Resume Next
        xml.open "GET", dataUrl, False
        xml.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        If cookie <> "" Then
            xml.setRequestHeader "Cookie", cookie
        End If
        xml.send
        
        If Err.Number = 0 Then
            responseText = xml.responseText
            If Left(responseText, 1) = ChrW(&hFEFF) Then
                responseText = Mid(responseText, 2)
            End If
            
            ' Verificar se retornou JSON valido
            If InStr(responseText, "Login - Interview") = 0 And InStr(responseText, "<!doctype html") = 0 Then
                Set objWorkbook = objExcel.Workbooks.Add()
                Set objWorksheet = objWorkbook.Sheets(1)
                
                recordCount = window.populateExcel(objWorksheet, responseText)
                objWorksheet.UsedRange.Columns.AutoFit
                
                nomeArquivo = reportListFile(i) & "_" & dataAtual & ".xlsx"
                caminhoCompleto = objFSO.BuildPath(diretorioDestino, nomeArquivo)
                
                ' Excluir arquivo antigo caso ja exista
                If objFSO.FileExists(caminhoCompleto) Then
                    objFSO.DeleteFile caminhoCompleto, True
                End If
                
                objWorkbook.SaveAs caminhoCompleto
                objWorkbook.Close True
                totalSuccess = totalSuccess + 1
            End If
        End If
        On Error GoTo 0
        WScript.Sleep 500 ' Pausa curta para nao sobrecarregar o servidor
    Next
    
    objExcel.Quit
    
    MsgBox "Processo de Lote Concluido!" & vbCrLf & vbCrLf & _
           "Relatorios baixados com sucesso: " & totalSuccess & " de 4." & vbCrLf & _
           "Arquivos salvos em: " & diretorioDestino, vbInformation, "Lote Finalizado"
Else
    ' Download Individual
    dataUrl = "http://reportview.cocal.com.br/relatorio/dados.php?id=" & reportId
    
    On Error Resume Next
    xml.open "GET", dataUrl, False
    xml.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    If cookie <> "" Then
        xml.setRequestHeader "Cookie", cookie
    End If
    xml.send
    
    If Err.Number <> 0 Then
        objExcel.Quit
        MsgBox "Erro ao baixar os dados do relatorio: " & Err.Description, vbCritical, "Erro de Rede"
        WScript.Quit
    End If
    On Error GoTo 0
    
    responseText = xml.responseText
    If Left(responseText, 1) = ChrW(&hFEFF) Then
        responseText = Mid(responseText, 2)
    End If
    
    Set objWorkbook = objExcel.Workbooks.Add()
    Set objWorksheet = objWorkbook.Sheets(1)
    
    On Error Resume Next
    recordCount = window.populateExcel(objWorksheet, responseText)
    If Err.Number <> 0 Then
        objWorkbook.Close False
        objExcel.Quit
        MsgBox "Falha ao processar os dados recebidos: " & Err.Description, vbCritical, "Erro no JSON"
        WScript.Quit
    End If
    On Error GoTo 0
    
    objWorksheet.UsedRange.Columns.AutoFit
    
    nomeArquivo = fileBaseName & "_" & dataAtual & ".xlsx"
    caminhoCompleto = objFSO.BuildPath(diretorioDestino, nomeArquivo)
    
    If objFSO.FileExists(caminhoCompleto) Then
        On Error Resume Next
        objFSO.DeleteFile caminhoCompleto, True
        On Error GoTo 0
    End If
    
    On Error Resume Next
    objWorkbook.SaveAs caminhoCompleto
    objWorkbook.Close True
    objExcel.Quit
    If Err.Number <> 0 Then
        MsgBox "Erro ao salvar a planilha em: " & caminhoCompleto & vbCrLf & Err.Description, vbCritical, "Erro ao Salvar"
        WScript.Quit
    End If
    On Error GoTo 0
    
    MsgBox "Integracao Concluida!" & vbCrLf & vbCrLf & _
           "Registros baixados: " & recordCount & vbCrLf & _
           "Salvo em: " & caminhoCompleto, vbInformation, "Sucesso"
End If

' Função recursiva para criar a árvore de pastas
Sub CreateFolderChain(fso, path)
    Dim parent
    parent = fso.GetParentFolderName(path)
    If Not fso.FolderExists(parent) And parent <> "" Then
        CreateFolderChain fso, parent
    End If
    If Not fso.FolderExists(path) Then
        fso.CreateFolder path
    End If
End Sub
