' VBScript para Baixar Relatório do Interview sem precisar de Python
Option Explicit

Dim objShell, objFSO, xml, html, window
Dim usuario, senha, reportId, loginUrl, dataUrl, payload
Dim responseText, cookieHeader, cookie, colunas, valores
Dim objExcel, objWorkbook, objWorksheet, recordCount
Dim diretorioDestino, dataAtual, nomeArquivo, caminhoCompleto

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' 1. Criar HTMLFile temporário para usar recursos de JavaScript (JSON.parse e encodeURIComponent)
Set html = CreateObject("htmlfile")
html.write "<meta http-equiv=""x-ua-compatible"" content=""IE=9"">"
Set window = html.parentWindow

' Inserir a função JavaScript que vai popular o Excel diretamente (super rápido e sem bugs de conversão de array)
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

' 2. Solicitar credenciais ao usuário de forma amigável
usuario = InputBox("Digite seu usuário do Interview:", "Login Interview - Cocal", "maicon.bahls")
If usuario = "" Then WScript.Quit

senha = InputBox("Digite sua senha do Interview:", "Login Interview - Cocal")
If senha = "" Then WScript.Quit

reportId = InputBox("Digite o ID do relatório que deseja baixar:" & vbCrLf & "(132 = Colaborador Senior, 372 = Posto de Trabalho, etc.)", "Selecionar Relatório", "132")
If reportId = "" Then WScript.Quit

' 3. Criar objeto HTTP com suporte a cookies e Headers de Navegador Real (Chrome)
Set xml = CreateObject("MSXML2.ServerXMLHTTP.6.0")

' Configurar URLs
loginUrl = "http://reportview.cocal.com.br/login.php"
dataUrl = "http://reportview.cocal.com.br/relatorio/dados.php?id=" & reportId

' Codificar dados para evitar problemas com caracteres especiais (como @, #, $, etc.)
payload = "usuario=" & window.encodeURIComponent(usuario) & "&senha=" & window.encodeURIComponent(senha)

' 4. Efetuar Login (com cabeçalho disfarçado para simular Chrome)
On Error Resume Next
xml.open "POST", loginUrl, False
xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
xml.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
xml.send payload

If Err.Number <> 0 Then
    MsgBox "Erro de conexão ao servidor do Interview: " & Err.Description, vbCritical, "Erro de Rede"
    WScript.Quit
End If
On Error GoTo 0

' Extrair Cookie de Sessão
cookie = ""
cookieHeader = xml.getResponseHeader("Set-Cookie")
If cookieHeader <> "" Then
    If InStr(cookieHeader, ";") > 0 Then
        cookie = Split(cookieHeader, ";")(0)
    Else
        cookie = cookieHeader
    End If
End If

' 5. Fazer a requisição dos dados do Relatório
On Error Resume Next
xml.open "GET", dataUrl, False
xml.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
If cookie <> "" Then
    xml.setRequestHeader "Cookie", cookie
End If
xml.send

If Err.Number <> 0 Then
    MsgBox "Erro ao baixar os dados do relatório: " & Err.Description, vbCritical, "Erro de Rede"
    WScript.Quit
End If
On Error GoTo 0

responseText = xml.responseText

' Remover UTF-8 BOM se presente
If Left(responseText, 1) = ChrW(&hFEFF) Then
    responseText = Mid(responseText, 2)
End If

' Verificar se fomos redirecionados para a tela de login
If InStr(responseText, "Login - Interview") > 0 Or InStr(responseText, "<!doctype html") > 0 Then
    MsgBox "Falha de login ou acesso negado. Verifique seu usuário e senha do Interview.", vbExclamation, "Falha de Autenticação"
    WScript.Quit
End If

' 6. Iniciar Excel e Criar Planilha
On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Não foi possível abrir o Microsoft Excel. Verifique se ele está instalado.", vbCritical, "Erro no Excel"
    WScript.Quit
End If
On Error GoTo 0

objExcel.Visible = False
Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Sheets(1)

' Popular o Excel usando o código JavaScript otimizado
On Error Resume Next
recordCount = window.populateExcel(objWorksheet, responseText)
If Err.Number <> 0 Then
    objWorkbook.Close False
    objExcel.Quit
    MsgBox "Falha ao processar os dados recebidos: " & Err.Description, vbCritical, "Erro no JSON"
    WScript.Quit
End If
On Error GoTo 0

' Auto-ajustar a largura das colunas
objWorksheet.UsedRange.Columns.AutoFit

' 7. Definir pasta de destino e salvar
diretorioDestino = "C:\Users\maicon.bahls\Cocal\Recursos Humanos - 09_Selects\AUTOMAÇÃO"

' Criar pasta caso não exista
If Not objFSO.FolderExists(diretorioDestino) Then
    On Error Resume Next
    CreateFolderChain objFSO, diretorioDestino
    If Err.Number <> 0 Then
        ' Se falhar por falta de permissão, salva na mesma pasta do script
        diretorioDestino = objFSO.GetParentFolderName(WScript.ScriptFullName)
    End If
    On Error GoTo 0
End If

dataAtual = Replace(FormatDateTime(Date, 2), "/", "-")
nomeArquivo = "Colaboradores_ID_" & reportId & "_" & dataAtual & ".xlsx"
caminhoCompleto = objFSO.BuildPath(diretorioDestino, nomeArquivo)

' Excluir arquivo antigo caso já exista para evitar mensagens de substituição
If objFSO.FileExists(caminhoCompleto) Then
    On Error Resume Next
    objFSO.DeleteFile caminhoCompleto, True
    On Error GoTo 0
End If

' Salvar e fechar
On Error Resume Next
objWorkbook.SaveAs caminhoCompleto
objWorkbook.Close True
objExcel.Quit
If Err.Number <> 0 Then
    MsgBox "Erro ao salvar a planilha em: " & caminhoCompleto & vbCrLf & Err.Description, vbCritical, "Erro ao Salvar"
    WScript.Quit
End If
On Error GoTo 0

MsgBox "Integração Concluída!" & vbCrLf & vbCrLf & _
       "Registros baixados: " & recordCount & vbCrLf & _
       "Salvo em: " & caminhoCompleto, vbInformation, "Sucesso"

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
