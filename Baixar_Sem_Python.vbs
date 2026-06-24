' VBScript para Baixar Relatorio do Interview sem precisar de Python
Option Explicit

Dim objShell, objFSO, xml, html, window
Dim usuario, senha, reportId, loginUrl, dataUrl, payload
Dim responseText, cookieHeader, cookie, colunas, valores
Dim objExcel, objWorkbook, objWorksheet, recordCount
Dim diretorioDestino, dataAtual, nomeArquivo, caminhoCompleto
Dim opcao

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
opcao = InputBox("Escolha o relatorio que deseja baixar (digite de 1 a 5):" & vbCrLf & vbCrLf & _
                 "1 - Colaborador Senior (ID 132)" & vbCrLf & _
                 "2 - Colaboradores Grupo - Geral (ID 310)" & vbCrLf & _
                 "3 - Colaboradores Senior - MS (ID 316)" & vbCrLf & _
                 "4 - Colaboradores com Posto de Trabalho (ID 372)" & vbCrLf & _
                 "5 - Outro ID (Digitar ID personalizado)", "Selecionar Relatorio", "1")

If opcao = "" Then WScript.Quit

Select Case opcao
    Case "1"
        reportId = "132"
    Case "2"
        reportId = "310"
    Case "3"
        reportId = "316"
    Case "4"
        reportId = "372"
    Case "5"
        reportId = InputBox("Digite o ID do relatorio personalizado:", "ID Personalizado")
        If reportId = "" Then WScript.Quit
    Case Else
        ' Se o usuario digitar o ID direto (ex: 132) em vez do numero do menu
        If IsNumeric(opcao) Then
            reportId = opcao
        Else
            MsgBox "Opcao invalida!", vbExclamation, "Erro"
            WScript.Quit
        End If
End Select

' 3. Criar objeto HTTP com suporte a cookies e Headers de Navegador Real (Chrome)
Set xml = CreateObject("MSXML2.ServerXMLHTTP.6.0")

' Configurar URLs
loginUrl = "http://reportview.cocal.com.br/login.php"
dataUrl = "http://reportview.cocal.com.br/relatorio/dados.php?id=" & reportId

' Codificar dados para evitar problemas com caracteres especiais (como @, #, $, etc.)
payload = "usuario=" & window.encodeURIComponent(usuario) & "&senha=" & window.encodeURIComponent(senha)

' 4. Efetuar Login (com cabecalho disfarce para simular Chrome)
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

' 5. Fazer a requisicao dos dados do Relatorio
On Error Resume Next
xml.open "GET", dataUrl, False
xml.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
If cookie <> "" Then
    xml.setRequestHeader "Cookie", cookie
End If
xml.send

If Err.Number <> 0 Then
    MsgBox "Erro ao baixar os dados do relatorio: " & Err.Description, vbCritical, "Erro de Rede"
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
    MsgBox "Falha de login ou acesso negado. Verifique seu usuario e senha do Interview.", vbExclamation, "Falha de Autenticacao"
    WScript.Quit
End If

' 6. Iniciar Excel e Criar Planilha
On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Nao foi possivel abrir o Microsoft Excel. Verifique se ele esta instalado.", vbCritical, "Erro no Excel"
    WScript.Quit
End If
On Error GoTo 0

objExcel.Visible = False
Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Sheets(1)

' Popular o Excel usando o codigo JavaScript otimizado
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

' Criar pasta caso nao exista
If Not objFSO.FolderExists(diretorioDestino) Then
    On Error Resume Next
    CreateFolderChain objFSO, diretorioDestino
    If Err.Number <> 0 Then
        ' Se falhar por falta de permissao, salva na mesma pasta do script
        diretorioDestino = objFSO.GetParentFolderName(WScript.ScriptFullName)
    End If
    On Error GoTo 0
End If

dataAtual = Replace(FormatDateTime(Date, 2), "/", "-")
nomeArquivo = "Colaboradores_ID_" & reportId & "_" & dataAtual & ".xlsx"
caminhoCompleto = objFSO.BuildPath(diretorioDestino, nomeArquivo)

' Excluir arquivo antigo caso ja exista para evitar mensagens de substituicao
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

MsgBox "Integracao Concluida!" & vbCrLf & vbCrLf & _
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
