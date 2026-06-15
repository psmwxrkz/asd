Attribute VB_Name = "mod02_Atualizar_Celulas"
' =========================================================
' MÓDULO 02 - SITE DAS CÉLULAS / NOTEETEG
' Atualiza A4:H4 com as células disponíveis e grava em
' C5/D5/F5/G5/H5 se cada moega está como caminhão ou vagão.
' =========================================================

Sub AtualizarCelulas()
    Dim botUpdate As WebDriver
    Dim ws As Worksheet

    On Error GoTo TrataErro

    Set ws = ThisWorkbook.Sheets("MACRO")
    Set botUpdate = New WebDriver

    botUpdate.Start "chrome"
    botUpdate.Get "https://noteeteg.vercel.app"
    botUpdate.Wait 10000

    botUpdate.FindElementById("email").SendKeys "cco@tegporto.com.br"
    botUpdate.Wait 500
    botUpdate.FindElementById("password").SendKeys "qwerty12"
    botUpdate.Wait 500
    botUpdate.FindElementByXPath("(//button[normalize-space()='Entrar'])[1]").Click

    If WaitForElementDriver(botUpdate, "xpath", "(//button[normalize-space()='Estocagem'])[1]", 35) Then
        botUpdate.Wait 300
        botUpdate.FindElementByXPath("(//button[normalize-space()='Estocagem'])[1]").Click
        botUpdate.Wait 1500
    Else
        RegistrarLog "ERRO", "AtualizarCelulas: botão Estocagem não encontrado no site noteeteg."
        GoTo Finalizar
    End If

    ' =========================================================
    ' Linha 4 = célula/local de armazenagem vindo do noteeteg
    ' Linha 5 = modo da moega: CAMINHÃO ou VAGÃO
    '
    ' A4 = Tombador 06          | A5 = CAMINHÃO fixo
    ' B4 = Tombador 07          | B5 = CAMINHÃO fixo
    ' C4 = Moega 01             | C5 = Dropdown Descarga Caminhão/Vagão
    ' D4 = Moega 02             | D5 = Dropdown Descarga Caminhão/Vagão
    ' E4 = Tombador 05          | E5 = CAMINHÃO fixo
    ' F4 = Moega 03             | F5 = Dropdown Descarga Caminhão/Vagão
    ' G4 = Moega 04             | G5 = Dropdown Descarga Caminhão/Vagão
    ' H4 = Moega 05             | H5 = Dropdown Descarga Caminhão/Vagão
    ' =========================================================

    ws.Range("A4").Value = LerValorInputPorIds(botUpdate, Array("teg-road"))
    ws.Range("B4").Value = LerValorInputPorIds(botUpdate, Array("teg-road-tombador"))
    ws.Range("C4").Value = LerValorInputPorIds(botUpdate, Array("teg-railway-moega-01"))
    ws.Range("D4").Value = LerValorInputPorIds(botUpdate, Array("teg-railway-moega-02"))

    ws.Range("E4").Value = LerValorInputPorIds(botUpdate, Array("teag-road"))
    ws.Range("F4").Value = LerValorInputPorIds(botUpdate, Array("teag-railway-moega-03", "teag-railway"))
    ws.Range("G4").Value = LerValorInputPorIds(botUpdate, Array("teag-railway-moega-04"))
    ws.Range("H4").Value = LerValorInputPorIds(botUpdate, Array("teag-railway-moega-05"))

    ' Tombadores são sempre caminhão.
    ws.Range("A5").Value = "CAMINHÃO"
    ws.Range("B5").Value = "CAMINHÃO"
    ws.Range("E5").Value = "CAMINHÃO"

    ' Dropdowns das moegas.
    ' IMPORTANTE:
    ' Caso o ID real do dropdown no noteeteg seja diferente, ajuste apenas os arrays abaixo.
    ' A função tenta localizar por id, parte do id, name e data-testid.
    ' Se não encontrar o dropdown, assume VAGÃO para manter a lógica antiga e evitar mandar vagão para tombador.
    ws.Range("C5").Value = LerModoMoega(botUpdate, Array("teg-railway-moega-01-mode", "teg-railway-moega-01-tipo", "teg-railway-moega-01-descarga", "teg-railway-moega-01"), "VAGÃO")
    ws.Range("D5").Value = LerModoMoega(botUpdate, Array("teg-railway-moega-02-mode", "teg-railway-moega-02-tipo", "teg-railway-moega-02-descarga", "teg-railway-moega-02"), "VAGÃO")
    ws.Range("F5").Value = LerModoMoega(botUpdate, Array("teag-railway-moega-03-mode", "teag-railway-moega-03-tipo", "teag-railway-moega-03-descarga", "teag-railway-moega-03", "teag-railway"), "VAGÃO")
    ws.Range("G5").Value = LerModoMoega(botUpdate, Array("teag-railway-moega-04-mode", "teag-railway-moega-04-tipo", "teag-railway-moega-04-descarga", "teag-railway-moega-04"), "VAGÃO")
    ws.Range("H5").Value = LerModoMoega(botUpdate, Array("teag-railway-moega-05-mode", "teag-railway-moega-05-tipo", "teag-railway-moega-05-descarga", "teag-railway-moega-05"), "VAGÃO")

    ws.Range("A5:H5").Font.Color = RGB(120, 120, 120)

    RegistrarLog "INFO", "Células atualizadas: A4:H4 / modos C5,D5,F5,G5,H5 - " & Now
    Debug.Print "Células atualizadas às " & Now

Finalizar:
    On Error Resume Next
    If Not botUpdate Is Nothing Then botUpdate.Quit
    Set botUpdate = Nothing
    On Error GoTo 0
    Exit Sub

TrataErro:
    RegistrarLog "ERRO", "Erro em AtualizarCelulas - " & Err.Number & " - " & Err.Description
    Resume Finalizar
End Sub
