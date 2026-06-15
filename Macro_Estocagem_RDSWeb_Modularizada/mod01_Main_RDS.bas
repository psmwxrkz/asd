Attribute VB_Name = "mod01_Main_RDS"
' =========================================================
' MÓDULO 01 - MACRO PRINCIPAL / RDSWEB
' Ponto de entrada da automação: BaixarDadosSite.
' Mantém o loop principal, leitura das fichas e estocagem.
' =========================================================

Sub BaixarDadosSite()
    Application.EnableCancelKey = xlDisabled
    Application.EnableCancelKey = xlErrorHandler
    MacroRodando = True
    RegistrarLog "INICIO", "Macro iniciada pelo botão"
    MonitorarInterrupcao ' <--- MONITORAMENTO LIGADO
    
    Set bot = New WebDriver
    Dim divs As Object, divRow As Object
    Dim tipoVeiculo As String, produtoFicha As String, valorLocal As String
    Dim ws As Worksheet
    Dim i As Long
    Dim element As Selenium.WebElement
    Dim rowIds() As String
    Dim totalRows As Long

    Call PreventSleep
    
    ' Índices para alternância
    Dim idxTEGCaminhao As Long, idxTEGVagao As Long, idxTEAGCaminhao As Long, idxTEAGVagao As Long
    idxTEGCaminhao = 0
    idxTEGVagao = 0
    idxTEAGCaminhao = 0
    idxTEAGVagao = 0
    
    Dim startTime As Date, lastUpdate As Date
    startTime = Now
    lastUpdate = Now
    
    Set ws = ThisWorkbook.Sheets("MACRO")
    
    ' Limpa apenas as células A4, B4, C4 ate F4
    ws.Range("A4:F4").ClearContents
    
    ' Exibe UserForm
    UserForm1.Show vbModeless
    StartKeepAlive 30   ' mantém vivo a cada 30s 'ALTERAR CASO CONTINUE APAGANDO A TELA (PODE SER SEGURANÇA DA CARGILL)
    On Error GoTo SairComErro
    bot.Wait 4000
    
    ' Atualiza células antes de iniciar estocagem
    Call AtualizarCelulas
    
    ' Inicia RDS
    DoEvents
    bot.Start "chrome"
    bot.Wait 700
    bot.Get "https://rdswebguaruja.la.cargill.com/"
    bot.Wait 300
    bot.Refresh
    bot.Wait 5000
    
    ' Login RDS
    If WaitForElement("xpath", "//span[@class='input-group-addon bl-none btn-xs liveListview-icon-caret-bt']", 25) Then
        bot.Wait 300
        bot.FindElementByXPath("//span[@class='input-group-addon bl-none btn-xs liveListview-icon-caret-bt']").Click
        bot.Wait 1500
    End If

    bot.Wait 1200

    If WaitForElement("xpath", "//button[normalize-space()='Entrar']", 25) Then
        bot.Wait 300
        bot.FindElementByXPath("//button[normalize-space()='Entrar']").Click
    End If

    bot.Wait 1200

    If WaitForElement("xpath", "//img[@src='/images/menu/dashboard.png']", 25) Then
        bot.Wait 300
        bot.FindElementByXPath("//img[@src='/images/menu/dashboard.png']").Click
    End If

    bot.Wait 1200

    If WaitForElement("xpath", "//a[normalize-space()='Posição de Veículos']", 25) Then
        bot.Wait 300
        bot.FindElementByXPath("//a[normalize-space()='Posição de Veículos']").Click
    End If

    SafeWait 18000, bot

' Loop principal
Do
    On Error GoTo TratarError
    KeepAliveTick bot
    DoEvents
    ' Parada automática após 5h25min (para que não permaneça logado o user do funcionario, ja que o turno é de 6h) FLUXO FINAL (END SUB)
    If DateDiff("n", startTime, Now) >= 318 Then ' Controle o tempo de execucao aqui
        RegistrarLog "FINALIZADO", "Execução concluída normalmente"
        MacroRodando = False
        StopKeepAlive
        Call RestoreSleep
        bot.Quit
        bot.Wait 1500
        Unload UserForm1
        MsgBox "MACRO ENCERRADA (Tempo de uso atingido).", vbInformation, "Aviso"
        RegistrarLog "FINALIZADO", "Execução concluída pelo tempo máximo"
        Exit Do
    End If
    
    ' Atualização das células a cada 07 minutos
    If DateDiff("n", lastUpdate, Now) >= 7 Then
        Call AtualizarCelulas
        lastUpdate = Now
    End If

    bot.Wait 3000

    ' Abrir lista de "veículos aguardando estocagem"
    Set element = bot.FindElementByXPath("(//a[@onclick=""bindIndicador('local-armazena', true)""])[1]")
    element.Click
    SafeWait 130000, bot ' aguarda um tempo para necessidades operacionais

    ' Capturar a lista de divs presentes na pagina
    Set divs = bot.FindElementsByCss("div[id^='div-row-id-']")
    If divs Is Nothing Then GoTo ContinueLoop
    If divs.Count = 0 Then GoTo ContinueLoop
    On Error GoTo 0
    ' -------- Processo de estocagem --------
    
    totalRows = divs.Count
    ReDim rowIds(1 To totalRows)
    
    For i = 1 To totalRows  'For para captar as div row
        KeepAliveTick bot
        DoEvents
        Dim idText As String
        On Error Resume Next
        idText = divs.Item(i).Attribute("id")
        On Error GoTo 0
    
        ' Fallback JS se vier vazio
        If Len(idText) = 0 Then
            On Error Resume Next
            idText = bot.ExecuteScript("return arguments[0].id;", divs.Item(i))
            On Error GoTo 0
        End If

        rowIds(i) = idText
    Next i
    
    ' 2) For do Next i para Iterar usando os IDs (O NEXT I VOLTA SEMPRE AQUI) for principal
    For i = 1 To totalRows

        On Error Resume Next
        bot.FindElementById("toolbarButtonClose").Click
        bot.Wait 3000
        If rowIds(i) = "" Then GoTo ProximaFicha  ' Caso esteja em branco, volta para o Do
        On Error GoTo 0
    
        Set divRow = Nothing
    
        ' Primeiro, tenta pegar pelo ID fixo
        On Error Resume Next
        Set divRow = bot.FindElementById(rowIds(i))
        On Error GoTo 0
    
        ' Se falhar (DOM atualizou), tenta pelo índice como fallback
        If divRow Is Nothing Then
            On Error Resume Next
            Set divRow = bot.FindElementByXPath("(//div[starts-with(@id,'div-row-id-')])[" & i & "]")
            On Error GoTo 0
            If divRow Is Nothing Then GoTo ProximaFicha
        End If
        
        ' --------------------------
        ' 1) ABRIR "Alterar Ficha" PARA LER OBSERVAÇÕES
        ' --------------------------
        On Error GoTo TratarError
        bot.Wait 1500
        ' Abrir dropdown (três pontinhos) da própria linha
        If Not TryClickChild(divRow, ".//button[@class='dropdown-toggle ficha-drop']", 8) Then GoTo ProximaFicha
        bot.Wait 1700
        If Not TryClickChild(divRow, ".//a[contains(.,'Alterar Ficha')]", 8) Then GoTo ProximaFicha
        
        SafeWait 8000, bot

        ' Ler observações
        Dim observacaoFicha As String, ladoTerminal As String
        observacaoFicha = "": ladoTerminal = ""
    
        If WaitForElement("id", "Info_Observacao", 25) Then
            bot.Wait 200
            observacaoFicha = bot.FindElementById("Info_Observacao").Attribute("value")
        Else
            observacaoFicha = ""
        End If
    
        observacaoFicha = Trim$(UCase$(observacaoFicha))
    
        ' Caso observação em BRANCO: fecha e pula
        If observacaoFicha = "" Then
            bot.FindElementById("toolbarButtonClose").Click
            bot.Wait 4000
            GoTo ProximaFicha
        End If

        ' Validar TEAG / TEG aceitando pontuação: TEG, TEAG; TEG-; (TEAG), etc.
        ladoTerminal = IdentificarTerminalObservacao(observacaoFicha)
        
        ' Se digitado errado (nem TEG nem TEAG), tratar igual ao em branco
        If ladoTerminal = "" Then
            bot.FindElementById("toolbarButtonClose").Click
            bot.Wait 4000
            GoTo ProximaFicha
        End If

        bot.Wait 1500
        ' Fechar Alterar Ficha
        bot.FindElementById("toolbarButtonClose").Click
        On Error GoTo 0
        
        bot.Wait 4000
    
        ' Re-capturar a mesma linha pelo ID (evitar stale)
        On Error Resume Next
        Set divRow = bot.FindElementById(rowIds(i))
        On Error GoTo 0
        If divRow Is Nothing Then
            On Error Resume Next
            Set divRow = bot.FindElementByXPath("(//div[starts-with(@id,'div-row-id-')])[" & i & "]")
            On Error GoTo 0
            If divRow Is Nothing Then GoTo ProximaFicha
        End If
    
        ' --------------------------
        ' 2) ABRIR A JANELA DE ESTOCAGEM (ticket-row-menu-passo) na LINHA CORRENTE
        ' --------------------------
        bot.Wait 1000
        If Not TryClickChild(divRow, ".//span[@id='ticket-row-menu-passo']", 8) Then GoTo ProximaFicha
        bot.Wait 5000
    
        On Error GoTo TratarError
        If WaitForElement("id", "TipoVeiculo_TranslationDesc", 25) Then
            bot.Wait 200
            tipoVeiculo = bot.FindElementById("TipoVeiculo_TranslationDesc").Attribute("value")
        End If
        On Error GoTo 0
        bot.Wait 500

        ' --------------------------
        ' 3) LÓGICA DE CÉLULAS E ALTERNÂNCIA (CONFORME A4:H4 E MODOS EM C5/D5/F5/G5/H5)
        ' --------------------------
        valorLocal = EscolherCelulaEstocagem(ws, ladoTerminal, tipoVeiculo, _
                                             idxTEGCaminhao, idxTEGVagao, _
                                             idxTEAGCaminhao, idxTEAGVagao)

        ' Se não houver nenhuma célula compatível com o tipo do modal, pular ficha
        If Trim$(valorLocal) = "" Then
            RegistrarLog "INFO", "Sem célula disponível/compatível para " & ladoTerminal & " - Tipo: " & tipoVeiculo & ". Ficha ignorada neste ciclo."

            On Error Resume Next
            bot.FindElementById("toolbarButtonClose").Click
            On Error GoTo TratarError

            GoTo ProximaFicha
        End If

        ' --------------------------
        ' 4) PREENCHER CAMPOS E SALVAR
        ' --------------------------
        On Error GoTo TratarError
        PreencherInputXPath bot, "//input[@id='liveListCombinacao']", ladoTerminal
        bot.Wait 3200

        PreencherInputXPath bot, "//input[@id='liveLinhasLocalArmazenagem']", ladoTerminal
        bot.Wait 3200

        PreencherInputXPath bot, "(//input[@id='liveLocalArmazenagem'])[1]", valorLocal
        bot.Wait 3200

        bot.FindElementById("toolbarButtonSave").Click ' Salvar
        SafeWait 13000, bot
        On Error GoTo 0

ProximaFicha:
    Next i

ContinueLoop:
    DoEvents
Loop

TratarError:
    RegistrarLog "ERRO", "Erro no loop principal - " & Err.Description
    bot.Refresh
    SafeWait 13000, bot
    On Error GoTo 0
    Resume ContinueLoop

SairComErro:
    RegistrarLog "ERRO", Err.Description
    MacroRodando = False
    On Error Resume Next
    ' Encerra heartbeat
    StopKeepAlive
    Call RestoreSleep
    ' Fecha o driver se estiver aberto
    If Not bot Is Nothing Then bot.Quit
    Set bot = Nothing
    ' Fecha UserForm se necessário
    Unload UserForm1
    On Error GoTo 0
    
    
End Sub

' Sub para atualizar células
