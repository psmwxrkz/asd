Private bot As WebDriver

' =========================================================
' HELPERS GERAIS - Selenium / Waits / Textos
' =========================================================

Private Function SegundosDecorridos(ByVal inicio As Single) As Double
    ' Evita problema quando a macro atravessa a meia-noite.
    If Timer >= inicio Then
        SegundosDecorridos = Timer - inicio
    Else
        SegundosDecorridos = (86400 - inicio) + Timer
    End If
End Function

Private Sub WaitDriver(ByVal driver As Object, ByVal ms As Long)
    On Error Resume Next
    If Not driver Is Nothing Then
        driver.Wait ms
    Else
        DoEvents
    End If
    On Error GoTo 0
End Sub

Private Function ElementoVisivelHabilitado(ByVal el As Object) As Boolean
    On Error Resume Next

    ElementoVisivelHabilitado = False
    If el Is Nothing Then Exit Function

    ' SeleniumBasic pode variar entre IsDisplayed/IsEnabled e Displayed/Enabled.
    ElementoVisivelHabilitado = (el.IsDisplayed And el.IsEnabled)

    If Err.Number <> 0 Then
        Err.Clear
        ElementoVisivelHabilitado = (el.Displayed And el.Enabled)
    End If

    On Error GoTo 0
End Function

' Aguarda que um elemento esteja visível e habilitado
Private Function WaitClickable(ByVal el As Object, Optional ByVal timeoutSec As Long = 6) As Boolean
    Dim t As Single
    t = Timer

    Do
        If ElementoVisivelHabilitado(el) Then
            WaitClickable = True
            Exit Function
        End If

        WaitDriver bot, 120
        If SegundosDecorridos(t) >= timeoutSec Then Exit Do
    Loop
End Function

' Tenta clicar em um filho relativo ao parent; usa click padrão e fallback JS
Private Function TryClickChild(ByVal parent As Object, ByVal xpath As String, Optional ByVal timeoutSec As Long = 8) As Boolean
    Dim t As Single
    Dim btn As Object
    t = Timer

    Do
        On Error Resume Next
        Set btn = parent.FindElementByXPath(xpath)
        On Error GoTo 0

        If Not btn Is Nothing Then
            If WaitClickable(btn, 1) Then
                On Error Resume Next
                btn.Click
                If Err.Number = 0 Then
                    On Error GoTo 0
                    TryClickChild = True
                    Exit Function
                End If

                Err.Clear
                bot.ExecuteScript "arguments[0].click();", btn
                If Err.Number = 0 Then
                    On Error GoTo 0
                    TryClickChild = True
                    Exit Function
                End If
                On Error GoTo 0
            End If
        End If

        WaitDriver bot, 150
        If SegundosDecorridos(t) >= timeoutSec Then Exit Do
    Loop
End Function

Function WaitForElementDriver(ByVal driver As Object, ByVal ByType As String, ByVal Identifier As String, Optional ByVal Timeout As Integer = 10) As Boolean
    Dim startTime As Single
    Dim el As Object

    startTime = Timer

    Do While SegundosDecorridos(startTime) < Timeout
        Set el = Nothing

        On Error Resume Next
        Select Case LCase$(ByType)
            Case "id"
                Set el = driver.FindElementById(Identifier)
            Case "xpath"
                Set el = driver.FindElementByXPath(Identifier)
            Case "css"
                Set el = driver.FindElementByCss(Identifier)
        End Select
        On Error GoTo 0

        If Not el Is Nothing Then
            If ElementoVisivelHabilitado(el) Then
                WaitForElementDriver = True
                Exit Function
            End If
        End If

        WaitDriver driver, 200
        DoEvents
    Loop

    WaitForElementDriver = False
End Function

' Mantém compatibilidade com as chamadas antigas da macro principal.
Function WaitForElement(ByVal ByType As String, ByVal Identifier As String, Optional ByVal Timeout As Integer = 10) As Boolean
    WaitForElement = WaitForElementDriver(bot, ByType, Identifier, Timeout)
End Function

Private Function TextoSemAcento(ByVal valor As String) As String
    valor = UCase$(Trim$(valor))

    valor = Replace(valor, "Á", "A")
    valor = Replace(valor, "À", "A")
    valor = Replace(valor, "Â", "A")
    valor = Replace(valor, "Ã", "A")
    valor = Replace(valor, "Ä", "A")
    valor = Replace(valor, "É", "E")
    valor = Replace(valor, "È", "E")
    valor = Replace(valor, "Ê", "E")
    valor = Replace(valor, "Ë", "E")
    valor = Replace(valor, "Í", "I")
    valor = Replace(valor, "Ì", "I")
    valor = Replace(valor, "Î", "I")
    valor = Replace(valor, "Ï", "I")
    valor = Replace(valor, "Ó", "O")
    valor = Replace(valor, "Ò", "O")
    valor = Replace(valor, "Ô", "O")
    valor = Replace(valor, "Õ", "O")
    valor = Replace(valor, "Ö", "O")
    valor = Replace(valor, "Ú", "U")
    valor = Replace(valor, "Ù", "U")
    valor = Replace(valor, "Û", "U")
    valor = Replace(valor, "Ü", "U")
    valor = Replace(valor, "Ç", "C")

    TextoSemAcento = valor
End Function

Private Function EhCaminhao(ByVal tipoVeiculo As String) As Boolean
    Dim t As String
    t = TextoSemAcento(tipoVeiculo)
    EhCaminhao = (InStr(1, t, "CAMINH", vbTextCompare) > 0)
End Function

Private Function IdentificarTerminalObservacao(ByVal observacao As String) As String
    Dim t As String

    t = " " & TextoSemAcento(observacao) & " "

    ' Troca separadores comuns por espaço para reconhecer: TEG, TEAG; TEG-; (TEAG), etc.
    t = Replace(t, ".", " ")
    t = Replace(t, ",", " ")
    t = Replace(t, ";", " ")
    t = Replace(t, ":", " ")
    t = Replace(t, "-", " ")
    t = Replace(t, "_", " ")
    t = Replace(t, "/", " ")
    t = Replace(t, "\", " ")
    t = Replace(t, "(", " ")
    t = Replace(t, ")", " ")
    t = Replace(t, "[", " ")
    t = Replace(t, "]", " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")

    Do While InStr(1, t, "  ", vbBinaryCompare) > 0
        t = Replace(t, "  ", " ")
    Loop

    If InStr(1, t, " TEAG ", vbTextCompare) > 0 Then
        IdentificarTerminalObservacao = "TEAG"
    ElseIf InStr(1, t, " TEG ", vbTextCompare) > 0 Then
        IdentificarTerminalObservacao = "TEG"
    Else
        IdentificarTerminalObservacao = ""
    End If
End Function

Private Function ValorCelulaTratado(ByVal valor As String) As String
    valor = Trim$(valor)

    If TextoSemAcento(valor) = "PARADO" Then
        ValorCelulaTratado = ""
    Else
        ValorCelulaTratado = valor
    End If
End Function

Private Function NormalizarModoDescarga(ByVal valor As String, Optional ByVal modoPadrao As String = "VAGÃO") As String
    Dim t As String
    t = TextoSemAcento(valor)

    If InStr(1, t, "CAMINH", vbTextCompare) > 0 Then
        NormalizarModoDescarga = "CAMINHÃO"
    ElseIf InStr(1, t, "VAG", vbTextCompare) > 0 Or InStr(1, t, "FERRO", vbTextCompare) > 0 Then
        NormalizarModoDescarga = "VAGÃO"
    Else
        NormalizarModoDescarga = modoPadrao
    End If
End Function

Private Sub AdicionarEnderecoRepetido(ByRef seq As Collection, ByVal ws As Worksheet, ByVal enderecoValor As String, Optional ByVal vezes As Long = 1)
    Dim n As Long

    If Trim$(ws.Range(enderecoValor).Value) = "" Then Exit Sub

    For n = 1 To vezes
        seq.Add enderecoValor
    Next n
End Sub

Private Sub AdicionarMoegaSeModo(ByRef seq As Collection, ByVal ws As Worksheet, ByVal enderecoValor As String, ByVal enderecoModo As String, ByVal modoDesejado As String, Optional ByVal vezes As Long = 1)
    Dim modoAtual As String

    If Trim$(ws.Range(enderecoValor).Value) = "" Then Exit Sub

    modoAtual = NormalizarModoDescarga(CStr(ws.Range(enderecoModo).Value), "VAGÃO")

    If modoAtual = modoDesejado Then
        AdicionarEnderecoRepetido seq, ws, enderecoValor, vezes
    End If
End Sub

Private Function ValorDaProximaCelula(ByVal ws As Worksheet, ByVal seq As Collection, ByRef indiceAtual As Long) As String
    If seq.Count = 0 Then
        ValorDaProximaCelula = ""
        Exit Function
    End If

    indiceAtual = indiceAtual + 1
    If indiceAtual > seq.Count Then indiceAtual = 1

    ValorDaProximaCelula = Trim$(ws.Range(CStr(seq(indiceAtual))).Value)
End Function

Private Sub PreencherInputXPath(ByVal driver As Object, ByVal xpath As String, ByVal valor As String)
    Dim el As Object

    Set el = driver.FindElementByXPath(xpath)

    On Error Resume Next
    el.Clear
    On Error GoTo 0

    el.SendKeys valor
End Sub

Private Function EscolherCelulaEstocagem(ByVal ws As Worksheet, _
                                        ByVal ladoTerminal As String, _
                                        ByVal tipoVeiculo As String, _
                                        ByRef idxTEGCaminhao As Long, _
                                        ByRef idxTEGVagao As Long, _
                                        ByRef idxTEAGCaminhao As Long, _
                                        ByRef idxTEAGVagao As Long) As String
    Dim seq As New Collection
    Dim caminhao As Boolean

    caminhao = EhCaminhao(tipoVeiculo)
    ladoTerminal = UCase$(Trim$(ladoTerminal))

    If ladoTerminal = "TEG" Then

        If caminhao Then
            ' TEG caminhões:
            ' Mantém a alternância antiga dos tombadores: A4 4 vezes / B4 1 vez.
            AdicionarEnderecoRepetido seq, ws, "A4", 4
            AdicionarEnderecoRepetido seq, ws, "B4", 1

            ' Moegas C/D entram para caminhão somente se o dropdown estiver como "Descarga Caminhão".
            ' Mantém o peso antigo das moegas: C4 2 vezes / D4 1 vez.
            AdicionarMoegaSeModo seq, ws, "C4", "C5", "CAMINHÃO", 2
            AdicionarMoegaSeModo seq, ws, "D4", "D5", "CAMINHÃO", 1

            EscolherCelulaEstocagem = ValorDaProximaCelula(ws, seq, idxTEGCaminhao)

        Else
            ' TEG vagões:
            ' Moegas C/D só recebem vagão se o dropdown estiver como "Descarga Vagão".
            AdicionarMoegaSeModo seq, ws, "C4", "C5", "VAGÃO", 2
            AdicionarMoegaSeModo seq, ws, "D4", "D5", "VAGÃO", 1

            EscolherCelulaEstocagem = ValorDaProximaCelula(ws, seq, idxTEGVagao)
        End If

    ElseIf ladoTerminal = "TEAG" Then

        If caminhao Then
            ' TEAG caminhões:
            ' E4 é apenas caminhão.
            ' F/G/H entram para caminhão somente se o dropdown estiver como "Descarga Caminhão".
            AdicionarEnderecoRepetido seq, ws, "E4", 1
            AdicionarMoegaSeModo seq, ws, "F4", "F5", "CAMINHÃO", 1
            AdicionarMoegaSeModo seq, ws, "G4", "G5", "CAMINHÃO", 1
            AdicionarMoegaSeModo seq, ws, "H4", "H5", "CAMINHÃO", 1

            EscolherCelulaEstocagem = ValorDaProximaCelula(ws, seq, idxTEAGCaminhao)

        Else
            ' TEAG vagões:
            ' F/G/H só recebem vagão se o dropdown estiver como "Descarga Vagão".
            AdicionarMoegaSeModo seq, ws, "F4", "F5", "VAGÃO", 1
            AdicionarMoegaSeModo seq, ws, "G4", "G5", "VAGÃO", 1
            AdicionarMoegaSeModo seq, ws, "H4", "H5", "VAGÃO", 1

            EscolherCelulaEstocagem = ValorDaProximaCelula(ws, seq, idxTEAGVagao)
        End If

    End If
End Function

Private Function LerValorInputPorIds(ByVal driver As Object, ByVal ids As Variant) As String
    Dim id As Variant
    Dim el As Object
    Dim valor As String

    For Each id In ids
        Set el = Nothing
        valor = ""

        On Error Resume Next
        Set el = driver.FindElementByXPath("//input[@id='" & CStr(id) & "' and @data-slot='input']")
        If el Is Nothing Then Set el = driver.FindElementById(CStr(id))
        On Error GoTo 0

        If Not el Is Nothing Then
            On Error Resume Next
            valor = el.Attribute("value")
            If Trim$(valor) = "" Then valor = el.Text
            On Error GoTo 0

            LerValorInputPorIds = ValorCelulaTratado(valor)
            Exit Function
        End If
    Next id

    LerValorInputPorIds = ""
End Function

Private Function LerTextoElemento(ByVal el As Object) As String
    Dim valor As String

    If el Is Nothing Then
        LerTextoElemento = ""
        Exit Function
    End If

    On Error Resume Next
    valor = el.Attribute("value")
    If Trim$(valor) = "" Then valor = el.Attribute("aria-label")
    If Trim$(valor) = "" Then valor = el.Attribute("title")
    If Trim$(valor) = "" Then valor = el.Text
    On Error GoTo 0

    LerTextoElemento = Trim$(valor)
End Function

Private Function LerDropdownPorIds(ByVal driver As Object, ByVal ids As Variant) As String
    Dim id As Variant
    Dim el As Object
    Dim valor As String

    For Each id In ids
        Set el = Nothing
        valor = ""

        On Error Resume Next
        Set el = driver.FindElementById(CStr(id))
        If el Is Nothing Then Set el = driver.FindElementByXPath("//*[@id='" & CStr(id) & "']")
        If el Is Nothing Then Set el = driver.FindElementByXPath("//*[contains(@id,'" & CStr(id) & "')]")
        If el Is Nothing Then Set el = driver.FindElementByXPath("//*[contains(@name,'" & CStr(id) & "')]")
        If el Is Nothing Then Set el = driver.FindElementByXPath("//*[contains(@data-testid,'" & CStr(id) & "')]")
        On Error GoTo 0

        valor = LerTextoElemento(el)

        If Trim$(valor) <> "" Then
            LerDropdownPorIds = valor
            Exit Function
        End If
    Next id

    LerDropdownPorIds = ""
End Function

Private Function LerModoMoega(ByVal driver As Object, ByVal ids As Variant, Optional ByVal modoPadrao As String = "VAGÃO") As String
    Dim textoDropdown As String

    textoDropdown = LerDropdownPorIds(driver, ids)

    If Trim$(textoDropdown) = "" Then
        LerModoMoega = modoPadrao
    Else
        LerModoMoega = NormalizarModoDescarga(textoDropdown, modoPadrao)
    End If
End Function


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

