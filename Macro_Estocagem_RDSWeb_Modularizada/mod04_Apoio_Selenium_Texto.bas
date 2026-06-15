Attribute VB_Name = "mod04_Apoio_Selenium_Texto"
' =========================================================
' MÓDULO 04 - APOIO / SELENIUM / TEXTOS
' Fica aqui o WebDriver principal (bot), esperas, cliques,
' leitura de inputs/dropdowns e normalização de textos.
' =========================================================

Public bot As WebDriver

Public Function SegundosDecorridos(ByVal inicio As Single) As Double
    ' Evita problema quando a macro atravessa a meia-noite.
    If Timer >= inicio Then
        SegundosDecorridos = Timer - inicio
    Else
        SegundosDecorridos = (86400 - inicio) + Timer
    End If
End Function

Public Sub WaitDriver(ByVal driver As Object, ByVal ms As Long)
    On Error Resume Next
    If Not driver Is Nothing Then
        driver.Wait ms
    Else
        DoEvents
    End If
    On Error GoTo 0
End Sub

Public Function ElementoVisivelHabilitado(ByVal el As Object) As Boolean
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

Public Function WaitClickable(ByVal el As Object, Optional ByVal timeoutSec As Long = 6) As Boolean
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

Public Function TryClickChild(ByVal parent As Object, ByVal xpath As String, Optional ByVal timeoutSec As Long = 8) As Boolean
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

Public Function TextoSemAcento(ByVal valor As String) As String
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

Public Function EhCaminhao(ByVal tipoVeiculo As String) As Boolean
    Dim t As String
    t = TextoSemAcento(tipoVeiculo)
    EhCaminhao = (InStr(1, t, "CAMINH", vbTextCompare) > 0)
End Function

Public Function IdentificarTerminalObservacao(ByVal observacao As String) As String
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

Public Function ValorCelulaTratado(ByVal valor As String) As String
    valor = Trim$(valor)

    If TextoSemAcento(valor) = "PARADO" Then
        ValorCelulaTratado = ""
    Else
        ValorCelulaTratado = valor
    End If
End Function

Public Function NormalizarModoDescarga(ByVal valor As String, Optional ByVal modoPadrao As String = "VAGÃO") As String
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

Public Sub PreencherInputXPath(ByVal driver As Object, ByVal xpath As String, ByVal valor As String)
    Dim el As Object

    Set el = driver.FindElementByXPath(xpath)

    On Error Resume Next
    el.Clear
    On Error GoTo 0

    el.SendKeys valor
End Sub

Public Function LerValorInputPorIds(ByVal driver As Object, ByVal ids As Variant) As String
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

Public Function LerTextoElemento(ByVal el As Object) As String
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

Public Function LerDropdownPorIds(ByVal driver As Object, ByVal ids As Variant) As String
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

Public Function LerModoMoega(ByVal driver As Object, ByVal ids As Variant, Optional ByVal modoPadrao As String = "VAGÃO") As String
    Dim textoDropdown As String

    textoDropdown = LerDropdownPorIds(driver, ids)

    If Trim$(textoDropdown) = "" Then
        LerModoMoega = modoPadrao
    Else
        LerModoMoega = NormalizarModoDescarga(textoDropdown, modoPadrao)
    End If
End Function
