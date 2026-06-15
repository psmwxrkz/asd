Attribute VB_Name = "mod03_Regras_Alternancia"
' =========================================================
' MÓDULO 03 - REGRAS DE CÉLULAS E ALTERNÂNCIA
' Altere aqui quando precisar mudar a distribuição entre
' tombadores/moegas ou o peso da alternância.
' =========================================================

Public Sub AdicionarEnderecoRepetido(ByRef seq As Collection, ByVal ws As Worksheet, ByVal enderecoValor As String, Optional ByVal vezes As Long = 1)
    Dim n As Long

    If Trim$(ws.Range(enderecoValor).Value) = "" Then Exit Sub

    For n = 1 To vezes
        seq.Add enderecoValor
    Next n
End Sub

Public Sub AdicionarMoegaSeModo(ByRef seq As Collection, ByVal ws As Worksheet, ByVal enderecoValor As String, ByVal enderecoModo As String, ByVal modoDesejado As String, Optional ByVal vezes As Long = 1)
    Dim modoAtual As String

    If Trim$(ws.Range(enderecoValor).Value) = "" Then Exit Sub

    modoAtual = NormalizarModoDescarga(CStr(ws.Range(enderecoModo).Value), "VAGÃO")

    If modoAtual = modoDesejado Then
        AdicionarEnderecoRepetido seq, ws, enderecoValor, vezes
    End If
End Sub

Public Function ValorDaProximaCelula(ByVal ws As Worksheet, ByVal seq As Collection, ByRef indiceAtual As Long) As String
    If seq.Count = 0 Then
        ValorDaProximaCelula = ""
        Exit Function
    End If

    indiceAtual = indiceAtual + 1
    If indiceAtual > seq.Count Then indiceAtual = 1

    ValorDaProximaCelula = Trim$(ws.Range(CStr(seq(indiceAtual))).Value)
End Function

Public Function EscolherCelulaEstocagem(ByVal ws As Worksheet, _
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
