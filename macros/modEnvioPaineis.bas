Attribute VB_Name = "modEnvioPaineis"
Option Explicit

'==========================================================================
' ENVIO DOS LOTES DE "ROTAS E FAMILIAS" PARA OS PAINÉIS (SEM FÓRMULAS)
'
' EnviarLotesParaPaineis     -> distribui os concatenados nos painéis
' DesfazerUltimoEnvio        -> apaga o que o último envio escreveu
' ConverterFormulasEmValores -> (uma vez) congela as fórmulas antigas
'
' Os lotes enviados ficam pintados de cinza claro ("Prog. em andamento"),
' para a analista ver o que veio automático e trocar a cor ao ajustar.
'
' A macro só ESCREVE em células vazias e nunca move nem apaga o que a
' analista já posicionou. Um lote que já está no bloco da família (em
' qualquer data) não é enviado de novo, então a analista pode arrastar
' os lotes livremente para mudar a prioridade.
'==========================================================================

'---------------- CONFIGURAÇÃO ----------------
Private Const ABA_ROTAS As String = "ROTAS E FAMILIAS"
Private Const TABELA_ROTAS As String = "Tabela3_1"
Private Const COL_CONCAT As String = "CONCATENADO"
Private Const COL_DATA_BANCADA As String = "D+8 (-2 dias)"
Private Const COL_DATA_CROMATO As String = "D+8 (-3 dias)"

Private Const LINHA_CABECALHO As Long = 4      ' linha das datas nos painéis
Private Const COL_FAMILIA As Long = 2          ' coluna B dos painéis
' Senha de proteção dos painéis. Se ficar vazia e o painel tiver senha,
' a macro pergunta a senha na hora (uma vez por clique).
Private Const SENHA_PAINEIS As String = ""

' Lotes com data anterior a (hoje - DIAS_RETROATIVOS) não são enviados.
Private Const DIAS_RETROATIVOS As Long = 0

' True = atualiza a consulta de ROTAS E FAMILIAS antes de enviar.
Private Const ATUALIZAR_ROTAS_ANTES As Boolean = False

Private Const ABA_LOG As String = "LOG ENVIO"

' Células preenchidas pela macro recebem a cor da legenda abaixo (linhas 1 a 3
' do painel). Se a legenda não for encontrada, usa o cinza claro padrão.
Private Const PINTAR_ENVIADOS As Boolean = True
Private Const TEXTO_LEGENDA_AUTO As String = "Prog. em andamento"
Private Const COR_AUTO_PADRAO As Long = 12566463   ' RGB(191, 191, 191) - cinza claro

'-----------------------------------------------

Private Const RES_ENVIADO As String = "Enviado"
Private Const RES_DESFEITO As String = "Desfeito"
Private Const RES_PASSADO As String = "Data já passou: não enviado"
Private Const RES_MANTIDO As String = "Não desfeito (célula já alterada)"
Private Const MAX_ENVIOS_LOG As Long = 10   ' quantos envios o LOG ENVIO guarda (menos = arquivo mais leve)

Private mPaineis As Object      ' nome da aba -> informações do painel
Private mLog As Collection
Private mIdEnvio As String        ' identificação (data/hora) do envio em andamento
Private mSenhaDigitada As String  ' senha informada pelo usuário nesta execução

' Família (cabeçalho em ROTAS E FAMILIAS) -> painel -> coluna de data usada.
' O nome da família deve ser igual ao da coluna B do painel.
Private Function MapaFamilias() As Variant
    MapaFamilias = Array( _
        Array("FAM FAST EQUIPAMENTO - UV - Vis", "PAINEL - BANCADA", COL_DATA_BANCADA), _
        Array("FAM FAST EQUIPAMENTO - IR/NIR", "PAINEL - BANCADA", COL_DATA_BANCADA), _
        Array("FAM FAST EQUIPAMENTO - DSC/RAIOX/TGA/DRX", "PAINEL - BANCADA", COL_DATA_BANCADA), _
        Array("FAM FAST - ICP / AA", "PAINEL - BANCADA", COL_DATA_BANCADA), _
        Array("FAM FAST - BANCADA", "PAINEL - BANCADA", COL_DATA_BANCADA), _
        Array("FAM EXCIP", "PAINEL - BANCADA", COL_DATA_BANCADA), _
        Array("FAM IFA", "PAINEL - BANCADA", COL_DATA_BANCADA), _
        Array("Teste CG", "PAINEL - CG", COL_DATA_CROMATO), _
        Array("Teste HPLC", "PAINEL - HPLC", COL_DATA_CROMATO))
End Function


'==========================================================================
' 1) ENVIAR
'==========================================================================
Public Sub EnviarLotesParaPaineis()
    Dim lo As ListObject, mapa As Variant, dados As Variant
    Dim cConcat As Long, i As Long, j As Long
    Dim colFam() As Long, colData() As Long, faltando As String
    Dim nEnviados As Long, nJaNoPainel As Long, nPassados As Long, nProblemas As Long
    Dim calcAnterior As XlCalculation

    Set lo = TabelaRotas()
    If lo Is Nothing Then Exit Sub

    If ATUALIZAR_ROTAS_ANTES Then
        On Error Resume Next
        lo.QueryTable.Refresh BackgroundQuery:=False
        If Err.Number <> 0 Then
            MsgBox "Não foi possível atualizar a tabela " & TABELA_ROTAS & "." & vbLf & _
                   "O envio vai usar os dados que já estão na aba." & vbLf & vbLf & Err.Description, vbExclamation
            Err.Clear
        End If
        On Error GoTo 0
    End If

    If lo.DataBodyRange Is Nothing Then
        MsgBox "A tabela " & TABELA_ROTAS & " está vazia.", vbInformation
        Exit Sub
    End If

    ' Localiza as colunas pelo nome do cabeçalho (não depende da letra).
    mapa = MapaFamilias()
    ReDim colFam(LBound(mapa) To UBound(mapa))
    ReDim colData(LBound(mapa) To UBound(mapa))
    cConcat = IndiceColuna(lo, COL_CONCAT)
    If cConcat = 0 Then faltando = faltando & vbLf & "  - " & COL_CONCAT
    For j = LBound(mapa) To UBound(mapa)
        colFam(j) = IndiceColuna(lo, CStr(mapa(j)(0)))
        colData(j) = IndiceColuna(lo, CStr(mapa(j)(2)))
        If colFam(j) = 0 Then faltando = faltando & vbLf & "  - " & mapa(j)(0)
        If colData(j) = 0 And InStr(faltando, CStr(mapa(j)(2))) = 0 Then faltando = faltando & vbLf & "  - " & mapa(j)(2)
    Next j
    If Len(faltando) > 0 Then
        MsgBox "Colunas não encontradas na tabela " & TABELA_ROTAS & ":" & faltando & vbLf & vbLf & _
               "Confira os nomes em MapaFamilias() no módulo modEnvioPaineis.", vbCritical
        Exit Sub
    End If

    calcAnterior = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Falha

    Set mPaineis = CreateObject("Scripting.Dictionary")
    Set mLog = New Collection
    mIdEnvio = Format(Now, "dd/mm/yyyy hh:mm:ss")
    mSenhaDigitada = ""
    For j = LBound(mapa) To UBound(mapa)
        If Not mPaineis.Exists(CStr(mapa(j)(1))) Then
            mPaineis.Add CStr(mapa(j)(1)), PrepararPainel(ThisWorkbook.Worksheets(CStr(mapa(j)(1))))
        End If
    Next j

    dados = lo.DataBodyRange.Value

    For i = 1 To UBound(dados, 1)
        Dim conc As String
        conc = TextoCelula(dados(i, cConcat))
        If conc <> "" Then
            For j = LBound(mapa) To UBound(mapa)
                If TextoCelula(dados(i, colFam(j))) <> "" Then
                    Select Case EnviarLote(conc, CStr(mapa(j)(0)), mPaineis(CStr(mapa(j)(1))), dados(i, colData(j)))
                        Case 1: nEnviados = nEnviados + 1
                        Case 2: nJaNoPainel = nJaNoPainel + 1
                        Case 3: nPassados = nPassados + 1
                        Case Else: nProblemas = nProblemas + 1
                    End Select
                End If
            Next j
        End If
    Next i

    ReprotegerPaineis
    GravarLog
    Application.Calculation = calcAnterior
    Application.ScreenUpdating = True

    MsgBox "Envio concluído." & vbLf & vbLf & _
           "Enviados agora: " & nEnviados & vbLf & _
           "Já estavam no painel: " & nJaNoPainel & vbLf & _
           "Data já passou (não enviados): " & nPassados & vbLf & _
           "Não enviados (sem vaga / sem data / sem coluna): " & nProblemas & _
           IIf(nEnviados + nPassados + nProblemas > 0, vbLf & vbLf & "Lote a lote na aba """ & ABA_LOG & """.", ""), _
           IIf(nPassados + nProblemas > 0, vbExclamation, vbInformation)
    Exit Sub

Falha:
    Dim msg As String
    msg = Err.Description
    On Error Resume Next
    ReprotegerPaineis
    GravarLog
    Application.Calculation = calcAnterior
    Application.ScreenUpdating = True
    MsgBox "Erro durante o envio: " & msg & vbLf & _
           "O que já foi escrito está registrado na aba """ & ABA_LOG & """.", vbCritical
End Sub

' Retorna 1 = enviado, 2 = já estava no painel, 3 = data passada, 0 = não enviado.
Private Function EnviarLote(ByVal conc As String, ByVal familia As String, painel As Object, ByVal valorData As Variant) As Long
    Dim ws As Worksheet, d As Long, chave As String, famU As String
    Dim col As Long, linhas As Collection, r As Variant, corAnterior As Long

    Set ws = painel("ws")
    famU = Normalizar(familia)

    ' Lote que já está no painel (em qualquer data do bloco) não precisa de nada.
    chave = famU & "|" & ChaveLote(conc)
    If painel("chaves").Exists(chave) Then
        EnviarLote = 2
        Exit Function
    End If

    d = ParaDataSerial(valorData)
    If d = 0 Then
        RegistrarLog "Sem data programada", ws.Name, familia, "", "", conc
        Exit Function
    End If
    ' Programação só de hoje em diante; o atrasado fica visível no LOG ENVIO.
    If d < CLng(Date) - DIAS_RETROATIVOS Then
        RegistrarLog RES_PASSADO, ws.Name, familia, Format(CDate(d), "dd/mm/yyyy"), "", conc
        EnviarLote = 3
        Exit Function
    End If

    If Not painel("blocos").Exists(famU) Then
        RegistrarLog "Família não existe na coluna B do painel", ws.Name, familia, Format(CDate(d), "dd/mm/yyyy"), "", conc
        Exit Function
    End If
    If Not painel("datas").Exists(d) Then
        RegistrarLog "Data não existe na linha " & LINHA_CABECALHO & " do painel", ws.Name, familia, Format(CDate(d), "dd/mm/yyyy"), "", conc
        Exit Function
    End If

    col = painel("datas")(d)
    Set linhas = painel("blocos")(famU)
    For Each r In linhas
        If CelulaVazia(ws.Cells(CLng(r), col)) Then
            If Not DesprotegerPainel(painel) Then
                RegistrarLog "Painel protegido: senha não informada ou incorreta", ws.Name, familia, Format(CDate(d), "dd/mm/yyyy"), "", conc
                Exit Function
            End If
            corAnterior = CorDaCelula(ws.Cells(CLng(r), col))
            ws.Cells(CLng(r), col).Value = conc
            If PINTAR_ENVIADOS Then ws.Cells(CLng(r), col).Interior.Color = painel("cor")
            painel("chaves").Item(chave) = True
            RegistrarLog RES_ENVIADO, ws.Name, familia, Format(CDate(d), "dd/mm/yyyy"), ws.Cells(CLng(r), col).Address(False, False), conc, corAnterior
            EnviarLote = 1
            Exit Function
        End If
    Next r

    RegistrarLog "Sem vaga: bloco cheio nesta data", ws.Name, familia, Format(CDate(d), "dd/mm/yyyy"), "", conc
End Function


'==========================================================================
' 2) DESFAZER O ÚLTIMO ENVIO
'    Cada clique desfaz um envio, do mais recente para o mais antigo.
'    Só apaga a célula se ela ainda estiver no lugar e com o mesmo texto:
'    o que a analista já moveu ou editou é mantido.
'==========================================================================
Public Sub DesfazerUltimoEnvio()
    Dim wsLog As Worksheet, ultLin As Long, r As Long, n As Long, ignorados As Long
    Dim ws As Worksheet, cel As Range, alvo As String, achou As Boolean, total As Long

    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets(ABA_LOG)
    On Error GoTo 0
    If Not wsLog Is Nothing Then ultLin = wsLog.Cells(wsLog.Rows.Count, 2).End(xlUp).Row

    ' O log fica com o envio mais recente no topo: o primeiro "Enviado" é o alvo.
    For r = 2 To ultLin
        If wsLog.Cells(r, 2).Value = RES_ENVIADO Then
            alvo = CStr(wsLog.Cells(r, 9).Value)
            achou = True
            Exit For
        End If
    Next r
    If Not achou Then
        MsgBox "Não há envio para desfazer.", vbInformation
        Exit Sub
    End If

    For r = 2 To ultLin
        If wsLog.Cells(r, 2).Value = RES_ENVIADO And CStr(wsLog.Cells(r, 9).Value) = alvo Then total = total + 1
    Next r
    If MsgBox("Desfazer o envio " & IIf(alvo = "", "anterior", "de " & alvo) & "?" & vbLf & vbLf & _
              total & " lote(s) serão apagados dos painéis." & vbLf & _
              "Lotes que a analista já moveu ou editou não são mexidos.", vbQuestion + vbYesNo, "Desfazer envio") = vbNo Then Exit Sub

    Set mPaineis = CreateObject("Scripting.Dictionary")
    mSenhaDigitada = ""
    Application.ScreenUpdating = False
    For r = 2 To ultLin
        If wsLog.Cells(r, 2).Value = RES_ENVIADO And CStr(wsLog.Cells(r, 9).Value) = alvo Then
            Set ws = Nothing
            Set cel = Nothing
            On Error Resume Next
            Set ws = ThisWorkbook.Worksheets(CStr(wsLog.Cells(r, 3).Value))
            If Not ws Is Nothing Then Set cel = ws.Range(CStr(wsLog.Cells(r, 6).Value))
            On Error GoTo 0
            If cel Is Nothing Then
                wsLog.Cells(r, 2).Value = RES_MANTIDO
                ignorados = ignorados + 1
            ElseIf TextoCelula(cel.Value) <> TextoCelula(wsLog.Cells(r, 7).Value) Then
                wsLog.Cells(r, 2).Value = RES_MANTIDO
                ignorados = ignorados + 1
            Else
                If Not mPaineis.Exists(ws.Name) Then mPaineis.Add ws.Name, InfoProtecao(ws)
                If DesprotegerPainel(mPaineis(ws.Name)) Then
                    cel.ClearContents
                    If TextoCelula(wsLog.Cells(r, 8).Value) <> "" Then RestaurarCor cel, CLng(wsLog.Cells(r, 8).Value)
                    wsLog.Cells(r, 2).Value = RES_DESFEITO
                    n = n + 1
                Else
                    ignorados = ignorados + 1
                End If
            End If
        End If
    Next r
    ReprotegerPaineis
    Application.ScreenUpdating = True

    MsgBox n & " lote(s) removido(s) dos painéis." & _
           IIf(ignorados > 0, vbLf & ignorados & " lote(s) mantido(s): a analista já tinha movido/editado a célula.", "") & _
           vbLf & vbLf & "Clique de novo em Desfazer para voltar mais um envio.", vbInformation, "Desfazer envio"
End Sub


'==========================================================================
' 3) CONGELAR AS FÓRMULAS ANTIGAS (rodar uma única vez)
'==========================================================================
Public Sub ConverterFormulasEmValores()
    Dim mapa As Variant, j As Long, nomes As Object, nome As Variant
    Dim ws As Worksheet, info As Object, area As Range, formulas As Range, cel As Range, n As Long

    If MsgBox("Isto troca as fórmulas antigas dos painéis pelo texto que elas mostram hoje." & vbLf & _
              "Depois disso, use apenas o botão de envio." & vbLf & vbLf & "Continuar?", _
              vbQuestion + vbYesNo) = vbNo Then Exit Sub

    mapa = MapaFamilias()
    Set nomes = CreateObject("Scripting.Dictionary")
    For j = LBound(mapa) To UBound(mapa)
        nomes(CStr(mapa(j)(1))) = True
    Next j

    Set mPaineis = CreateObject("Scripting.Dictionary")
    mSenhaDigitada = ""
    Application.ScreenUpdating = False
    For Each nome In nomes.Keys
        Set ws = ThisWorkbook.Worksheets(CStr(nome))
        Set info = PrepararPainel(ws)
        mPaineis.Add ws.Name, info
        If info("primeiraCol") > 0 And info("ultLin") > LINHA_CABECALHO Then
            Set area = ws.Range(ws.Cells(LINHA_CABECALHO + 1, info("primeiraCol")), ws.Cells(info("ultLin"), info("ultCol")))
            Set formulas = Nothing
            On Error Resume Next
            Set formulas = area.SpecialCells(xlCellTypeFormulas)
            On Error GoTo 0
            If Not formulas Is Nothing Then
                If Not DesprotegerPainel(info) Then
                    MsgBox "As fórmulas do " & ws.Name & " não foram convertidas (senha não informada).", vbExclamation
                    Set formulas = Nothing
                End If
            End If
            If Not formulas Is Nothing Then
                For Each cel In formulas.Cells
                    If cel.HasArray Then
                        cel.CurrentArray.Value = cel.CurrentArray.Value
                    Else
                        cel.Value = cel.Value
                    End If
                    n = n + 1
                Next cel
            End If
        End If
    Next nome
    ReprotegerPaineis
    Application.ScreenUpdating = True

    MsgBox n & " fórmula(s) convertida(s) em valor.", vbInformation
End Sub


'==========================================================================
' AUXILIARES
'==========================================================================
Private Function TabelaRotas() As ListObject
    On Error Resume Next
    Set TabelaRotas = ThisWorkbook.Worksheets(ABA_ROTAS).ListObjects(TABELA_ROTAS)
    On Error GoTo 0
    If TabelaRotas Is Nothing Then
        MsgBox "Tabela """ & TABELA_ROTAS & """ não encontrada na aba """ & ABA_ROTAS & """.", vbCritical
    End If
End Function

Private Function IndiceColuna(lo As ListObject, ByVal nome As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If Normalizar(lc.Name) = Normalizar(nome) Then
            IndiceColuna = lc.Index
            Exit Function
        End If
    Next lc
End Function

' Lê a estrutura do painel: colunas de data, linhas de cada família e lotes já presentes.
Private Function PrepararPainel(ws As Worksheet) As Object
    Dim info As Object, datas As Object, blocos As Object, chaves As Object
    Dim ultCol As Long, ultLin As Long, primeiraCol As Long
    Dim c As Long, r As Long, d As Long, dDireita As Long, fam As String
    Dim valores As Variant, txt As String

    Set info = InfoProtecao(ws)
    Set datas = CreateObject("Scripting.Dictionary")
    Set blocos = CreateObject("Scripting.Dictionary")
    Set chaves = CreateObject("Scripting.Dictionary")

    ultCol = ws.Cells(LINHA_CABECALHO, ws.Columns.Count).End(xlToLeft).Column
    ultLin = ws.Cells(ws.Rows.Count, COL_FAMILIA).End(xlUp).Row

    ' Da direita para a esquerda: cabeçalhos "dd/mm" herdam o ano da coluna seguinte.
    For c = ultCol To COL_FAMILIA + 1 Step -1
        d = DataDoCabecalho(ws.Cells(LINHA_CABECALHO, c).Value, dDireita)
        If d > 0 Then
            If Not datas.Exists(d) Then datas.Add d, c
            dDireita = d
            primeiraCol = c
        End If
    Next c

    If ultLin > LINHA_CABECALHO And primeiraCol > 0 Then
        valores = ws.Range(ws.Cells(1, 1), ws.Cells(ultLin, ultCol)).Value
        For r = LINHA_CABECALHO + 1 To ultLin
            fam = Normalizar(valores(r, COL_FAMILIA))
            If fam <> "" Then
                If Not blocos.Exists(fam) Then blocos.Add fam, New Collection
                blocos(fam).Add r
                For c = primeiraCol To ultCol
                    txt = TextoCelula(valores(r, c))
                    If txt <> "" Then chaves(fam & "|" & ChaveLote(txt)) = True
                Next c
            End If
        Next r
    End If

    info.Add "ws", ws
    info.Add "datas", datas
    info.Add "blocos", blocos
    info.Add "chaves", chaves
    info.Add "primeiraCol", primeiraCol
    info.Add "ultCol", ultCol
    info.Add "ultLin", ultLin
    info.Add "cor", CorDaLegenda(ws)
    Set PrepararPainel = info
End Function

Private Function InfoProtecao(ws As Worksheet) As Object
    Dim info As Object
    Set info = CreateObject("Scripting.Dictionary")
    info.Add "wsProt", ws
    info.Add "protegida", ws.ProtectContents
    info.Add "desprotegida", False
    info.Add "recusada", False
    info.Add "senha", ""
    Set InfoProtecao = info
End Function

' Desprotege o painel antes de escrever. Tenta a SENHA_PAINEIS, depois a senha
' já digitada neste clique e, se nenhuma servir, pergunta ao usuário.
' Retorna False se não conseguir (o painel é pulado, sem travar o envio).
Private Function DesprotegerPainel(info As Object) As Boolean
    Dim ws As Worksheet, senha As String, tentativa As Long

    If Not info("protegida") Or info("desprotegida") Then
        DesprotegerPainel = True
        Exit Function
    End If
    If info("recusada") Then Exit Function

    Set ws = info("wsProt")
    If TentarDesproteger(ws, SENHA_PAINEIS) Then
        senha = SENHA_PAINEIS
    ElseIf mSenhaDigitada <> "" And TentarDesproteger(ws, mSenhaDigitada) Then
        senha = mSenhaDigitada
    Else
        Application.ScreenUpdating = True
        For tentativa = 1 To 3
            senha = InputBox("A aba """ & ws.Name & """ está protegida com senha." & vbLf & vbLf & _
                             "Digite a senha para a macro poder escrever os lotes" & _
                             IIf(tentativa > 1, " (senha incorreta, tentativa " & tentativa & " de 3)", "") & ":", _
                             "Senha do painel")
            If senha = "" Then Exit For
            If TentarDesproteger(ws, senha) Then Exit For
            senha = ""
        Next tentativa
        Application.ScreenUpdating = False
        If senha = "" Then
            info("recusada") = True
            Exit Function
        End If
        mSenhaDigitada = senha
    End If

    info("senha") = senha
    info("desprotegida") = True
    DesprotegerPainel = True
End Function

Private Function TentarDesproteger(ws As Worksheet, ByVal senha As String) As Boolean
    On Error Resume Next
    ws.Unprotect Password:=senha
    TentarDesproteger = (Err.Number = 0 And Not ws.ProtectContents)
    Err.Clear
End Function

' Protege de novo, com a mesma senha, e com as permissões que o PAINEL - BANCADA usa hoje.
Private Sub ReprotegerPaineis()
    Dim k As Variant, info As Object
    If mPaineis Is Nothing Then Exit Sub
    For Each k In mPaineis.Keys
        Set info = mPaineis(k)
        If info("desprotegida") Then
            info("wsProt").Protect Password:=CStr(info("senha")), DrawingObjects:=False, Contents:=True, Scenarios:=False, _
                AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
                AllowSorting:=True, AllowFiltering:=True
            info("desprotegida") = False
        End If
    Next k
End Sub

' Cabeçalho "dd/mm/aaaa", "dd/mm" ou data de verdade -> número de série da data (0 se não for data).
Private Function DataDoCabecalho(ByVal v As Variant, ByVal dDireita As Long) As Long
    Dim p() As String, s As String, ano As Long, dt As Date
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        DataDoCabecalho = CLng(Int(CDbl(v)))
        Exit Function
    End If
    s = Trim(CStr(v))
    p = Split(s, "/")
    If UBound(p) < 1 Or UBound(p) > 2 Then Exit Function
    If Not (IsNumeric(p(0)) And IsNumeric(p(1))) Then Exit Function
    On Error GoTo Invalida
    If UBound(p) = 2 Then
        If Not IsNumeric(p(2)) Then Exit Function
        dt = DateSerial(CLng(p(2)), CLng(p(1)), CLng(p(0)))
    Else
        If dDireita > 0 Then ano = Year(CDate(dDireita)) Else ano = Year(Date)
        dt = DateSerial(ano, CLng(p(1)), CLng(p(0)))
        If dDireita > 0 And CLng(dt) > dDireita Then dt = DateSerial(ano - 1, CLng(p(1)), CLng(p(0)))
    End If
    DataDoCabecalho = CLng(dt)
Invalida:
End Function

Private Function ParaDataSerial(ByVal v As Variant) As Long
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Or IsNumeric(v) Then
        If CDbl(v) > 0 Then ParaDataSerial = CLng(Int(CDbl(v)))
    ElseIf IsDate(v) Then
        ParaDataSerial = CLng(Int(CDbl(CDate(v))))
    End If
End Function

' Identifica o lote pelo trecho antes de "| D+8" (código, descrição e lotes),
' para que uma mudança de data em ROTAS E FAMILIAS não gere duplicata no painel.
Private Function ChaveLote(ByVal s As String) As String
    Dim p As Long
    p = InStr(1, s, "| D+8", vbTextCompare)
    If p > 0 Then s = Left$(s, p - 1)
    ChaveLote = UCase$(Replace(Replace(s, " ", ""), vbLf, ""))
End Function

Private Function Normalizar(ByVal v As Variant) As String
    Normalizar = UCase$(TextoCelula(v))
End Function

Private Function TextoCelula(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    TextoCelula = Trim$(CStr(v))
End Function

' Vazia = sem texto. Fórmula que devolve "" conta como vazia; erro (#REF! etc.) conta como ocupada.
Private Function CelulaVazia(cel As Range) As Boolean
    If IsError(cel.Value) Then Exit Function
    CelulaVazia = (TextoCelula(cel.Value) = "")
End Function

' Cor da célula da legenda (linhas 1 a 3) cujo texto é TEXTO_LEGENDA_AUTO.
Private Function CorDaLegenda(ws As Worksheet) As Long
    Dim cel As Range
    CorDaLegenda = COR_AUTO_PADRAO
    For Each cel In ws.Range("A1:Z3").Cells
        If Normalizar(cel.Value) = Normalizar(TEXTO_LEGENDA_AUTO) Then
            If cel.Interior.Pattern <> xlNone Then CorDaLegenda = cel.Interior.Color
            Exit Function
        End If
    Next cel
End Function

' -1 = célula sem preenchimento.
Private Function CorDaCelula(cel As Range) As Long
    If cel.Interior.Pattern = xlNone Then
        CorDaCelula = -1
    Else
        CorDaCelula = cel.Interior.Color
    End If
End Function

Private Sub RestaurarCor(cel As Range, ByVal cor As Long)
    If cor = -1 Then
        cel.Interior.Pattern = xlNone
    Else
        cel.Interior.Color = cor
    End If
End Sub

Private Sub RegistrarLog(ByVal resultado As String, ByVal painel As String, ByVal familia As String, _
                         ByVal dataProg As String, ByVal celula As String, ByVal conc As String, _
                         Optional ByVal corAnterior As Variant = "")
    mLog.Add Array(Now, resultado, painel, familia, dataProg, celula, conc, corAnterior, mIdEnvio)
End Sub

' Acrescenta o envio atual no topo do LOG ENVIO (mais recente primeiro) e
' mantém só os últimos MAX_ENVIOS_LOG envios. Um clique que não gerou nenhuma
' linha não apaga o histórico, então o Desfazer continua funcionando.
Private Sub GravarLog()
    Dim ws As Worksheet, i As Long, saida() As Variant, k As Long, linha As Variant

    If mLog Is Nothing Then Exit Sub
    If mLog.Count = 0 Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ABA_LOG)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = ABA_LOG
    End If

    ws.Range("A1:I1").Value = Array("Data/hora", "Resultado", "Painel", "Família", "Data programada", "Célula", "Concatenado", "Cor anterior", "Envio")
    ws.Range("A1:I1").Font.Bold = True

    ReDim saida(1 To mLog.Count, 1 To 9)
    For i = 1 To mLog.Count
        linha = mLog(i)
        For k = 0 To 8
            saida(i, k + 1) = linha(k)
        Next k
    Next i
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Rows("2:" & mLog.Count + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    ws.Range("A:A").NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Range("E:F").NumberFormat = "@"
    ws.Range("I:I").NumberFormat = "@"
    ws.Range("A2").Resize(mLog.Count, 9).Value = saida

    LimparLogAntigo ws
    ws.Range("A1:I" & ws.Cells(ws.Rows.Count, 2).End(xlUp).Row).AutoFilter
    ws.Columns("A:F").AutoFit
    ws.Columns("I").AutoFit
End Sub

Private Sub LimparLogAntigo(ws As Worksheet)
    Dim ultLin As Long, r As Long, ids As Object, id As String
    ultLin = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    Set ids = CreateObject("Scripting.Dictionary")
    For r = 2 To ultLin
        id = CStr(ws.Cells(r, 9).Value)
        If Not ids.Exists(id) Then
            If ids.Count >= MAX_ENVIOS_LOG Then
                ws.Rows(r & ":" & ultLin).Delete
                Exit Sub
            End If
            ids.Add id, True
        End If
    Next r
End Sub
