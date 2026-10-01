Attribute VB_Name = "modEnvioPaineis"
Option Explicit

'==========================================================================
' ENVIO DOS LOTES DE "ROTAS E FAMILIAS" PARA OS PAINÉIS (SEM FÓRMULAS)
'
' EnviarLotesParaPaineis     -> distribui os concatenados nos painéis
' DesfazerUltimoEnvio        -> apaga o que o último envio escreveu
' ConverterFormulasEmValores -> (uma vez) congela as fórmulas antigas
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
Private Const SENHA_PAINEIS As String = ""     ' senha de proteção dos painéis (se houver)

' Lotes com data anterior a (hoje - DIAS_RETROATIVOS) não são enviados.
Private Const DIAS_RETROATIVOS As Long = 0

' True = atualiza a consulta de ROTAS E FAMILIAS antes de enviar.
Private Const ATUALIZAR_ROTAS_ANTES As Boolean = False

Private Const ABA_LOG As String = "LOG ENVIO"

'-----------------------------------------------

Private Const RES_ENVIADO As String = "Enviado"
Private Const RES_DESFEITO As String = "Desfeito"

Private mPaineis As Object      ' nome da aba -> informações do painel
Private mLog As Collection

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
           "Data já passou (ignorados): " & nPassados & vbLf & _
           "Não enviados (sem vaga / sem data / sem coluna): " & nProblemas & _
           IIf(nEnviados + nProblemas > 0, vbLf & vbLf & "Detalhes na aba """ & ABA_LOG & """.", ""), _
           IIf(nProblemas > 0, vbExclamation, vbInformation)
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
    Dim col As Long, linhas As Collection, r As Variant

    Set ws = painel("ws")
    famU = Normalizar(familia)

    d = ParaDataSerial(valorData)
    If d = 0 Then
        RegistrarLog "Sem data programada", ws.Name, familia, "", "", conc
        Exit Function
    End If
    If d < CLng(Date) - DIAS_RETROATIVOS Then
        EnviarLote = 3
        Exit Function
    End If

    chave = famU & "|" & ChaveLote(conc)
    If painel("chaves").Exists(chave) Then
        EnviarLote = 2
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
            DesprotegerPainel painel
            ws.Cells(CLng(r), col).Value = conc
            painel("chaves").Item(chave) = True
            RegistrarLog RES_ENVIADO, ws.Name, familia, Format(CDate(d), "dd/mm/yyyy"), ws.Cells(CLng(r), col).Address(False, False), conc
            EnviarLote = 1
            Exit Function
        End If
    Next r

    RegistrarLog "Sem vaga: bloco cheio nesta data", ws.Name, familia, Format(CDate(d), "dd/mm/yyyy"), "", conc
End Function


'==========================================================================
' 2) DESFAZER O ÚLTIMO ENVIO
'==========================================================================
Public Sub DesfazerUltimoEnvio()
    Dim wsLog As Worksheet, ultLin As Long, r As Long, n As Long, ignorados As Long
    Dim ws As Worksheet, cel As Range

    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets(ABA_LOG)
    On Error GoTo 0
    If wsLog Is Nothing Then
        MsgBox "Não há envio registrado para desfazer.", vbInformation
        Exit Sub
    End If

    ultLin = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    If ultLin < 2 Then
        MsgBox "Não há envio registrado para desfazer.", vbInformation
        Exit Sub
    End If
    If MsgBox("Apagar dos painéis os lotes escritos no último envio?" & vbLf & _
              "(Células que a analista já alterou não são mexidas.)", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Set mPaineis = CreateObject("Scripting.Dictionary")
    Application.ScreenUpdating = False
    For r = 2 To ultLin
        If wsLog.Cells(r, 2).Value = RES_ENVIADO Then
            Set ws = ThisWorkbook.Worksheets(CStr(wsLog.Cells(r, 3).Value))
            Set cel = ws.Range(CStr(wsLog.Cells(r, 6).Value))
            If TextoCelula(cel.Value) = TextoCelula(wsLog.Cells(r, 7).Value) Then
                If Not mPaineis.Exists(ws.Name) Then mPaineis.Add ws.Name, InfoProtecao(ws)
                DesprotegerPainel mPaineis(ws.Name)
                cel.ClearContents
                wsLog.Cells(r, 2).Value = RES_DESFEITO
                n = n + 1
            Else
                ignorados = ignorados + 1
            End If
        End If
    Next r
    ReprotegerPaineis
    Application.ScreenUpdating = True

    MsgBox n & " lote(s) removido(s) dos painéis." & _
           IIf(ignorados > 0, vbLf & ignorados & " célula(s) já tinham sido alteradas e foram mantidas.", ""), vbInformation
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
                DesprotegerPainel info
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
    Set PrepararPainel = info
End Function

Private Function InfoProtecao(ws As Worksheet) As Object
    Dim info As Object
    Set info = CreateObject("Scripting.Dictionary")
    info.Add "wsProt", ws
    info.Add "protegida", ws.ProtectContents
    info.Add "desprotegida", False
    Set InfoProtecao = info
End Function

Private Sub DesprotegerPainel(info As Object)
    If info("protegida") And Not info("desprotegida") Then
        info("wsProt").Unprotect Password:=SENHA_PAINEIS
        info("desprotegida") = True
    End If
End Sub

' Protege de novo com as mesmas permissões que o PAINEL - BANCADA usa hoje.
Private Sub ReprotegerPaineis()
    Dim k As Variant, info As Object
    If mPaineis Is Nothing Then Exit Sub
    For Each k In mPaineis.Keys
        Set info = mPaineis(k)
        If info("desprotegida") Then
            info("wsProt").Protect Password:=SENHA_PAINEIS, DrawingObjects:=False, Contents:=True, Scenarios:=False, _
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

Private Sub RegistrarLog(ByVal resultado As String, ByVal painel As String, ByVal familia As String, _
                         ByVal dataProg As String, ByVal celula As String, ByVal conc As String)
    mLog.Add Array(Now, resultado, painel, familia, dataProg, celula, conc)
End Sub

Private Sub GravarLog()
    Dim ws As Worksheet, i As Long, saida() As Variant, k As Long, linha As Variant
    If mLog Is Nothing Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ABA_LOG)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = ABA_LOG
    End If

    ws.Cells.ClearContents
    ws.Range("A1:G1").Value = Array("Data/hora", "Resultado", "Painel", "Família", "Data programada", "Célula", "Concatenado")
    ws.Range("A1:G1").Font.Bold = True
    If mLog.Count > 0 Then
        ReDim saida(1 To mLog.Count, 1 To 7)
        For i = 1 To mLog.Count
            linha = mLog(i)
            For k = 0 To 6
                saida(i, k + 1) = linha(k)
            Next k
        Next i
        ws.Range("A:A").NumberFormat = "dd/mm/yyyy hh:mm"
        ws.Range("E:F").NumberFormat = "@"
        ws.Range("A2").Resize(mLog.Count, 7).Value = saida
    End If
    ws.Columns("A:F").AutoFit
End Sub
