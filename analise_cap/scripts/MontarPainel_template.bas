Attribute VB_Name = "MontarPainel"
Option Explicit
' ============================================================================
' Monta a aba PAINEL da CAP proposta no arquivo ATIVO.
' - Adiciona 9 consultas novas (v2_FONTES_CAL, fn..., v2_SEMANAS, v2_CAP_PERIODOS).
'   As consultas e abas atuais NAO sao alteradas nem apagadas.
' - Cria as abas PAINEL_AUX, CAP_PERIODOS e PAINEL.
' Rode numa COPIA do arquivo. Pede a pasta onde estao os arquivos .m do kit.
' ============================================================================

Private Const CONSULTAS As String = "v2_FONTES_CAL,fnCALENDARIO_ATIVO,fnFAM_PROPOSTA,fnDEMANDA_MPR,fnTC_COM_CAMPANHA,fnCAP,v2_SEMANAS,v2_CAP_PERIODOS"

Public Sub MontarPainelCAP()
    Dim wb As Workbook, pasta As String, nome As Variant, inicio As Double
    Set wb = ActiveWorkbook

    If MsgBox("Este macro vai adicionar consultas novas e as abas PAINEL, PAINEL_AUX e CAP_PERIODOS em:" & vbCrLf & _
              wb.Name & vbCrLf & vbCrLf & "Use uma CÓPIA do arquivo. Continuar?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    For Each nome In Array("PAINEL", "PAINEL_AUX", "CAP_PERIODOS")
        If ExisteAba(wb, CStr(nome)) Then
            MsgBox "A aba " & nome & " já existe. Apague-a (ou use outra cópia) e rode de novo.", vbExclamation
            Exit Sub
        End If
    Next

    With Application.FileDialog(4) ' msoFileDialogFolderPicker
        .Title = "Escolha a pasta 'consultas' do kit (arquivos .m)"
        If .Show <> -1 Then Exit Sub
        pasta = .SelectedItems(1)
    End With

    On Error GoTo Falha
    inicio = Timer
    Application.ScreenUpdating = False

    ' 1) consultas
    For Each nome In Split(CONSULTAS, ",")
        Application.StatusBar = "Adicionando consulta " & nome & "..."
        AdicionarConsulta wb, CStr(nome), LerUtf8(pasta & Application.PathSeparator & nome & ".m")
    Next

    ' 2) aba auxiliar (listas dos filtros + semanas disponíveis)
    Dim aux As Worksheet, i As Long, meses As Variant
    Set aux = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    aux.Name = "PAINEL_AUX"
    meses = Array("Jan/2026", "Fev/2026", "Mar/2026", "Abr/2026", "Mai/2026", "Jun/2026", _
                  "Jul/2026", "Ago/2026", "Set/2026", "Out/2026", "Nov/2026", "Dez/2026")
    aux.Range("A1").Value = "Mês"
    For i = 0 To 11
        aux.Cells(i + 2, 1).Value = meses(i)
    Next
    aux.Range("E1").Value = "Visão"
    aux.Range("E2").Value = "Semanal"
    aux.Range("E3").Value = "Mensal"
    Application.StatusBar = "Carregando semanas disponíveis (SharePoint)..."
    CarregarConsulta wb, "v2_SEMANAS", aux, "C1", "v2_SEMANAS"

    ' 3) CAP de todos os períodos (é a parte demorada: 12 meses + semanas)
    Dim capWs As Worksheet
    Set capWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    capWs.Name = "CAP_PERIODOS"
    Application.StatusBar = "Calculando a CAP de todos os períodos... (pode levar alguns minutos)"
    CarregarConsulta wb, "v2_CAP_PERIODOS", capWs, "A1", "CAP_PERIODOS"

    ' 4) PAINEL
    Application.StatusBar = "Montando a aba PAINEL..."
    wb.Names.Add Name:="lstSemanas", RefersTo:="=v2_SEMANAS[Periodo]"
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    ws.Name = "PAINEL"
'__CHAMADAS__
    ws.Range("C4").Value = ws.Range("C4").Value   ' primeira semana disponível, como texto
    FormatarPainel ws

    Application.Calculate
    Application.StatusBar = False
    Application.ScreenUpdating = True
    ws.Activate
    MsgBox "PAINEL montado em " & Format((Timer - inicio) / 60, "0.0") & " min." & vbCrLf & vbCrLf & _
           "Confira: com Visão = Semanal e a mesma semana do FILTRO_DATA, as três diferenças da " & _
           "CONFERÊNCIA (F19:H21) devem ser 0,0000.", vbInformation
    Exit Sub

Falha:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Parou com erro: " & Err.Description & vbCrLf & _
           "Nada foi apagado. As consultas e abas criadas até aqui podem ser removidas à mão.", vbCritical
End Sub

Private Function ExisteAba(wb As Workbook, nome As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nome)
    ExisteAba = Not ws Is Nothing
End Function

Private Function LerUtf8(caminho As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile caminho
    LerUtf8 = st.ReadText(-1)
    st.Close
End Function

Private Sub AdicionarConsulta(wb As Workbook, nome As String, formula As String)
    Dim q As Object
    On Error Resume Next
    Set q = wb.Queries(nome)
    On Error GoTo 0
    If q Is Nothing Then
        wb.Queries.Add Name:=nome, Formula:=formula
    Else
        q.Formula = formula
    End If
End Sub

Private Sub CarregarConsulta(wb As Workbook, consulta As String, ws As Worksheet, celula As String, nomeTabela As String)
    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & consulta & ";Extended Properties=""""", _
        Destination:=ws.Range(celula))
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & consulta & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False
    End With
    lo.Name = nomeTabela
End Sub

Private Sub FormatarPainel(ws As Worksheet)
    Const AZUL As Long = 7884319      ' RGB(31, 78, 120)
    Const CINZA As Long = 8421504     ' RGB(128, 128, 128)
    Dim r As Long, r2 As Variant

    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 40
    ws.Columns("E:U").ColumnWidth = 13
    ws.Columns("I").ColumnWidth = 24

    With ws.Range("B1")
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = AZUL
    End With
    ws.Range("B2").Font.Size = 12

    ' filtros
    ws.Range("B3:B6").Font.Bold = True
    With ws.Range("C3:C6")
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Font.Bold = True
        .Font.Size = 11
    End With
    ws.Range("C3").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=PAINEL_AUX!$E$2:$E$3"
    ws.Range("C4").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=lstSemanas"
    ws.Range("C5").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=PAINEL_AUX!$A$2:$A$13"
    ws.Range("C6").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=PAINEL_AUX!$A$2:$A$13"
    ws.Range("E3:F7").Font.Color = CINZA

    ' indicadores
    Cabecalho ws.Range("B8:D8")
    ws.Range("C9:C21").Font.Bold = True
    ws.Range("C9:C21").Font.Size = 12
    ws.Range("C9:C21").HorizontalAlignment = xlRight
    ws.Range("D9:D21").Font.Color = CINZA
    ws.Range("D9:D21").Font.Italic = True
    For Each r2 In Array(13, 17, 21)
        ws.Range("B" & r2 & ":C" & r2).Font.Size = 14
        ws.Range("B" & r2 & ":C" & r2).Font.Bold = True
    Next
    ws.Range("B9:D21").Borders(xlInsideHorizontal).LineStyle = xlContinuous
    ws.Range("B9:D21").Borders(xlInsideHorizontal).Color = RGB(217, 217, 217)
    StatusCor ws.Range("C13"), "$C$13"
    StatusCor ws.Range("C17"), "$C$17"

    ' complexidade (barra de texto proporcional ao %)
    Cabecalho ws.Range("F8:J8")
    ws.Range("F9:J9").Font.Bold = True
    For r = 10 To 13
        ws.Range("J" & r).Formula = "=REPT(" & Chr(34) & ChrW(9608) & Chr(34) & ",ROUND(H" & r & "*40,0))"
        ws.Range("J" & r).Font.Color = AZUL
    Next
    ws.Range("F14:I15").Font.Color = CINZA

    ' conferência
    Cabecalho ws.Range("F17:J17")
    ws.Range("F18").Font.Italic = True
    ws.Range("F18").Font.Color = CINZA
    ws.Range("H19:H21").Font.Bold = True

    ' quadro mês a mês
    ws.Range("B24").Font.Bold = True
    ws.Range("B24").Font.Color = AZUL
    Cabecalho ws.Range("B25:Y25")
    ws.Range("B25:Y25").WrapText = True
    ws.Range("B26:Y38").Borders.LineStyle = xlContinuous
    ws.Range("B26:Y38").Borders.Color = RGB(217, 217, 217)
    ws.Range("B38:Y38").Interior.Color = RGB(242, 242, 242)
    ws.Range("D26:D38").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=1").Interior.Color = RGB(255, 242, 204)

    ' CAP do período
    ws.Range("B41").Font.Bold = True
    ws.Range("B41").Font.Color = AZUL
    Cabecalho ws.Range("B42:U42")
    ws.Range("B42:U42").WrapText = True
    ws.Range("S42:U102").Font.Color = CINZA
    ws.Range("B42:U102").AutoFilter
    ws.Activate
    ws.Range("M43").Select
    With ws.Range("M43:M102,R43:R102").FormatConditions
        .Add(Type:=xlExpression, Formula1:="=$R43=""Sobrecarregado""").Font.Color = RGB(192, 0, 0)
        .Add(Type:=xlExpression, Formula1:="=$R43=""Adequado""").Font.Color = RGB(0, 128, 0)
        .Add(Type:=xlExpression, Formula1:="=$R43=""Sub-utilizado""").Font.Color = RGB(31, 112, 214)
    End With
    ws.Range("M43:M102,R43:R102").Font.Bold = True

    ActiveWindow.DisplayGridlines = False
    ws.Range("C3").Select
End Sub

Private Sub Cabecalho(rg As Range)
    rg.Interior.Color = 7884319
    rg.Font.Color = vbWhite
    rg.Font.Bold = True
End Sub

Private Sub StatusCor(rg As Range, ref As String)
    ' sem funções nem decimais na fórmula, para funcionar em Excel em português
    rg.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & ref & ">1").Font.Color = RGB(192, 0, 0)
    rg.FormatConditions.Add(Type:=xlExpression, Formula1:="=(" & ref & ">=4/5)*(" & ref & "<=1)").Font.Color = RGB(0, 128, 0)
    rg.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & ref & "<4/5").Font.Color = RGB(31, 112, 214)
End Sub

'__SUBS__
