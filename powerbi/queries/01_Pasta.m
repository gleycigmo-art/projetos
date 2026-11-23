// Pasta de trabalho inteira (não carregar no modelo).
let
    Fonte = Excel.Workbook(File.Contents(CaminhoArquivo), null, true)
in
    Fonte
