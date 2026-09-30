shared Tabela_DemandaSemanal = let
    URLSite      = "https://hypermarcas.sharepoint.com/sites/ControledeQualidade-PA",
    CaminhoPasta = URLSite & "/Documentos Compartilhados/MFV/",
    NomeArquivo  = "Base Quarentena tratada.xlsx",
    NomeAba      = "Tbela_DemandaSemanal_futura",

    Conteudo =
        SharePoint.Files(URLSite, [ApiVersion = 15])
            {[Name = NomeArquivo, #"Folder Path" = CaminhoPasta]}[Content],

    Objetos = Excel.Workbook(Binary.Buffer(Conteudo), null, false),

    Filtrado =
        Table.SelectRows(
            Objetos,
            each [Kind] = "Sheet"
                 and Comparer.OrdinalIgnoreCase(Text.Trim([Item]), NomeAba) = 0
        ),

    Aba =
        if Table.IsEmpty(Filtrado)
        then error "Aba '" & NomeAba & "' não encontrada em " & NomeArquivo
        else Filtrado{0}[Data],

    Promovido = Table.PromoteHeaders(Aba, [PromoteAllScalars = true]),

    #"Linhas Válidas" =
        Table.SelectRows(
            Promovido,
            each [Cod] <> null and [Cod] <> "" and [#"Chave semana"] <> null
        ),

    #"Tipo Alterado" =
        Table.TransformColumnTypes(
            #"Linhas Válidas",
            {{"Cod",             type text},
             {"Chave semana",    Int64.Type},
             {"STATUS",          type text},
             {"Lote semana",     Int64.Type},
             {"MÉTODO DOCNIX",   type text},
             {"TC levantados?",  type text}}
        ),

    #"Status2 Adicionado" =
        Table.AddColumn(
            #"Tipo Alterado",
            "Status2",
            each
                if [MÉTODO DOCNIX] = "Método não encontrado" then "Sem tempos"
                else if [MÉTODO DOCNIX] = null then "Sem tempos"
                else if [#"TC levantados?"] = "NÃO" then "Sem tempos"
                else "Tempos OK",
            type text
        ),

    #"Colunas Selecionadas" =
        Table.SelectColumns(
            #"Status2 Adicionado",
            {"Cod", "Chave semana", "STATUS", "Lote semana",
             "MÉTODO DOCNIX", "TC levantados?", "Status2"}
        )
in
    #"Colunas Selecionadas";
