// Converte qualquer valor em número: nulo, vazio, erro ou texto não numérico viram 0.
// Mesmo comportamento do ToNum das queries originais.
(v as any) as number =>
    let
        r = try (
                if v = null or v = "" then 0
                else if Value.Is(v, type number) then v
                else Number.FromText(Text.Replace(Text.Trim(Text.From(v)), ",", "."), "en-US")
            ) otherwise 0
    in
        if r = null then 0 else r
