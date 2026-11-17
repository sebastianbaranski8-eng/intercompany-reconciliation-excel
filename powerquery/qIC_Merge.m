// qIC_Merge: full outer join of both ledgers with the same status logic as the formulas
let
    PL = Table.SelectColumns(Excel.CurrentWorkbook(){[Name="tPL"]}[Content], {"Invoice no.", "Posting date", "Amount EUR"}),
    DE = Table.SelectColumns(Excel.CurrentWorkbook(){[Name="tDE"]}[Content], {"Invoice no.", "Posting date", "Amount EUR"}),
    PLRenamed = Table.RenameColumns(PL, {{"Invoice no.", "Invoice PL"}, {"Posting date", "Date PL"}, {"Amount EUR", "Amount PL"}}),
    DERenamed = Table.RenameColumns(DE, {{"Invoice no.", "Invoice DE"}, {"Posting date", "Date DE"}, {"Amount EUR", "Amount DE"}}),
    Merged = Table.Join(PLRenamed, "Invoice PL", DERenamed, "Invoice DE", JoinKind.FullOuter),
    WithInvoice = Table.AddColumn(Merged, "Invoice no.", each if [Invoice PL] <> null then [Invoice PL] else [Invoice DE], type text),
    WithCount = Table.AddColumn(WithInvoice, "Count", each Table.RowCount(Table.SelectRows(WithInvoice, (r) => r[#"Invoice no."] = [#"Invoice no."])), Int64.Type),
    WithSum = Table.AddColumn(WithCount, "Sum", each Number.Round((if [Amount PL] = null then 0 else [Amount PL]) + (if [Amount DE] = null then 0 else [Amount DE]), 2), type number),
    WithStatus = Table.AddColumn(WithSum, "Status", each
        if [Count] > 1 then "Duplicate"
        else if [Amount DE] = null then "Missing in DE"
        else if [Amount PL] = null then "Missing in PL"
        else if [Sum] <> 0 then "Amount difference"
        else if [Date PL] <> [Date DE] then "Date difference"
        else "OK", type text),
    Typed = Table.TransformColumnTypes(WithStatus, {{"Date PL", type date}, {"Date DE", type date}, {"Amount PL", type number}, {"Amount DE", type number}}),
    Result = Table.SelectColumns(Typed, {"Invoice no.", "Date PL", "Amount PL", "Date DE", "Amount DE", "Sum", "Count", "Status"})
in
    Result
