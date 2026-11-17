// qNBP_EUR: NBP average EUR/PLN rates (Table A), 30 Sep - 30 Oct 2026, from the NBP API
let
    Source = Json.Document(Web.Contents("https://api.nbp.pl/api/exchangerates/rates/a/eur/2026-09-30/2026-10-30/?format=json")),
    Rates = Source[rates],
    RatesTable = Table.FromRecords(Rates),
    Typed = Table.TransformColumnTypes(RatesTable, {{"effectiveDate", type date}, {"mid", type number}, {"no", type text}}),
    Renamed = Table.RenameColumns(Typed, {{"no", "NBP table no."}, {"effectiveDate", "Date"}, {"mid", "EUR/PLN"}}),
    Reordered = Table.ReorderColumns(Renamed, {"Date", "EUR/PLN", "NBP table no."})
in
    Reordered
