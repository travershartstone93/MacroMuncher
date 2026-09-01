let
    Source = Excel.CurrentWorkbook(){[Name="Sales"]}[Content],
    Filtered = Table.SelectRows(Source, each [Amount] > 1000),
    Sorted = Table.Sort(Filtered, {{"Date", Order.Descending}}),
    Final = Table.RemoveColumns(Sorted, {"TempColumn"})
in Final
