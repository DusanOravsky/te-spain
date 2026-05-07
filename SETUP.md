# Liquidaciones — Excel automation setup

One-time setup of the `Liquidaciones de combustible` workbook so the monthly
process becomes: **drop vendor file in folder → Refresh All → Close month → copy Email sheet**.

You'll do this once on a copy of the workbook. Save the result as `.xlsm`.

---

## Prerequisites

- Excel 2016+ (Power Query built in)
- A working copy: open `Liquidaciones de combustible 29.04.2026_incluyendo gasolina.xlsx`
  → File → Save As → keep as `.xlsx` for now (we'll change extension later) →
  name it whatever you want.
- The user will choose a folder each month; we'll wire that to a `Settings` sheet.

---

## STEP 1 — Add a "Settings" sheet

1. Right-click any sheet tab → Insert → Worksheet → name it **`Settings`**
2. In `Settings`, type:

   | Cell | Value |
   |------|-------|
   | `A1` | `Folder path` |
   | `B1` | (the absolute folder path the user will use, e.g. `C:\Covestro\Liquidaciones`) |

3. Select `B1` → Formulas → Define Name → Name = `FolderPath`. Click OK.

The user updates `B1` each month if they change folder. Everything else reads `FolderPath`.

---

## STEP 2 — Query 1: WebPrices (fetches geoportal averages)

1. Data tab → Get Data → From Other Sources → **Blank Query**
2. Power Query Editor opens. Home → **Advanced Editor**.
3. Replace everything with:

```m
let
    Source = Excel.Workbook(
        Web.Contents("https://geoportalgasolineras.es/geoportal/resources/files/preciosEESS_es.xls"),
        null, true
    ),
    Sheet = Source{0}[Data],
    Promoted = Table.PromoteHeaders(Table.Skip(Sheet, 3), [PromoteAllScalars=true]),
    ToNum = (v) => try Number.FromText(Text.From(v), "es-ES") otherwise null,
    G95 = List.Select(List.Transform(Table.Column(Promoted, "Precio gasolina 95 E5"), ToNum), each _ <> null),
    GA  = List.Select(List.Transform(Table.Column(Promoted, "Precio gasóleo A"),       ToNum), each _ <> null),
    Snapshot = try Text.From(Sheet{0}[Column2]) otherwise "",
    Result = #table(
        {"Fuel", "Average", "Stations", "Snapshot"},
        {
            {"Gasolina 95 E5", List.Average(G95), List.Count(G95), Snapshot},
            {"Gasóleo A",      List.Average(GA),  List.Count(GA),  Snapshot}
        }
    )
in
    Result
```

4. Done → name the query **`WebPrices`** in the right panel
5. If a "Connection" / "Privacy" prompt appears, set **Anonymous** and tick "ignore privacy levels"
6. Home → Close & Load To… → Table → New worksheet — call it **`Prices`**

You should now see the two-row table on `Prices`. Note the row positions (typically B2 = Gasolina avg, B3 = Gasóleo avg).

---

## STEP 3 — Query 2: VendorFleet (newest Flota Viva file in folder)

1. Data → Get Data → From File → **From Folder**
2. Type a placeholder path (we'll replace with `FolderPath` in a moment), click OK → Combine & Edit
3. After it opens, Home → Advanced Editor. Replace everything with:

```m
let
    Folder = Folder.Files(Excel.CurrentWorkbook(){[Name="FolderPath"]}[Content]{0}[Column1]),
    Filtered = Table.SelectRows(Folder, each
        Text.StartsWith([Name], "Flota Viva Covestro") and
        Text.EndsWith(Text.Lower([Name]), ".xlsx") and
        not Text.StartsWith([Name], "~$")
    ),
    Sorted = Table.Sort(Filtered, {{"Date modified", Order.Descending}}),
    Newest = Table.First(Sorted),
    Wb = Excel.Workbook(Newest[Content], null, true),
    Sheet = Wb{0}[Data],
    Promoted = Table.PromoteHeaders(Table.Skip(Sheet, 9), [PromoteAllScalars=true]),
    Trimmed = Table.SelectRows(Promoted, each [Matrícula] <> null and [Matrícula] <> ""),
    Cols = Table.SelectColumns(Trimmed,
        {"Matrícula","Propietario","Marca","Modelo","Versión","Usuario: Nombre Completo",
         "Usuario: Nivel","Usuario: Centro Coste","Contrato: F.Inicio Contrato",
         "Contrato: F.Fin Contrato","Consumo Fabricante","Tipo Combustible",
         "Contrato: Meses Contratados","Contrato: Km Contratados","Cuota Total Mensual"})
in
    Cols
```

4. Done → name **`VendorFleet`**
5. Close & Load To… → Table → **Existing worksheet** → `flota!$B$10` (overwrites old fleet)
   ⚠ When prompted "the data already exists there, replace?" → **Yes**

Refresh confirms vendor file is read correctly.

---

## STEP 4 — Query 3: FleetConsumption (replaces the consumo pivot)

1. Data → Get Data → **From Other Sources** → **Blank Query**
2. Advanced Editor:

```m
let
    Source = VendorFleet,
    Cleaned = Table.SelectRows(Source, each
        [Consumo Fabricante] <> null and [Consumo Fabricante] <> 0
    ),
    Categorized = Table.AddColumn(Cleaned, "Cat", each
        if Text.Contains(Text.Lower([Tipo Combustible] ?? ""), "diésel") or
           Text.Contains(Text.Lower([Tipo Combustible] ?? ""), "diesel") then "Diesel"
        else if Text.Contains(Text.Lower([Tipo Combustible] ?? ""), "gasolina") then "Gasolina"
        else null
    ),
    Filtered = Table.SelectRows(Categorized, each [Cat] <> null),
    Grouped = Table.Group(Filtered, {"Cat"}, {{"Average", each List.Average([Consumo Fabricante])}})
in
    Grouped
```

3. Done → name **`FleetConsumption`**
4. Close & Load To… → Table → New worksheet — call it **`Consumo_avg`**

After load, find which row Diesel and Gasolina ended up on — typically `B2` and `B3` of the loaded table. Note them.

---

## STEP 5 — Wire up row 145 of `Liquidaciones de combustible`

Open the `Liquidaciones de combustible` sheet. Replace the existing row 145 with these formulas (the data row directly under the header on row 144):

| Cell | Formula |
|------|---------|
| `A145` | `=EOMONTH(TODAY(),0)` *(end of current month — overridable by typing)* |
| `H145` | `=VLOOKUP("Gasóleo A",Prices!A:B,2,FALSE)` |
| `I145` | `=VLOOKUP("Diesel",Consumo_avg!A:B,2,FALSE)` |
| `J145` | `=H145` |
| `K145` | `=I145*J145/100` |
| `L145` | `=VLOOKUP("Gasolina 95 E5",Prices!A:B,2,FALSE)` |
| `M145` | `=VLOOKUP("Gasolina",Consumo_avg!A:B,2,FALSE)` |
| `N145` | `=L145` |
| `O145` | `=M145*N145/100` |

Format `A145` as date (`dd/mm/yyyy`), `H/L` as `0.0000`, `I/M` as `0.00`, `K/O` as `0.0000`.

---

## STEP 6 — Add the "Email" sheet

1. Insert a new sheet → name **`Email`**.
2. Cell-by-cell:

   | Cell | Value / Formula |
   |------|-----------------|
   | `A1` | `Dear colleagues,` |
   | `A2` | (leave empty) |
   | `A3` | `="Find below price for "&TEXT('Liquidaciones de combustible'!A145,"mmmm yyyy")&":"` |
   | `A4` | (leave empty) |
   | `A5` | `Fecha del cálculo` |
   | `H5` | `Gasóleo A` |
   | `I5` | `Consumo Gasoil l/100km` |
   | `J5` | `Precio €/l Gasoil` |
   | `K5` | `Tarifa €/km Gasoil` |
   | `L5` | `Gasolina 95` |
   | `M5` | `Consumo Gasolina 95 l/100km` |
   | `N5` | `Precio €/l Gasolina 95` |
   | `O5` | `Tarifa €/km Gasolina 95` |
   | `A6` | `=TEXT('Liquidaciones de combustible'!A145,"d/mm/yyyy")` |
   | `H6` | `=TEXT('Liquidaciones de combustible'!H145,"#,##0.0000")` |
   | `I6` | `=TEXT('Liquidaciones de combustible'!I145,"#,##0.00")` |
   | `J6` | `=TEXT('Liquidaciones de combustible'!J145,"#,##0.0000")` |
   | `K6` | `=TEXT('Liquidaciones de combustible'!K145,"#,##0.0000")` |
   | `L6` | `=TEXT('Liquidaciones de combustible'!L145,"#,##0.0000")` |
   | `M6` | `=TEXT('Liquidaciones de combustible'!M145,"#,##0.00")` |
   | `N6` | `=TEXT('Liquidaciones de combustible'!N145,"#,##0.0000")` |
   | `O6` | `=TEXT('Liquidaciones de combustible'!O145,"#,##0.0000")` |
   | `A8` | `=TEXT('Liquidaciones de combustible'!K145,"#,##0.0000")&CHAR(9)&"Diesel"` |
   | `A9` | `=TEXT('Liquidaciones de combustible'!O145,"#,##0.0000")&CHAR(9)&"Gasoline"` |
   | `A11` | `Thanks` |

The user copies range `A1:O11` and pastes into Outlook — the tab characters become a table.

---

## STEP 7 — VBA "Close month" macro

1. Press **Alt+F11** to open the VBA editor
2. Insert → Module
3. Paste:

```vba
Sub CloseMonth()
    Const SHEET As String = "Liquidaciones de combustible"
    Const HEADER_TEXT As String = "Fecha del cálculo"
    Const CURRENT_ROW As Long = 145
    Const HEADER_ROW As Long = 144
    Const FIRST_COL As Long = 1   ' A
    Const LAST_COL As Long = 15   ' O

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET)

    ' Sanity check the header is where we expect.
    If Trim(CStr(ws.Cells(HEADER_ROW, 1).Value)) <> HEADER_TEXT Then
        MsgBox "Header '" & HEADER_TEXT & "' not found on row " & HEADER_ROW & _
               ". Aborting.", vbExclamation
        Exit Sub
    End If

    ' 1. Snapshot the current calculated values from row 145.
    Dim vals() As Variant
    ReDim vals(FIRST_COL To LAST_COL)
    Dim c As Long
    For c = FIRST_COL To LAST_COL
        vals(c) = ws.Cells(CURRENT_ROW, c).Value
    Next c

    ' 2. Insert a new blank row above the header. Header shifts to row 145,
    '    formulas on the old row 145 shift to row 146.
    ws.Rows(HEADER_ROW).EntireRow.Insert Shift:=xlDown

    ' 3. Write the snapshot values into the newly-created row (still HEADER_ROW number).
    For c = FIRST_COL To LAST_COL
        ws.Cells(HEADER_ROW, c).Value = vals(c)
    Next c

    ' 4. Done. Tell the user.
    MsgBox "Month closed. Refresh All to populate the new current row.", vbInformation
End Sub
```

4. Close VBA editor.
5. On the `Liquidaciones de combustible` sheet: Insert → Shapes → pick a button shape → label it **"Close month"** → right-click → Assign Macro → `CloseMonth`.
6. **File → Save As → choose "Excel Macro-Enabled Workbook (*.xlsm)"** to preserve the macro.

---

## Monthly procedure

```
1. Vendor emails Flota Viva Covestro <date>.xlsx
2. User saves it into the folder defined in Settings!B1
3. User opens Liquidaciones.xlsm
4. Data → Refresh All
   → all queries refresh; row 145 auto-populates
5. User opens the Email sheet → Ctrl+A → Ctrl+C → paste in Outlook
6. After sending, click the "Close month" button on the Liquidaciones sheet
   → row 145's values become history above row 144
   → row 145's formulas remain, ready for next month
```

Total monthly clicks: **5** (Refresh, copy email, paste in Outlook, Send, Close month).

---

## Troubleshooting

- **"FolderPath" name not found** — Settings sheet missing or the named range wasn't created. Re-do Step 1.
- **WebPrices fails with credential prompt** — set Anonymous in Data → Get Data → Data Source Settings.
- **VendorFleet says no rows** — folder path wrong, or no `Flota Viva Covestro*.xlsx` in the folder, or there's a `~$` lock file from an open file: close any open vendor file.
- **VLOOKUPs in row 145 return `#N/A`** — the labels in `Prices` / `Consumo_avg` shifted. Check that `Prices` column A still has "Gasolina 95 E5" and "Gasóleo A", and `Consumo_avg` column A has "Diesel" and "Gasolina".
- **Pivot on `consumo` sheet** — leave it alone; nothing references it anymore. Can be deleted later if you want.
