using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

public class WorkbookService : IDisposable {

  private static readonly List<string> requiredSheetNames = new List<string> {
        "INPUT DATA",
        "1ST",
        "2ND",
        "Final Semestral Grade"
    };

  public string FilePath { get; private set; }

  public List<string> WeightedScores { get; private set; } = new List<string>();

  public List<string> WrittenWorks { get; private set; } = new List<string>();

  public List<string> PerformanceTasks { get; private set; } = new List<string>();

  public string MissingSheets { get; set; }

  public string Exam { get; private set; }

  public bool HasFileLoaded => !string.IsNullOrWhiteSpace(FilePath);

  public bool IsFileECR() {

    using (var doc = SpreadsheetDocument.Open(FilePath, false)) {
      var wbPart = doc.WorkbookPart;
      var sheetNames = wbPart.Workbook.Sheets
          .OfType<Sheet>()
          .Select(s => s.Name)
          .ToList();

      MissingSheets = string.Join(
          ", ",
          requiredSheetNames
          .Where(req => !sheetNames.Contains(req))
          .ToList()
      );

      // Check if all required sheets are present
      return requiredSheetNames.All(req => sheetNames.Contains(req));
    }
  }



  public void LoadWorkbook(string path) {
    FilePath = path;
  }

  public (List<string>maleNames, List<string> femaleNames) ReadAllNames() {
    using (SpreadsheetDocument doc = SpreadsheetDocument.Open(FilePath, false)) {
      var maleNames = ReadNames(doc, true);
      var femaleNames = ReadNames(doc, false);
      ReadHighestPossibleScores(doc);
      return (maleNames, femaleNames);
    }
  }

  public List<string> ReadNames(SpreadsheetDocument doc, bool male = true) {
    string columnLetter = "B";
    uint startRow = male ? 13u : 64u;
    uint endRow = male ? 37u : 88u;

    var values = new List<string>();

    WorkbookPart wbPart = doc.WorkbookPart;
    Sheet sheet = wbPart.Workbook.Sheets
        .OfType<Sheet>()
        .FirstOrDefault(s => s.Name == requiredSheetNames[0]);

    if (sheet == null) return values;

    WorksheetPart wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    SheetData sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();

    for (uint rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
      string cellRef = $"{columnLetter}{rowIndex}";
      Cell cell = sheetData.Descendants<Cell>()
                           .FirstOrDefault(c => c.CellReference == cellRef);

      string val = GetCellValue(doc, cell);
      if (!string.IsNullOrWhiteSpace(val)) {
        values.Add(val);
      }
    }

    return values;
  }

  public List<string> AppendAndSortNames(string newValue, bool male = true) {
    string columnLetter = "B";
    uint startRow = male ? 13u : 64u;
    uint endRow = male ? 37u : 88u;

    using (SpreadsheetDocument doc = SpreadsheetDocument.Open(FilePath, true)) {
      WorkbookPart wbPart = doc.WorkbookPart;
      Sheet sheet = wbPart.Workbook.Sheets
          .OfType<Sheet>()
          .FirstOrDefault(s => s.Name == requiredSheetNames[0]);

      if (sheet == null) return new List<string>();

      WorksheetPart wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
      SheetData sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();

      // 1️⃣ Read existing values
      List<string> values = new List<string>();
      for (uint rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
        string cellRef = $"{columnLetter}{rowIndex}";
        Cell cell = sheetData.Descendants<Cell>()
                             .FirstOrDefault(c => c.CellReference == cellRef);
        string val = GetCellValue(doc, cell);
        if (!string.IsNullOrWhiteSpace(val))
          values.Add(val);
      }

      // 2️⃣ Append new value
      if (!string.IsNullOrWhiteSpace(newValue))
        values.Add(newValue);

      // 3️⃣ Sort alphabetically
      values = values.OrderBy(v => v, StringComparer.OrdinalIgnoreCase).ToList();

      // 4️⃣ Find a style to reuse
      uint? styleIndex = sheetData.Descendants<Cell>()
          .FirstOrDefault(c => GetColumnName(c.CellReference) == columnLetter && c.StyleIndex != null)
          ?.StyleIndex;

      // 5️⃣ Write sorted values back
      uint writeRow = startRow;
      foreach (var val in values) {
        if (writeRow > endRow) break;

        Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == writeRow);
        if (row == null) {
          row = new Row() { RowIndex = writeRow };
          sheetData.Append(row);
        }

        string cellRef = $"{columnLetter}{writeRow}";
        Cell cell = row.Elements<Cell>()
                       .FirstOrDefault(c => c.CellReference == cellRef);

        if (cell == null) {
          cell = new Cell() {
            CellReference = cellRef,
            StyleIndex = styleIndex
          };
          Cell refCell = row.Elements<Cell>()
                            .FirstOrDefault(c => string.Compare(c.CellReference.Value, cellRef, true) > 0);
          row.InsertBefore(cell, refCell);
        }

        // ✅ Update value without replacing the cell object
        cell.CellValue = new CellValue(val);
        cell.DataType = CellValues.String;

        writeRow++;
      }

      // 6️⃣ Clear remaining cells
      for (; writeRow <= endRow; writeRow++) {
        Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == writeRow);
        if (row != null) {
          string cellRef = $"{columnLetter}{writeRow}";
          Cell cell = row.Elements<Cell>()
                         .FirstOrDefault(c => c.CellReference == cellRef);
          if (cell != null) {
            cell.CellValue = null;
            cell.DataType = null; // keep style, clear value
          }
        }
      }

      // 7️⃣ Force Excel to recalc formulas on open
      var calcProps = wbPart.Workbook.CalculationProperties;
      if (calcProps == null) {
        calcProps = new CalculationProperties();
        wbPart.Workbook.Append(calcProps);
      }
      calcProps.FullCalculationOnLoad = true;
      calcProps.ForceFullCalculation = true;
      calcProps.CalculationOnSave = true;

      wsPart.Worksheet.Save();
      wbPart.Workbook.Save();

      return ReadNames(doc, male);
    }
  }

  public void ReadHighestPossibleScores(SpreadsheetDocument doc, bool sem1 = true) {
    uint fixedRow = 11u;
    WorkbookPart wbPart = doc.WorkbookPart;
    Sheet sheet = wbPart.Workbook.Sheets
        .OfType<Sheet>()
        .FirstOrDefault(s => s.Name == (sem1 ? requiredSheetNames[1] : requiredSheetNames[2]));

    if (sheet == null) return;

    WorksheetPart wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    SheetData sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();

    int startIdx, endIdx;
    string[,] Cols = {
                { "F", "O" },
                { "S", "AB" }
            };
    Cell cell;
    string val;

    for (int i = 0; i < 2; i++) {
      startIdx = ColNameToNumber(Cols[i, 0]);
      endIdx = ColNameToNumber(Cols[i, 1]);

      for (int colIdx = startIdx; colIdx <= endIdx; colIdx++) {
        string colName = ColNumberToName(colIdx);
        string cellRef = $"{colName}{fixedRow}";

        cell = sheetData.Descendants<Cell>()
                             .FirstOrDefault(c => c.CellReference == cellRef);

        val = GetCellValue(doc, cell); // ← your existing helper
        //if (!string.IsNullOrWhiteSpace(val)) {
          if (i == 0) {
            WrittenWorks.Add(val);
          } else if (i == 1) {
            PerformanceTasks.Add(val);
          }
        //}
      }
    }

    cell = sheetData.Descendants<Cell>()
                             .FirstOrDefault(c => c.CellReference == "AF11");
    Exam = GetCellValue(doc, cell);

    cell = sheetData.Descendants<Cell>()
                             .FirstOrDefault(c => c.CellReference == "R11");
    WeightedScores.Add(GetCellValue(doc, cell));

    cell = sheetData.Descendants<Cell>()
                             .FirstOrDefault(c => c.CellReference == "AE11");
    WeightedScores.Add(GetCellValue(doc, cell));

    cell = sheetData.Descendants<Cell>()
                             .FirstOrDefault(c => c.CellReference == "AH11");
    WeightedScores.Add(GetCellValue(doc, cell));


  }

  public (List<string> writtenWorks, List<string> performanceTasks) ReadStudentScores(
    uint row,
    bool sem1 = true
  ) {

    var values1 = new List<string>();
    var values2 = new List<string>();

    using (SpreadsheetDocument doc = SpreadsheetDocument.Open(FilePath, false)) {
      WorkbookPart wbPart = doc.WorkbookPart;
      Sheet sheet = wbPart.Workbook.Sheets
          .OfType<Sheet>()
          .FirstOrDefault(s => s.Name == (sem1 ? requiredSheetNames[1] : requiredSheetNames[2]));

      if (sheet == null) return (values1, values2);

      WorksheetPart wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
      SheetData sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();

      int startIdx, endIdx;
      string[,] Cols = {
        { "F", "O" },
        { "S", "AB" }
      };
      Cell cell;
      string val;

      for (int i = 0; i < 2; i++) {
        startIdx = ColNameToNumber(Cols[i, 0]);
        endIdx = ColNameToNumber(Cols[i, 1]);

        for (int colIdx = startIdx; colIdx <= endIdx; colIdx++) {
          string colName = ColNumberToName(colIdx);
          string cellRef = $"{colName}{row}";

          cell = sheetData.Descendants<Cell>()
                               .FirstOrDefault(c => c.CellReference == cellRef);

          val = GetCellValue(doc, cell);
          if (i == 0) {
            values1.Add(val);
          } else {
            values2.Add(val);
          }

        }
      }

      return (values1, values2);
    }
  }

  public void AddHighestPossibleScore(string valueToInsert, bool sem1 = true, bool isWrittenWork = true) {
    uint fixedRow = 11u;

    using (SpreadsheetDocument doc = SpreadsheetDocument.Open(FilePath, true)) {
      WorkbookPart wbPart = doc.WorkbookPart;
      Sheet sheet = wbPart.Workbook.Sheets
          .OfType<Sheet>()
          .FirstOrDefault(s => s.Name == (sem1 ? requiredSheetNames[1] : requiredSheetNames[2]));

      if (sheet == null) return;

      WorksheetPart wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
      SheetData sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();

      string[,] Cols = {
                { "F", "O" },  // WrittenWorks
                { "S", "AB" }  // PerformanceTasks
            };

      int rangeIndex = isWrittenWork ? 0 : 1;
      int startIdx = ColNameToNumber(Cols[rangeIndex, 0]);
      int endIdx = ColNameToNumber(Cols[rangeIndex, 1]);

      for (int colIdx = startIdx; colIdx <= endIdx; colIdx++) {
        string colName = ColNumberToName(colIdx);
        string cellRef = $"{colName}{fixedRow}";

        Cell cell = sheetData.Descendants<Cell>()
                             .FirstOrDefault(c => c.CellReference == cellRef);

        string val = GetCellValue(doc, cell);
        if (string.IsNullOrWhiteSpace(val)) {

          cell.CellValue = new CellValue(val);
          cell.DataType = CellValues.String;

          var calcProps = wbPart.Workbook.CalculationProperties;
          if (calcProps == null) {
            calcProps = new CalculationProperties();
            wbPart.Workbook.Append(calcProps);
          }
          calcProps.FullCalculationOnLoad = true;
          calcProps.ForceFullCalculation = true;
          calcProps.CalculationOnSave = true;

          wsPart.Worksheet.Save();
          wbPart.Workbook.Save();

          if (isWrittenWork)
            WrittenWorks.Add(valueToInsert);
          else
            PerformanceTasks.Add(valueToInsert);
          return; // Inserted, exit
        }
      }

      // All cells are full — do nothing
    }
  }

  public void SetStudentScore(
      uint row,
      string col,
      string valueToInsert,
      bool sem1 = true,
      bool isWrittenWork = true
  ) {

    using (SpreadsheetDocument doc = SpreadsheetDocument.Open(FilePath, true)) {
      WorkbookPart wbPart = doc.WorkbookPart;
      Sheet sheet = wbPart.Workbook.Sheets
          .OfType<Sheet>()
          .FirstOrDefault(s => s.Name == (sem1 ? requiredSheetNames[1] : requiredSheetNames[2]));

      if (sheet == null) return;

      WorksheetPart wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
      SheetData sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();

      string[,] Cols = {
                { "F", "O" },  // WrittenWorks
                { "S", "AB" }  // PerformanceTasks
            };

      int rangeIndex = isWrittenWork ? 0 : 1;
      int startIdx = ColNameToNumber(Cols[rangeIndex, 0]);
      int endIdx = ColNameToNumber(Cols[rangeIndex, 1]);

      for (int colIdx = startIdx; colIdx <= endIdx; colIdx++) {
        string colName = ColNumberToName(colIdx);
        string cellRef = $"{colName}{row}";

        Cell cell = sheetData.Descendants<Cell>()
                             .FirstOrDefault(c => c.CellReference == cellRef);

        string val = GetCellValue(doc, cell);
        if (string.IsNullOrWhiteSpace(val)) {
          InsertCellValue(sheetData, cellRef, valueToInsert);
          wsPart.Worksheet.Save();
          if (isWrittenWork)
            WrittenWorks.Add(valueToInsert);
          else
            PerformanceTasks.Add(valueToInsert);
          return; // Inserted, exit
        }
      }

      // All cells are full — do nothing
    }
  }

  private int ColNameToNumber(string colName) {
    int sum = 0;
    foreach (char c in colName.ToUpper()) {
      sum *= 26;
      sum += (c - 'A' + 1);
    }
    return sum;
  }

  private string ColNumberToName(int colNumber) {
    string colName = "";
    while (colNumber > 0) {
      int rem = (colNumber - 1) % 26;
      colName = (char)('A' + rem) + colName;
      colNumber = (colNumber - 1) / 26;
    }
    return colName;
  }

  // Helpers
  private static string GetColumnName(string cellReference) =>
      new string(cellReference.Where(char.IsLetter).ToArray());

  private static string GetCellValue(SpreadsheetDocument doc, Cell cell) {
    if (cell == null || cell.CellValue == null)
      return "";

    string value = cell.CellValue.InnerText;

    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
      return doc.WorkbookPart.SharedStringTablePart.SharedStringTable
                .Elements<SharedStringItem>()
                .ElementAt(int.Parse(value))
                .InnerText;
    }
    return value;
  }


  private void InsertCellValue(SheetData sheetData, string cellReference, string value) {
    string rowNumber = new string(cellReference.Where(char.IsDigit).ToArray());
    Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == uint.Parse(rowNumber));

    if (row == null) {
      row = new Row() { RowIndex = uint.Parse(rowNumber) };
      sheetData.Append(row);
    }

    Cell cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference == cellReference);
    if (cell == null) {
      cell = new Cell() { CellReference = cellReference };
      row.Append(cell);
    }

    cell.DataType = CellValues.String;
    cell.CellValue = new CellValue(value);
  }

  public void Dispose() {
    // Nothing to dispose unless you keep a persistent open document
  }
}
