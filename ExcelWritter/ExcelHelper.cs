using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;

namespace ExcelWritter
{
    public class ExcelHelper
    {

        public void Write(WorkBookDefinition conf)
        {
            using (var fs = File.Open(conf.FileName,FileMode.OpenOrCreate))
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(fs, SpreadsheetDocumentType.Workbook))
                {

                    var workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    workbookPart.Workbook.AppendChild(new Sheets());

                    var stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                    stylePart.Stylesheet = GenerateStyleSheet();
                    stylePart.Stylesheet.Save();

                    var worksheetPart = InsertWorksheet(document.WorkbookPart, conf.SheetName, conf.Rows.Length, conf);

                    var rowIdx = 1;
                    var col = 'A';
                    foreach (var colName in conf.ColumnsNames)
                    {
                        WriteCellContent(col.ToString(), rowIdx, colName, worksheetPart);
                        col++;
                    }                    

                    rowIdx++;
                    foreach(var row in conf.Rows)
                    {
                        col = 'A';
                        foreach (var colVal in row)
                        {
                            WriteCellContent(col.ToString(), rowIdx, colVal, worksheetPart);
                            col++;
                        }
                        rowIdx++;
                    }
                    worksheetPart.Worksheet.Save();
                }
            }
           
        }

        private Stylesheet GenerateStyleSheet()
        {
            Fonts fonts = new Fonts(
                new Font( // Index 0 - default
                    new FontSize() { Val = 10 },
                    new Color() { Rgb = "000000" }
                ),
                   new Font( // Index 1 - header
                    new FontSize() { Val = 10 },
                    new Bold(),
                    new Color() { Rgb = "FFFFFF" }
                ),
                new Font( // Index 2 - row
                    new FontSize() { Val = 10 },
                    new Color() { Rgb = "2B4270" }
                )
             );

            Fills fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = "2B4270" }) { PatternType = PatternValues.Solid }), // Index 2 - header
                    new Fill(new PatternFill(new ForegroundColor { Rgb = "EBF3F5" }) { PatternType = PatternValues.Solid }), // Index 3 - pair
                    new Fill(new PatternFill(new ForegroundColor { Rgb = "FFFFFF" }) { PatternType = PatternValues.Solid }) // Index 4 - unpair
                );

            Borders borders = new Borders(
                    new Border(), // index 0 default
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                );

            CellFormats cellFormats = new CellFormats(
                    new CellFormat(), // default
                    new CellFormat
                    {
                        FontId = 1,
                        FillId = 2,
                        BorderId = 1,
                        ApplyBorder = true,
                        Alignment = new Alignment
                        {
                            Horizontal = HorizontalAlignmentValues.Center,
                            Vertical = VerticalAlignmentValues.Center
                        }
                    }, // header
                    new CellFormat { FontId = 2, FillId = 3, BorderId = 1, ApplyFill = true }, // pair
                    new CellFormat { FontId = 2, FillId = 4, BorderId = 1, ApplyFill = true } // unpair
                );

            Stylesheet styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);

            return styleSheet;
        }

        private void WriteCellContent(string column, int row, string text, WorksheetPart worksheetPart)
        {
            Cell cell = InsertCellInWorksheet(column, (uint)row, worksheetPart);
            cell.CellValue = new CellValue(text);
            cell.DataType = CellValues.String;

            cell.StyleIndex = row == 1
                ? 1U
                : (row % 2 == 0
                ? 2U
                : 3U);
        }

        // Given a WorkbookPart, inserts a new worksheet.
        private WorksheetPart InsertWorksheet(WorkbookPart workbookPart, string sheetName, int linesCount, WorkBookDefinition conf)
        {
            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            Columns columns = new Columns();
            columns.Append(conf.ColumnsNames.Select((c,idx) => new Column() { Min = (uint)(idx+1), Max = (uint)(idx + 1),
                Width = 10, CustomWidth = true }));

            newWorksheetPart.Worksheet = new Worksheet();

            newWorksheetPart.Worksheet.Append(columns);

            newWorksheetPart.Worksheet.AppendChild(new SheetData());

            var endColLetter = (char)('A' + conf.ColumnsNames.Length);
            AutoFilter autoFilter = new AutoFilter() { Reference = "A1:" + (endColLetter.ToString()) + linesCount };
            newWorksheetPart.Worksheet.AppendChild(autoFilter);

            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        private Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

    }

    public class WorkBookDefinition
    {
        public string FileName { get; set; }
        public string SheetName { get; set; }
        public string[] ColumnsNames { get; set; }
        public string[][] Rows { get; set; }
    }

}
