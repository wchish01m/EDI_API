using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualBasic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using oNonVisDrawProp = DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties;
using oNonVisDrawPicProp = DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureDrawingProperties;
using oNonVisPicProp = DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties;
using oStretch = DocumentFormat.OpenXml.Drawing.Stretch;
using oPicLocks = DocumentFormat.OpenXml.Drawing.PictureLocks;
using oBlipFill = DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill;
using oBlip = DocumentFormat.OpenXml.Drawing.Blip;
using oTransoform2D = DocumentFormat.OpenXml.Drawing.Transform2D;
using oOffset = DocumentFormat.OpenXml.Drawing.Offset;
using oExtents = DocumentFormat.OpenXml.Drawing.Extents;
using oShapeProperties = DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties;
using oPresetGeometry = DocumentFormat.OpenXml.Drawing.PresetGeometry;
using oPicture = DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture;
using oPosition = DocumentFormat.OpenXml.Drawing.Spreadsheet.Position;
using System.Text.RegularExpressions;
using EDI_API.Excel_Creator;

// Imports eCell_Font
// Imports eCell_FillColor
// Imports eCell_Borders
// Imports eCell_TextFormat
// Imports eCell_Style
// Imports eCell
// Imports eColumn
// Imports eSheet

public class Excel_Builder
{
    public static void create_Document(string filepath, List<eSheet> sSheets)
    {
        // Create a spreadsheet document by supplying the filepath.
        // By default, AutoSave = true, Editable = true, and Type = xlsx.
        // Dim start As Date
        // Dim finish As Date
        SpreadsheetDocument ssDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

        // Add a WorkbookPart to the document.
        WorkbookPart wbPart = ssDocument.AddWorkbookPart();
        wbPart.Workbook = new Workbook();
        // start = Now
        WorkbookStylesPart wbStyles = wbPart.AddNewPart<WorkbookStylesPart>();
        wbStyles.Stylesheet = createStyleSheet(sSheets[0].fonts, sSheets[0].fills, sSheets[0].borders, sSheets[0].textFormats, sSheets[0].styles);
        wbStyles.Stylesheet.Save();
        // finish = Now

        // My.Computer.FileSystem.WriteAllText("\\srv04\errorLogs$\openXML_Log.txt",
        // filepath & vbCrLf & "Stylesheet Creation, Start= " & start & ", End= " & finish & ", seconds= " & DateDiff(DateInterval.Second, start, finish) & vbCrLf, True)

        // Add Sheets to the Workbook.
        Sheets wsheets = ssDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

        for (uint s = 0; s <= sSheets.Count - 1; s++)
        {
            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
            Worksheet work_Sheet = new Worksheet();
            SheetData sData = new SheetData();

            Columns columns = new Columns();
            // start = Now
            for (int x = 0; x <= sSheets[(int)s].columns.Count - 1; x++)
                columns.Append(createColumnData(sSheets[(int)s].columns[x].Index, sSheets[(int)s].columns[x].Width, true));
            // finish = Now

            // My.Computer.FileSystem.WriteAllText("\\srv04\errorLogs$\openXML_Log.txt",
            // "Column Creation, Start= " & start & ", End= " & finish & ", seconds= " & DateDiff(DateInterval.Second, start, finish) & vbCrLf, True)

            work_Sheet.Append(columns);
            work_Sheet.Append(sData);

            wsPart.Worksheet = work_Sheet;

            // Append a new worksheet and associate it with the workbook.
            Sheet wsheet = new Sheet();
            wsheet.Id = ssDocument.WorkbookPart.GetIdOfPart(wsPart);
            wsheet.SheetId = s + 1;
            string sName = sSheets[(int)s].name;

            if (sName == "DEFAULT")
                sName = "Sheet" + (s + 1);

            wsheet.Name = sName;

            wsheets.Append(wsheet);


            // start = Now
            // For x As Integer = 0 To sSheets(s).cells.Count - 1
            InsertText(ssDocument, wsPart, sSheets[(int)s]);
            // MsgBox(eCells(x).merged)

            // Next
            // finish = Now

            // My.Computer.FileSystem.WriteAllText("\\srv04\errorLogs$\openXML_Log.txt",
            // "Cells Creation, Start= " & start & ", End= " & finish & ", seconds= " & DateDiff(DateInterval.Second, start, finish) & vbCrLf, True)

            // start = Now
            if (sSheets[(int)s].images.Count > 0)
            {
                for (int x = 0; x <= sSheets[(int)s].images.Count - 1; x++)
                    addImageToSheet(wsPart, wsPart.Worksheet, sSheets[(int)s].images[x]);
            }
            // finish = Now

            // My.Computer.FileSystem.WriteAllText("\\srv04\errorLogs$\openXML_Log.txt",
            // "Image Creation, Start= " & start & ", End= " & finish & ", seconds= " & DateDiff(DateInterval.Second, start, finish) & vbCrLf & vbCrLf, True)

            wsPart.Worksheet.Save();

            wbPart.Workbook.Save();
        }

        // Close the document.
        ssDocument.Close();
    }

    public static void InsertText(SpreadsheetDocument sDocument, WorksheetPart wsPart, eSheet sSheet)
    {

        // Imports (spreadSheet)
        // Get the SharedStringTablePart. If it does not exist, create a new one.
        SharedStringTablePart shareStringPart;

        if (sDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            shareStringPart = sDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
        else
            shareStringPart = sDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();

        for (int x = 0; x <= sSheet.cells.Count - 1; x++)
        {
            eCell dataCell = sSheet.cells[x];

            // Insert the text into the SharedStringTablePart.
            int index = InsertSharedStringItem(dataCell.value, shareStringPart);

            // Insert cell into the new worksheet.
            Cell sCell = InsertCellInWorksheet(dataCell.column, (uint)dataCell.row, wsPart);

            // To check if the cell has a formula or not
            if (dataCell.formula == "")
            {
                if (dataCell.cellType == "DEFAULT")
                {
                    int n;
                    bool isNum = int.TryParse(dataCell.value, out n);

                    DateTime da;
                    bool isDate = DateTime.TryParse(dataCell.value, out da);
                    // Assign the cell value
                    if (isNum == true)
                    {
                        sCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        sCell.CellValue = new CellValue(dataCell.value);
                    }
                    else if (isDate == true)
                    {
                        double d = DateTime.Parse(dataCell.value).ToOADate();
                        sCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        sCell.CellValue = new CellValue(d);
                    }
                    else
                    {
                        sCell.CellValue = new CellValue(index.ToString());
                        sCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    }
                }
                else
                    switch (dataCell.cellType)
                    {
                        case "NUMBER":
                            {
                                sCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                                sCell.CellValue = new CellValue(dataCell.value);
                                break;
                            }

                        case "DATE":
                            {
                                sCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                                double d = DateTime.Parse(dataCell.value).ToOADate();
                                sCell.CellValue = new CellValue(d);
                                break;
                            }

                        case "TEXT":
                            {
                                sCell.CellValue = new CellValue(index.ToString());
                                sCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                                break;
                            }
                    }
            }
            else
            {
                CellFormula cFormula = new CellFormula();
                cFormula.Text = dataCell.formula;
                sCell.CellFormula = cFormula;
            }

            sCell.StyleIndex = dataCell.styleID;
        }

        // Save the new worksheet.
        wsPart.Worksheet.Save();

        for (int x = 0; x <= sSheet.cells.Count - 1; x++)
        {
            if (sSheet.cells[x].merged == true)
            {
                styleMergedCells(sDocument, sSheet.cells[x], wsPart);
                MergeTwoCells(sDocument, sSheet.name, sSheet.cells[x].address, sSheet.cells[x].mergedAddress, (int)sSheet.cells[x].styleID);            //!!POSSIBLE!!ERROR!!//
            }
        }
    }

    public static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
    {
        // If the part does not contain a SharedStringTable, create one.
        if ((shareStringPart.SharedStringTable == null))
            shareStringPart.SharedStringTable = new SharedStringTable();

        int i = 0;

        // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
        {
            if ((item.InnerText == text))
                return i;
            i = (i + 1);
        }

        // The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
        shareStringPart.SharedStringTable.Save();

        return i;
    }

    public static void styleMergedCells(SpreadsheetDocument ssDocument, eCell styleCell, WorksheetPart wsPart)
    {
        for (int x = styleCell.row; x <= styleCell.mergedRow; x++)
        {
            for (int y = styleCell.colNum; y <= styleCell.mergedColNum; y++)
            {
                if (x != styleCell.row | y != styleCell.colNum)
                {
                    eCell mCell = new eCell();
                    mCell.value = "";
                    mCell.row = x;
                    mCell.column = y.ToString();
                    mCell.styleID = styleCell.styleID;
                    // Call InsertText(ssDocument, wsPart, mCell)

                    // Insert the text into the SharedStringTablePart.
                    // Dim index As Integer = InsertSharedStringItem(mCell.value, shareStringPart)

                    // Insert cell into the new worksheet.
                    Cell sCell = InsertCellInWorksheet(mCell.column, (uint)mCell.row, wsPart);
                    sCell.StyleIndex = mCell.styleID;
                }
            }
        }

        // Save the new worksheet.
        wsPart.Worksheet.Save();
    }

    public static string ConvertToLetter(int columnNum)
    {
        int a;
        int b;
        string columnLetters = "";
        a = columnNum;
        while (columnNum > 0)
        {
            a = (int)((columnNum - 1) / (double)26);
            b = (columnNum - 1) % 26;
            columnLetters = Strings.Chr(b + 65) + columnLetters;
            columnNum = a;
        }

        return columnLetters;
    }

    public static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
    {
        Worksheet worksheet = worksheetPart.Worksheet;
        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
        string cellReference = (columnName + rowIndex.ToString());

        // If the worksheet does not contain a row with the specified row index, insert one.
        Row row;
        if (sheetData.Elements<Row>().Where(r => r.RowIndex.Value == rowIndex).Count() != 0)
            row = sheetData.Elements<Row>().Where(r => r.RowIndex.Value == rowIndex).First();
        else
        {
            row = new Row();
            row.RowIndex = rowIndex;
            sheetData.Append(row);
        }

        // If there is not a cell with the specified column name, insert one.  
        if ((row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex.ToString()).Count() > 0))
            return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
        else
        {
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell refCell = null/* TODO Change to default(_) if this is not a reference type */;
            foreach (Cell cell in row.Elements<Cell>())
            {
                if ((string.Compare(cell.CellReference.Value, cellReference, true) > 0))
                {
                    refCell = cell;
                    break;
                }
            }

            Cell newCell = new Cell();
            newCell.CellReference = cellReference;

            row.InsertBefore(newCell, refCell);
            worksheet.Save();

            return newCell;
        }
    }

    public static void MergeTwoCells(SpreadsheetDocument sDocument, string sheetName, string cell1Name, string cell2Name, int styleID)
    {
        // Open the document for editing.
        // Dim document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)

        // Using (sDocument)
        Worksheet wSheet = GetWorksheet(sDocument, sheetName);
        if (((wSheet == null) || (string.IsNullOrEmpty(cell1Name) || string.IsNullOrEmpty(cell2Name))))
            return;

        // Verify if the specified cells exist, and if they do not exist, create them.
        CreateSpreadsheetCellIfNotExist(wSheet, cell1Name, styleID);
        CreateSpreadsheetCellIfNotExist(wSheet, cell2Name, styleID);

        MergeCells mergeCells;
        if ((wSheet.Elements<MergeCells>().Count() > 0))
            mergeCells = wSheet.Elements<MergeCells>().First();
        else
        {
            mergeCells = new MergeCells();
            // Insert a MergeCells object into the specified position.
            if ((wSheet.Elements<CustomSheetView>().Count() > 0))
                wSheet.InsertAfter(mergeCells, wSheet.Elements<CustomSheetView>().First());
            else if ((wSheet.Elements<DataConsolidate>().Count() > 0))
                wSheet.InsertAfter(mergeCells, wSheet.Elements<DataConsolidate>().First());
            else if ((wSheet.Elements<SortState>().Count() > 0))
                wSheet.InsertAfter(mergeCells, wSheet.Elements<SortState>().First());
            else if ((wSheet.Elements<AutoFilter>().Count() > 0))
                wSheet.InsertAfter(mergeCells, wSheet.Elements<AutoFilter>().First());
            else if ((wSheet.Elements<Scenarios>().Count() > 0))
                wSheet.InsertAfter(mergeCells, wSheet.Elements<Scenarios>().First());
            else if ((wSheet.Elements<ProtectedRanges>().Count() > 0))
                wSheet.InsertAfter(mergeCells, wSheet.Elements<ProtectedRanges>().First());
            else if ((wSheet.Elements<SheetProtection>().Count() > 0))
                wSheet.InsertAfter(mergeCells, wSheet.Elements<SheetProtection>().First());
            else if ((wSheet.Elements<SheetCalculationProperties>().Count() > 0))
                wSheet.InsertAfter(mergeCells, wSheet.Elements<SheetCalculationProperties>().First());
            else
                wSheet.InsertAfter(mergeCells, wSheet.Elements<SheetData>().First());
        }

        // Create the merged cell and append it to the MergeCells collection.
        MergeCell mergeCell = new MergeCell();
        mergeCell.Reference = new StringValue((cell1Name + (":" + cell2Name)));
        mergeCells.Append(mergeCell);

        wSheet.Save();
    }

    // Given a SpreadsheetDocument and a worksheet name, get the specified worksheet.
    public static Worksheet GetWorksheet(SpreadsheetDocument document, string worksheetName)
    {
        IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
        if (sheets.Count() == 0)
            // The specified worksheet does not exist.
            return null/* TODO Change to default(_) if this is not a reference type */;
        WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);

        return worksheetPart.Worksheet;
    }

    // Given a Worksheet and a cell name, verifies that the specified cell exists.
    // If it does not exist, creates a new cell.
    public static void CreateSpreadsheetCellIfNotExist(Worksheet worksheet, string cellName, int styleID)
    {
        string columnName = GetColumnName(cellName);
        uint rowIndex = GetRowIndex(cellName);

        IEnumerable<Row> rows = worksheet.Descendants<Row>().Where(r => r.RowIndex.Value.ToString() == rowIndex.ToString());

        // If the worksheet does not contain the specified row, create the specified row.
        // Create the specified cell in that row, and insert the row into the worksheet.
        if ((rows.Count() == 0))
        {
            Row row = new Row();
            row.RowIndex = new UInt32Value(rowIndex);

            Cell cell = new Cell();
            cell.CellReference = new StringValue(cellName);
            // cell.StyleIndex = styleID
            row.Append(cell);
            worksheet.Descendants<SheetData>().First().Append(row);
            worksheet.Save();
        }
        else
        {
            Row row = rows.First();
            IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellName);

            // If the row does not contain the specified cell, create the specified cell.
            if ((cells.Count() == 0))
            {
                Cell cell = new Cell();
                cell.CellReference = new StringValue(cellName);
                // cell.StyleIndex = styleID
                row.Append(cell);
                worksheet.Save();
            }
        }
    }

    // Given a cell name, parses the specified cell to get the column name.
    public static string GetColumnName(string cellName)
    {
        // Create a regular expression to match the column name portion of the cell name.
        Regex regex = new Regex("[A-Za-z]+");
        Match match = regex.Match(cellName);
        return match.Value;
    }

    // Given a cell name, parses the specified cell to get the row index.
    public static uint GetRowIndex(string cellName)
    {
        // Create a regular expression to match the row index portion the cell name.
        Regex regex = new Regex(@"\d+");
        Match match = regex.Match(cellName);
        return uint.Parse(match.Value);
    }


    public static Stylesheet createStyleSheet(List<eCell_Font> eFonts, List<eCell_FillColor> eFills, List<eCell_Borders> eBorders, List<eCell_TextFormat> eTextFormats, List<eCell_Style> eStyles)
    {
        Stylesheet newStyleSheet = new Stylesheet();


        // ##########################################################################################################################
        // ###
        // ###    Setting the fonts for the stylesheet for the worksheet
        // ###
        // ##########################################################################################################################
        Fonts fnts = new Fonts();

        for (int x = 0; x <= eFonts.Count - 1; x++)
        {
            DocumentFormat.OpenXml.Spreadsheet.Font fnt = new DocumentFormat.OpenXml.Spreadsheet.Font();
            FontName fntName = new FontName();
            fntName.Val = eFonts[x].name;
            FontSize fntSize = new FontSize();
            fntSize.Val = eFonts[x].size;
            fnt.Color = new DocumentFormat.OpenXml.Spreadsheet.Color();
            fnt.Color.Rgb = HexBinaryValue.FromString(eFonts[x].clr);
            if (eFonts[x].bold == true)
                fnt.Bold = new Bold();
            fnt.FontName = fntName;
            fnt.FontSize = fntSize;
            fnts.Append(fnt);
        }

        fnts.Count = System.Convert.ToUInt32(fnts.ChildElements.Count);


        // ##########################################################################################################################
        // ###
        // ###    Setting the cell colors for the stylehseet for the worksheet
        // ###
        // ##########################################################################################################################
        Fills cellFills = new Fills();

        for (int x = 0; x <= eFills.Count - 1; x++)
        {
            Fill cellFill = new Fill();
            PatternFill cellPatternFill = new PatternFill();
            switch (eFills[x].pattern)
            {
                case "NONE":
                    {
                        cellPatternFill.PatternType = PatternValues.None;
                        break;
                    }

                case "SOLID":
                    {
                        cellPatternFill.PatternType = PatternValues.Solid;
                        break;
                    }

                case "GRAY125":
                    {
                        cellPatternFill.PatternType = PatternValues.Gray125;
                        break;
                    }

                default:
                    {
                        cellPatternFill.PatternType = PatternValues.None;
                        break;
                    }
            }

            if (eFills[x].name != "NONE")
            {
                cellPatternFill.ForegroundColor = new ForegroundColor();
                cellPatternFill.ForegroundColor.Rgb = HexBinaryValue.FromString(eFills[x].clr);
                cellPatternFill.BackgroundColor = new BackgroundColor();
                cellPatternFill.BackgroundColor.Rgb = cellPatternFill.ForegroundColor.Rgb;
            }

            cellFill.PatternFill = cellPatternFill;
            cellFills.Append(cellFill);
        }

        cellFills.Count = System.Convert.ToUInt32(cellFills.ChildElements.Count);


        // ##########################################################################################################################
        // ###
        // ###    Setting the cell borders for the stylehseet for the worksheet
        // ###
        // ##########################################################################################################################

        Borders cellBorders = new Borders();

        for (int x = 0; x <= eBorders.Count - 1; x++)
        {
            Border cellBorder = new Border();

            if (eBorders[x].name != "DEFAULT")
            {
                cellBorder.LeftBorder = new LeftBorder();
                switch (eBorders[x].leftBorder)
                {
                    case "NONE":
                        {
                            cellBorder.LeftBorder.Style = BorderStyleValues.None;
                            break;
                        }

                    case "DASHDOT":
                        {
                            cellBorder.LeftBorder.Style = BorderStyleValues.DashDot;
                            break;
                        }

                    case "DASHDOTDOT":
                        {
                            cellBorder.LeftBorder.Style = BorderStyleValues.DashDotDot;
                            break;
                        }

                    case "DASHED":
                        {
                            cellBorder.LeftBorder.Style = BorderStyleValues.Dashed;
                            break;
                        }

                    case "DOTTED":
                        {
                            cellBorder.LeftBorder.Style = BorderStyleValues.Dotted;
                            break;
                        }

                    case "DOUBLE":
                        {
                            cellBorder.LeftBorder.Style = BorderStyleValues.Double;
                            break;
                        }

                    case "HAIR":
                        {
                            cellBorder.LeftBorder.Style = BorderStyleValues.Hair;
                            break;
                        }

                    case "MEDIUM":
                        {
                            cellBorder.LeftBorder.Style = BorderStyleValues.Medium;
                            break;
                        }

                    case "MEDIUMDASHDOT":
                        {
                            cellBorder.LeftBorder.Style = BorderStyleValues.MediumDashDot;
                            break;
                        }

                    case "MEDIUMDASHDOTDO":
                        {
                            cellBorder.LeftBorder.Style = BorderStyleValues.MediumDashDotDot;
                            break;
                        }

                    case "MEDIUMDASHED":
                        {
                            cellBorder.LeftBorder.Style = BorderStyleValues.MediumDashed;
                            break;
                        }

                    case "SLANTDASHDOT":
                        {
                            cellBorder.LeftBorder.Style = BorderStyleValues.SlantDashDot;
                            break;
                        }

                    case "THICK":
                        {
                            cellBorder.LeftBorder.Style = BorderStyleValues.Thick;
                            break;
                        }

                    case "THIN":
                        {
                            cellBorder.LeftBorder.Style = BorderStyleValues.Thin;
                            break;
                        }

                    default:
                        {
                            cellBorder.LeftBorder.Style = BorderStyleValues.None;
                            break;
                        }
                }

                cellBorder.RightBorder = new RightBorder();
                switch (eBorders[x].rightBorder)
                {
                    case "NONE":
                        {
                            cellBorder.RightBorder.Style = BorderStyleValues.None;
                            break;
                        }

                    case "DASHDOT":
                        {
                            cellBorder.RightBorder.Style = BorderStyleValues.DashDot;
                            break;
                        }

                    case "DASHDOTDOT":
                        {
                            cellBorder.RightBorder.Style = BorderStyleValues.DashDotDot;
                            break;
                        }

                    case "DASHED":
                        {
                            cellBorder.RightBorder.Style = BorderStyleValues.Dashed;
                            break;
                        }

                    case "DOTTED":
                        {
                            cellBorder.RightBorder.Style = BorderStyleValues.Dotted;
                            break;
                        }

                    case "DOUBLE":
                        {
                            cellBorder.RightBorder.Style = BorderStyleValues.Double;
                            break;
                        }

                    case "HAIR":
                        {
                            cellBorder.RightBorder.Style = BorderStyleValues.Hair;
                            break;
                        }

                    case "MEDIUM":
                        {
                            cellBorder.RightBorder.Style = BorderStyleValues.Medium;
                            break;
                        }

                    case "MEDIUMDASHDOT":
                        {
                            cellBorder.RightBorder.Style = BorderStyleValues.MediumDashDot;
                            break;
                        }

                    case "MEDIUMDASHDOTDO":
                        {
                            cellBorder.RightBorder.Style = BorderStyleValues.MediumDashDotDot;
                            break;
                        }

                    case "MEDIUMDASHED":
                        {
                            cellBorder.RightBorder.Style = BorderStyleValues.MediumDashed;
                            break;
                        }

                    case "SLANTDASHDOT":
                        {
                            cellBorder.RightBorder.Style = BorderStyleValues.SlantDashDot;
                            break;
                        }

                    case "THICK":
                        {
                            cellBorder.RightBorder.Style = BorderStyleValues.Thick;
                            break;
                        }

                    case "THIN":
                        {
                            cellBorder.RightBorder.Style = BorderStyleValues.Thin;
                            break;
                        }

                    default:
                        {
                            cellBorder.RightBorder.Style = BorderStyleValues.None;
                            break;
                        }
                }


                cellBorder.TopBorder = new TopBorder();
                switch (eBorders[x].topBorder)
                {
                    case "NONE":
                        {
                            cellBorder.TopBorder.Style = BorderStyleValues.None;
                            break;
                        }

                    case "DASHDOT":
                        {
                            cellBorder.TopBorder.Style = BorderStyleValues.DashDot;
                            break;
                        }

                    case "DASHDOTDOT":
                        {
                            cellBorder.TopBorder.Style = BorderStyleValues.DashDotDot;
                            break;
                        }

                    case "DASHED":
                        {
                            cellBorder.TopBorder.Style = BorderStyleValues.Dashed;
                            break;
                        }

                    case "DOTTED":
                        {
                            cellBorder.TopBorder.Style = BorderStyleValues.Dotted;
                            break;
                        }

                    case "DOUBLE":
                        {
                            cellBorder.TopBorder.Style = BorderStyleValues.Double;
                            break;
                        }

                    case "HAIR":
                        {
                            cellBorder.TopBorder.Style = BorderStyleValues.Hair;
                            break;
                        }

                    case "MEDIUM":
                        {
                            cellBorder.TopBorder.Style = BorderStyleValues.Medium;
                            break;
                        }

                    case "MEDIUMDASHDOT":
                        {
                            cellBorder.TopBorder.Style = BorderStyleValues.MediumDashDot;
                            break;
                        }

                    case "MEDIUMDASHDOTDO":
                        {
                            cellBorder.TopBorder.Style = BorderStyleValues.MediumDashDotDot;
                            break;
                        }

                    case "MEDIUMDASHED":
                        {
                            cellBorder.TopBorder.Style = BorderStyleValues.MediumDashed;
                            break;
                        }

                    case "SLANTDASHDOT":
                        {
                            cellBorder.TopBorder.Style = BorderStyleValues.SlantDashDot;
                            break;
                        }

                    case "THICK":
                        {
                            cellBorder.TopBorder.Style = BorderStyleValues.Thick;
                            break;
                        }

                    case "THIN":
                        {
                            cellBorder.TopBorder.Style = BorderStyleValues.Thin;
                            break;
                        }

                    default:
                        {
                            cellBorder.TopBorder.Style = BorderStyleValues.None;
                            break;
                        }
                }


                cellBorder.BottomBorder = new BottomBorder();
                switch (eBorders[x].bottomBorder)
                {
                    case "NONE":
                        {
                            cellBorder.BottomBorder.Style = BorderStyleValues.None;
                            break;
                        }

                    case "DASHDOT":
                        {
                            cellBorder.BottomBorder.Style = BorderStyleValues.DashDot;
                            break;
                        }

                    case "DASHDOTDOT":
                        {
                            cellBorder.BottomBorder.Style = BorderStyleValues.DashDotDot;
                            break;
                        }

                    case "DASHED":
                        {
                            cellBorder.BottomBorder.Style = BorderStyleValues.Dashed;
                            break;
                        }

                    case "DOTTED":
                        {
                            cellBorder.BottomBorder.Style = BorderStyleValues.Dotted;
                            break;
                        }

                    case "DOUBLE":
                        {
                            cellBorder.BottomBorder.Style = BorderStyleValues.Double;
                            break;
                        }

                    case "HAIR":
                        {
                            cellBorder.BottomBorder.Style = BorderStyleValues.Hair;
                            break;
                        }

                    case "MEDIUM":
                        {
                            cellBorder.BottomBorder.Style = BorderStyleValues.Medium;
                            break;
                        }

                    case "MEDIUMDASHDOT":
                        {
                            cellBorder.BottomBorder.Style = BorderStyleValues.MediumDashDot;
                            break;
                        }

                    case "MEDIUMDASHDOTDO":
                        {
                            cellBorder.BottomBorder.Style = BorderStyleValues.MediumDashDotDot;
                            break;
                        }

                    case "MEDIUMDASHED":
                        {
                            cellBorder.BottomBorder.Style = BorderStyleValues.MediumDashed;
                            break;
                        }

                    case "SLANTDASHDOT":
                        {
                            cellBorder.BottomBorder.Style = BorderStyleValues.SlantDashDot;
                            break;
                        }

                    case "THICK":
                        {
                            cellBorder.BottomBorder.Style = BorderStyleValues.Thick;
                            break;
                        }

                    case "THIN":
                        {
                            cellBorder.BottomBorder.Style = BorderStyleValues.Thin;
                            break;
                        }

                    default:
                        {
                            cellBorder.BottomBorder.Style = BorderStyleValues.None;
                            break;
                        }
                }


                cellBorder.DiagonalBorder = new DiagonalBorder();

                cellBorders.Append(cellBorder);
            }
            else
            {
                cellBorder.LeftBorder = new LeftBorder();
                cellBorder.RightBorder = new RightBorder();
                cellBorder.TopBorder = new TopBorder();
                cellBorder.BottomBorder = new BottomBorder();
                cellBorder.DiagonalBorder = new DiagonalBorder();
                cellBorders.Append(cellBorder);
            }
        }

        cellBorders.Count = System.Convert.ToUInt32(cellBorders.ChildElements.Count);


        // ##########################################################################################################################
        // ###
        // ###    Setting the number formats for the stylehseet for the worksheet
        // ###
        // ##########################################################################################################################

        NumberingFormats nFormats = new NumberingFormats();

        for (int x = 0; x <= eTextFormats.Count - 1; x++)
        {
            NumberingFormat nformat = new NumberingFormat();
            nformat.NumberFormatId = (uint)eTextFormats[x].indexID;
            nformat.FormatCode = eTextFormats[x].numFormat;
            nFormats.Append(nformat);
        }

        nFormats.Count = System.Convert.ToUInt32(nFormats.ChildElements.Count);


        // ##########################################################################################################################
        // ###
        // ###    Setting the cell formats for the stylehseet for the worksheet
        // ###
        // ##########################################################################################################################
        CellStyleFormats cStyleFormats = new CellStyleFormats();
        CellFormat cFormat = new CellFormat();
        cFormat.NumberFormatId = 0;
        cFormat.FontId = 0;
        cFormat.FillId = 0;
        cFormat.BorderId = 0;

        cStyleFormats.Append(cFormat);
        cStyleFormats.Count = System.Convert.ToUInt32(cStyleFormats.ChildElements.Count);

        CellFormats cFormats = new CellFormats();

        for (int x = 0; x <= eStyles.Count - 1; x++)
        {
            cFormat = new CellFormat();
            cFormat.NumberFormatId = (uint)eStyles[x].numberFormatID;
            cFormat.FontId = (uint)eStyles[x].fontID;
            cFormat.FillId = (uint)eStyles[x].fillID;
            cFormat.BorderId = (uint)eStyles[x].borderID;

            Alignment cellAlignment = new Alignment();

            switch (eStyles[x].horizontalAlignment)
            {
                case "LEFT":
                    {
                        cellAlignment.Horizontal = HorizontalAlignmentValues.Left;
                        break;
                    }

                case "CENTER":
                    {
                        cellAlignment.Horizontal = HorizontalAlignmentValues.Center;
                        break;
                    }

                case "RIGHT":
                    {
                        cellAlignment.Horizontal = HorizontalAlignmentValues.Right;
                        break;
                    }

                default:
                    {
                        cellAlignment.Horizontal = HorizontalAlignmentValues.Left;
                        break;
                    }
            }

            switch (eStyles[x].verticalAlignment)
            {
                case "TOP":
                    {
                        cellAlignment.Vertical = VerticalAlignmentValues.Top;
                        break;
                    }

                case "CENTER":
                    {
                        cellAlignment.Vertical = VerticalAlignmentValues.Center;
                        break;
                    }

                case "BOTTOM":
                    {
                        cellAlignment.Vertical = VerticalAlignmentValues.Bottom;
                        break;
                    }

                default:
                    {
                        cellAlignment.Vertical = VerticalAlignmentValues.Top;
                        break;
                    }
            }

            cFormat.Alignment = cellAlignment;

            if (eStyles[x].numberFormatID != 0)
                cFormat.ApplyNumberFormat = true;

            cFormats.Append(cFormat);
        }

        cFormats.Count = System.Convert.ToUInt32(cFormats.ChildElements.Count);


        // ##########################################################################################################################
        // ###
        // ###    Applying all styles to new stylesheet
        // ###
        // ##########################################################################################################################

        newStyleSheet.Append(nFormats);
        newStyleSheet.Append(fnts);
        newStyleSheet.Append(cellFills);
        newStyleSheet.Append(cellBorders);
        newStyleSheet.Append(cStyleFormats);
        newStyleSheet.Append(cFormats);


        // ##########################################################################################################################
        // ###
        // ###    Setting required default styling
        // ###
        // ##########################################################################################################################

        CellStyles cStyles = new CellStyles();
        CellStyle cStyle = new CellStyle();
        cStyle.Name = "Normal";
        cStyle.BuiltinId = 0;
        cStyle.FormatId = 0;
        cStyle.BuiltinId = 0;
        cStyles.Append(cStyle);
        cStyles.Count = System.Convert.ToUInt32(cStyles.ChildElements.Count);
        newStyleSheet.Append(cStyles);

        DifferentialFormats dFormats = new DifferentialFormats();
        dFormats.Count = 0;
        newStyleSheet.Append(dFormats);

        TableStyles tStyles = new TableStyles();
        tStyles.Count = 0;
        tStyles.DefaultTableStyle = "TableStyleMedium9";
        tStyles.DefaultPivotStyle = "PivotStyleLight16";
        newStyleSheet.Append(tStyles);

        return newStyleSheet;
    }

    public static Column createColumnData(int columnIndex, double columnWidth, bool custom)
    {
        Column col = new Column();
        col.Min = (uint)columnIndex;
        col.Max = (uint)columnIndex;
        col.Width = columnWidth;
        col.CustomWidth = custom;
        return col;
    }

    public static void addImageToSheet(WorksheetPart wPart, Worksheet wSheet, eImage sImage)
    {
        DrawingsPart sheetDrawingPart = wPart.AddNewPart<DrawingsPart>();
        ImagePart sheetImagePart = sheetDrawingPart.AddImagePart(ImagePartType.Png, wPart.GetIdOfPart(sheetDrawingPart));

        using (FileStream fs = new FileStream(sImage.path, FileMode.Open))
        {
            sheetImagePart.FeedData(fs);
        }

        oNonVisDrawProp nonVisDrawingPart = new oNonVisDrawProp();
        nonVisDrawingPart.Id = (uint)sImage.ID;
        nonVisDrawingPart.Name = sImage.name;
        nonVisDrawingPart.Description = "Image";

        oPicLocks picLocks = new oPicLocks();
        picLocks.NoChangeAspect = true;
        picLocks.NoChangeArrowheads = true;

        oNonVisDrawPicProp nonVisPicPart = new oNonVisDrawPicProp();
        nonVisPicPart.PictureLocks = picLocks;
        oNonVisPicProp nonVisPicProp = new oNonVisPicProp();
        nonVisPicProp.NonVisualDrawingProperties = nonVisDrawingPart;
        nonVisPicProp.NonVisualPictureDrawingProperties = nonVisPicPart;

        oStretch iStretch = new oStretch();
        iStretch.FillRectangle = new DocumentFormat.OpenXml.Drawing.FillRectangle();

        oBlipFill iBlipfill = new oBlipFill();
        oBlip iBlip = new oBlip();
        iBlip.Embed = sheetDrawingPart.GetIdOfPart(sheetImagePart);
        iBlip.CompressionState = DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print;
        iBlipfill.Blip = iBlip;
        iBlipfill.SourceRectangle = new DocumentFormat.OpenXml.Drawing.SourceRectangle();
        iBlipfill.Append(iStretch);

        oTransoform2D iT2D = new oTransoform2D();
        oOffset iOffset = new oOffset();
        iOffset.X = 0;
        iOffset.Y = 0;
        iT2D.Offset = iOffset;

        int inchToEMU = 914400;
        oExtents imageExtents = new oExtents();
        imageExtents.Cx = (Int64Value)(sImage.width * inchToEMU);
        imageExtents.Cy = (Int64Value)(sImage.height * inchToEMU);

        iT2D.Extents = imageExtents;

        oShapeProperties iShapeProperties = new oShapeProperties();
        iShapeProperties.BlackWhiteMode = DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues.Auto;
        iShapeProperties.Transform2D = iT2D;
        oPresetGeometry imagePresetGeom = new oPresetGeometry();
        imagePresetGeom.Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle;
        imagePresetGeom.AdjustValueList = new DocumentFormat.OpenXml.Drawing.AdjustValueList();
        iShapeProperties.Append(imagePresetGeom);
        iShapeProperties.Append(new DocumentFormat.OpenXml.Drawing.NoFill());

        oPicture iPicture = new oPicture();
        iPicture.NonVisualPictureProperties = nonVisPicProp;
        iPicture.BlipFill = iBlipfill;
        iPicture.ShapeProperties = iShapeProperties;

        oPosition iPosition = new oPosition();
        iPosition.X = sImage.xCord * inchToEMU;
        iPosition.Y = sImage.yCord * inchToEMU;
        Extent iExtent = new Extent();
        iExtent.Cx = imageExtents.Cx;
        iExtent.Cy = imageExtents.Cy;
        AbsoluteAnchor anchor = new AbsoluteAnchor();
        anchor.Position = iPosition;
        anchor.Extent = iExtent;
        anchor.Append(iPicture);
        anchor.Append(new ClientData());
        WorksheetDrawing wsDrawing = new WorksheetDrawing();
        wsDrawing.Append(anchor);
        Drawing iDrawing = new Drawing();
        iDrawing.Id = sheetDrawingPart.GetIdOfPart(sheetImagePart);

        wsDrawing.Save(sheetDrawingPart);

        wSheet.Append(iDrawing);
    }
}
