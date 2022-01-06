using EDI_API.Excel_Creator;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using static eCell_Borders;
using static eCell_FillColor;
using static eCell_Font;
using static eCell_Style;
using static eCell_TextFormat;
using static Excel_Builder;

namespace EDI_API
{
    public class Sheet_Creation
    {
        public static eSheet createSheet()
        {
            eSheet sSheet = new();
            eCell_Borders normalBorder = new();

            normalBorder.setBorder = "THIN";
            normalBorder.name = "ALL THIN";
            sSheet.addBorders = normalBorder;

            eCell_Font headerFont = new();
            eCell_FillColor headerFill = new();
            eCell_Style headerStyle = new();

            headerFont.name = "Calibri";
            headerFont.clr = "White";
            headerFont.size = 16;
            headerFont.bold = true;
            sSheet.addFont = headerFont;

            headerFill.clr = "DarkBlue";
            headerFill.pattern = "SOLID";
            sSheet.addFill = headerFill;

            headerStyle.name = "HEADER STYLE";
            headerStyle.borderID = getBorderByName("ALL THIN", sSheet.borders);
            headerStyle.fontID = getFontByName("Calibri 16 White Bold", sSheet.fonts);
            headerStyle.fillID = getFillByName("DarkBlue", sSheet.fills);
            sSheet.addStyle = headerStyle;

            eCell_Font normalFont = new() 
            {
                name = "Calibri"
            };

            eCell_Style normalStyle = new()
            {
                name = "NORMAL STYLE",
                fontID = getFontByName("Calibri 12 Black", sSheet.fonts),
                borderID = getBorderByName("ALL THIN", sSheet.borders),
                fillID = getFillByName("NONE", sSheet.fills)
            };

            sSheet.addFont = normalFont;

            normalBorder.setBorder = "THIN";
            normalBorder.name = "ALL THIN";
            sSheet.addBorders = normalBorder;
            
            sSheet.addStyle = normalStyle;

            // Create format and style for date type
            eCell_TextFormat dateFormat = new()
            {
                name = "DATE",
                numFormat = "m/d/yyyy H:mm",
                indexID = 1
            };

            sSheet.addTextFormat = dateFormat;

            eCell_Style dateStyle = new()
            {
                name = "DATE STYLE",
                fontID = getFontByName("Calibri 12 Black", sSheet.fonts),
                borderID = getBorderByName("ALL THIN", sSheet.borders),
                fillID = getFillByName("NONE", sSheet.fills),
                numberFormatID = getTextFormatByName("DATE", sSheet.textFormats)
            };

            sSheet.addStyle = dateStyle;

            eColumn sheetColumn = new()
            {
                Index = 1,
                Width = 10
            };
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn
            {
                Index = 2,
                Width = 24
            };
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn
            {
                Index = 3,
                Width = 24
            };
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn
            {
                Index = 4,
                Width = 10
            };
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn
            {
                Index = 5,
                Width = 19
            };
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn
            {
                Index = 6,
                Width = 14
            };
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn
            {
                Index = 7,
                Width = 18
            };
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn
            {
                Index = 8,
                Width = 15
            };
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn
            {
                Index = 9,
                Width = 13
            };
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn
            {
                Index = 10,
                Width = 20
            };
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn
            {
                Index = 11,
                Width = 20
            };
            sSheet.addColumn = sheetColumn;
            
            return sSheet;
        }

        public static void BuildSpreadSheet(string facility, string startDate, string endDate, string timestamp)
        {
            // Variables for database connections and commands
            SqlConnection cn;
            string myQuery;
            SqlCommand cmd;

            List<eSheet> wbSheets = new();

            eSheet excelFile = createSheet();

            excelFile.name = @"Shipping Results";
            string filePath = @"~\File_Export\".Replace(@"~\", "") + "shipping_results_" + timestamp + ".xlsx";

            int row = 1;
            int startingCol = 1;

            for (int x = 0; x <= 10; x++)
            {
                eCell headerCell = new()
                {
                    styleID = (uint)findStyle("HEADER STYLE", excelFile.styles)
                };

                if (x == 0)
                    headerCell.value = "TP Code";
                else if (x == 1)
                    headerCell.value = "Radley Part Number";
                else if (x == 2)
                    headerCell.value = "Part Number";
                else if (x == 3)
                    headerCell.value = "Quantity";
                else if (x == 4)
                    headerCell.value = "Shipper Number";
                else if (x == 5)
                    headerCell.value = "Ref Number";
                else if (x == 6)
                    headerCell.value = "Trading Partner";
                else if (x == 7)
                    headerCell.value = "Customer PO";
                else if (x == 8)
                    headerCell.value = "Destination";
                else if (x == 9)
                    headerCell.value = "Ship Date & Time";
                else if (x == 10)
                    headerCell.value = "ASN Date & Time";

                headerCell.row = row;
                headerCell.column = (startingCol + x).ToString();

                excelFile.addCell = headerCell;
            }

            row += 1;

            // Creating the sql query
            myQuery = @"SELECT ctpcode, cradleypn, cpartsnum, nshipqty, cshippernum, creferencenum, ctradingpartner, ccustpo, cdestination, dshiptime, dasntime 
                        FROM Shipping_Results
                        WHERE cfacility = '" + facility + "' AND dshipdate >= '" + startDate + "' AND dshipdate <= '" + endDate + "' " +
                       "ORDER BY dshiptime";

            cn = new SqlConnection(DatabaseConnection.GetALCS());

            // Connecting to the database
            cn.Open();

            // Executing the sql query
            cmd = new SqlCommand(myQuery, cn);

            // A variable to store the results of the sql query
            // Assigning the results of the query to the datareader variable
            SqlDataReader returned = cmd.ExecuteReader();

            while (returned.Read())
            {
                for (int x = 0; x <= 10; x++)
                {
                    eCell sheetCell = new()
                    {
                        styleID = (uint)findStyle("NORMAL STYLE", excelFile.styles)
                    };

                    switch (x)
                    {
                        case 0:
                            // TP Code
                            sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                            break;
                        case 1:
                            // Radley Part Number
                            sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                            break;
                        case 2:
                            // Part Number
                            sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                            break;
                        case 3:
                            // Ship Quantity
                            sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                            break;
                        case 4:
                            // Shipper Number
                            sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                            break;
                        case 5:
                            // Reference Number
                            sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                            break;
                        case 6:
                            // Trading Partner
                            sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                            break;
                        case 7:
                            // Customer PO
                            sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                            break;
                        case 8:
                            // Destination
                            sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                            break;
                        case 9:
                            // Ship Time
                            sheetCell.value = returned.GetValue(x).ToString().Trim();
                            sheetCell.styleID = (uint)findStyle("DATE STYLE", excelFile.styles);
                            break;
                        case 10:
                            // ASN Time
                            sheetCell.value = returned.GetValue(x).ToString().Trim();
                            sheetCell.styleID = (uint)findStyle("DATE STYLE", excelFile.styles);
                            break;
                    }

                    sheetCell.row = row;
                    sheetCell.column = (startingCol + x).ToString();

                    excelFile.addCell = sheetCell;
                }
                row += 1;
            }

            // Close the datareader
            returned.Close();

            // Close the connection to the database
            cn.Close();

            wbSheets.Add(excelFile);

            create_Document(filePath, wbSheets);
        }
    }
}
