using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using static eCell_Borders;
using static eCell_Font;
using static eCell_FillColor;
using static eCell_TextFormat;
using static eCell_Style;
using static EDI_API.Excel_Creator.eCell;
using static Excel_Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Data.SqlClient;
using System.Web;
using EDI_API.Excel_Creator;
using Microsoft.AspNetCore.Hosting;

namespace EDI_API
{
    public class Sheet_Creation
    {
        private readonly IWebHostEnvironment _hostingEnvironment;
        public static eSheet createSheet()
        {
            eSheet sSheet = new eSheet();

            eCell_Font attentionFont = new eCell_Font();
            eCell_Borders normalBorder = new eCell_Borders();
            eCell_Style attentionStyle = new eCell_Style();

            normalBorder.setBorder = "THIN";
            normalBorder.name = "ALL THIN";
            sSheet.addBorders = normalBorder;

            attentionFont.name = "Times New Roman";
            attentionFont.size = 16;
            sSheet.addFont = attentionFont;

            attentionStyle.name = "ATTENTION STYLE";
            attentionStyle.borderID = getBorderByName("ALL THIN", sSheet.borders);
            attentionStyle.fontID = getFontByName("Times New Roman 16 Black", sSheet.fonts);
            sSheet.addStyle = attentionStyle;

            eCell_Style attentionToStyle = new eCell_Style();

            attentionToStyle.name = "ATTENTION TO STYLE";
            attentionToStyle.borderID = getBorderByName("ALL THIN", sSheet.borders);
            attentionToStyle.fontID = getFontByName("Times New Roman 16 Black", sSheet.fonts);
            attentionToStyle.horizontalAlignment = "RIGHT";
            sSheet.addStyle = attentionToStyle;

            eCell_Font poFont = new eCell_Font();
            eCell_Style poNumberStyle = new eCell_Style();

            poFont.name = "Times New Roman";
            poFont.size = 16;
            poFont.bold = true;
            sSheet.addFont = poFont;

            poNumberStyle.name = "PO NUMBER STYLE";
            poNumberStyle.borderID = getBorderByName("ALL THIN", sSheet.borders);
            poNumberStyle.fontID = getFontByName("Times New Roman 16 Black Bold", sSheet.fonts);
            sSheet.addStyle = poNumberStyle;

            eCell_Style poStyle = new eCell_Style();

            poStyle.name = "PO STYLE";
            poStyle.borderID = getBorderByName("ALL THIN", sSheet.borders);
            poStyle.fontID = getFontByName("Times New Roman 16 Black Bold", sSheet.fonts);
            poStyle.horizontalAlignment = "RIGHT";
            sSheet.addStyle = poStyle;

            eCell_Font headerFont = new eCell_Font();
            eCell_FillColor headerFill = new eCell_FillColor();
            eCell_Style headerStyle = new eCell_Style();

            headerFont.name = "Times New Roman";
            headerFont.clr = "White";
            headerFont.size = 16;
            headerFont.bold = true;
            sSheet.addFont = headerFont;

            headerFill.clr = "DarkBlue";
            headerFill.pattern = "SOLID";
            sSheet.addFill = headerFill;



            headerStyle.name = "HEADER STYLE";
            headerStyle.borderID = getBorderByName("ALL THIN", sSheet.borders);
            headerStyle.fontID = getFontByName("Times New Roman 16 White Bold", sSheet.fonts);
            headerStyle.fillID = getFillByName("DarkBlue", sSheet.fills);
            sSheet.addStyle = headerStyle;

            eCell_Style orderHeaderStyle = new eCell_Style();

            orderHeaderStyle.name = "ORDER HEADER STYLE";
            orderHeaderStyle.borderID = getBorderByName("ALL THICK", sSheet.borders);
            orderHeaderStyle.fontID = getFontByName("Times New Roman 16 White Bold", sSheet.fonts);
            orderHeaderStyle.fillID = getFillByName("DarkBlue", sSheet.fills);
            orderHeaderStyle.horizontalAlignment = "CENTER";
            sSheet.addStyle = orderHeaderStyle;

            eCell_Style orderDateHeaderStyle = new eCell_Style();

            orderDateHeaderStyle.name = "ORDER DATE HEADER STYLE";
            orderDateHeaderStyle.borderID = getBorderByName("ALL THICK", sSheet.borders);
            orderDateHeaderStyle.fontID = getFontByName("Times New Roman 16 White Bold", sSheet.fonts);
            orderDateHeaderStyle.fillID = getFillByName("DarkBlue", sSheet.fills);
            orderDateHeaderStyle.horizontalAlignment = "CENTER";
            orderDateHeaderStyle.numberFormatID = 14;
            sSheet.addStyle = orderDateHeaderStyle;

            eCell_Font normalFont = new eCell_Font();
            eCell_Style normalStyle = new eCell_Style();

            normalFont.name = "Times New Roman";
            sSheet.addFont = normalFont;

            normalBorder.setBorder = "THIN";
            normalBorder.name = "ALL THIN";
            sSheet.addBorders = normalBorder;

            normalStyle.name = "NORMAL STYLE";
            normalStyle.fontID = getFontByName("Times New Roman 12 Black", sSheet.fonts);
            normalStyle.borderID = getBorderByName("ALL THIN", sSheet.borders);
            normalStyle.fillID = getFillByName("NONE", sSheet.fills);
            sSheet.addStyle = normalStyle;

            eCell_TextFormat currencyFormat = new eCell_TextFormat();
            currencyFormat.name = "CURRENCY";
            currencyFormat.numFormat = "$ #,##0.00";
            currencyFormat.indexID = 1;

            sSheet.addTextFormat = currencyFormat;

            eCell_Style currencyStyle = new eCell_Style();
            currencyStyle.name = "CURRENCY STYLE";
            currencyStyle.fontID = getFontByName("Times New Roman 12 Black", sSheet.fonts);
            currencyStyle.borderID = getBorderByName("ALL THIN", sSheet.borders);
            currencyStyle.fillID = getFillByName("NONE", sSheet.fills);
            currencyStyle.numberFormatID = getTextFormatByName("CURRENCY", sSheet.textFormats);
            sSheet.addStyle = currencyStyle;

            eCell_Font boldFont = new eCell_Font();
            eCell_FillColor orderFill = new eCell_FillColor();
            eCell_Style orderStyle = new eCell_Style();

            boldFont.name = "Times New Roman";
            boldFont.bold = true;
            sSheet.addFont = boldFont;

            orderFill.clr = "LightBlue";
            orderFill.pattern = "SOLID";
            sSheet.addFill = orderFill;

            orderStyle.name = "ORDER STYLE";
            orderStyle.fontID = getFontByName("Times New Roman 12 Black Bold", sSheet.fonts);
            orderStyle.borderID = getBorderByName("ALL THIN", sSheet.borders);
            orderStyle.fillID = getFillByName("LightBlue", sSheet.fills);
            orderStyle.horizontalAlignment = "RIGHT";
            sSheet.addStyle = orderStyle;

            eCell_Style orderQtyStyle = new eCell_Style();

            orderQtyStyle.name = "ORDER QTY STYLE";
            orderQtyStyle.fontID = getFontByName("Times New Roman 12 Black Bold", sSheet.fonts);
            orderQtyStyle.borderID = getBorderByName("ALL THIN", sSheet.borders);
            orderQtyStyle.fillID = getFillByName("LightBlue", sSheet.fills);
            orderQtyStyle.horizontalAlignment = "CENTER";
            sSheet.addStyle = orderQtyStyle;

            eCell_FillColor adjustFill = new eCell_FillColor();
            eCell_Style adjustStyle = new eCell_Style();

            adjustFill.clr = "Khaki";
            adjustFill.pattern = "SOLID";
            sSheet.addFill = adjustFill;

            adjustStyle.name = "ADJUST STYLE";
            adjustStyle.fontID = getFontByName("Times New Roman 12 Black Bold", sSheet.fonts);
            adjustStyle.borderID = getBorderByName("ALL THIN", sSheet.borders);
            adjustStyle.fillID = getFillByName("Khaki", sSheet.fills);
            adjustStyle.horizontalAlignment = "CENTER";
            sSheet.addStyle = adjustStyle;

            eCell_Style adjustMergedStyle = new eCell_Style();

            adjustMergedStyle.name = "ADJUST MERGED STYLE";
            adjustMergedStyle.fontID = getFontByName("Times New Roman 12 Black Bold", sSheet.fonts);
            adjustMergedStyle.borderID = getBorderByName("ALL THIN", sSheet.borders);
            adjustMergedStyle.fillID = getFillByName("Khaki", sSheet.fills);
            adjustMergedStyle.horizontalAlignment = "RIGHT";
            sSheet.addStyle = adjustMergedStyle;


            eColumn sheetColumn = new eColumn();
            sheetColumn.Index = 1;
            sheetColumn.Width = 24;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 2;
            sheetColumn.Width = 24;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 3;
            sheetColumn.Width = 24;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 4;
            sheetColumn.Width = 20;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 5;
            sheetColumn.Width = 16;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 6;
            sheetColumn.Width = 18;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 7;
            sheetColumn.Width = 18;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 8;
            sheetColumn.Width = 23;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 9;
            sheetColumn.Width = 32;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 10;
            sheetColumn.Width = 13;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 11;
            sheetColumn.Width = 11;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 12;
            sheetColumn.Width = 16;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 13;
            sheetColumn.Width = 16;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 14;
            sheetColumn.Width = 16;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 15;
            sheetColumn.Width = 16;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 16;
            sheetColumn.Width = 16;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 17;
            sheetColumn.Width = 16;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 18;
            sheetColumn.Width = 16;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 19;
            sheetColumn.Width = 11;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 20;
            sheetColumn.Width = 18;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 21;
            sheetColumn.Width = 18;
            sSheet.addColumn = sheetColumn;

            sheetColumn = new eColumn();
            sheetColumn.Index = 22;
            sheetColumn.Width = 18;
            sSheet.addColumn = sheetColumn;
            return sSheet;
        }

        public static void BuildSpreadSheet()
        {
            // Variables for database connections and commands
            SqlConnection cn;
            string myQuery;
            SqlCommand cmd;

            // A variable to store the results of the sql query
            SqlDataReader returned = null/* TODO Change to default(_) if this is not a reference type */;
            bool createFile = false;

            List<eSheet> wbSheets = new();

            eSheet excelFile = createSheet();
            excelFile.name = "TEST";

            int row = 13;
            int startingCol = 1;

            for (int x = 0; x <= 10; x++)
            {
                eCell headerCell = new eCell();
                headerCell.styleID = (uint)findStyle("HEADER STYLE", excelFile.styles);
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
                    headerCell.value = "ASN Date 7 Time";

                headerCell.row = row;
                headerCell.column = (startingCol + x).ToString();

                excelFile.addCell = headerCell;
            }

            row += 1;

            for (int x = 0; x <= 10; x++)
            {
                eCell headerCell = new eCell();
                headerCell.styleID = (uint)findStyle("HEADER STYLE", excelFile.styles);
                if (x == 0)
                    headerCell.value = "";
                else if (x == 1)
                    headerCell.value = "";
                else if (x == 2)
                    headerCell.value = "";
                else if (x == 3)
                    headerCell.value = "";
                else if (x == 4)
                    headerCell.value = "";
                else if (x == 5)
                    headerCell.value = "";
                else if (x == 6)
                    headerCell.value = "";
                else if (x == 7)
                    headerCell.value = "";
                else if (x == 8)
                    headerCell.value = "";
                else if (x == 9)
                    headerCell.value = "";
                else if (x == 10)
                    headerCell.value = "";

                headerCell.row = row;
                headerCell.column = (startingCol + x).ToString();

                excelFile.addCell = headerCell;
            }

            row += 1;

            // Creating the sql query
            myQuery = "SELECT ctpcode, cradleypn, cpartsnum, nshipqty, cshippernum, creferencenum, ctradingpartner, ccustpo, cdestination, dshiptime, dasntime FROM Shipping_Results ORDER BY dshiptime ";

            cn = new SqlConnection(@"Data Source=192.168.4.4;Initial Catalog=EDI;User ID=TopreWeb;Password=@topre123");

            // Connecting to the database
            cn.Open();
            // Executing the sql query
            cmd = new SqlCommand(myQuery, cn);

            // Assigning the results of the query to the datareader variable
            returned = cmd.ExecuteReader();

            while (returned.Read())
            {
                createFile = true;
                for (int x = 0; x <= 10; x++)
                {
                    eCell sheetCell = new eCell();
                    sheetCell.styleID = (uint)findStyle("NORMAL STYLE", excelFile.styles);

                    if (x == 0)
                        // Part Number
                        sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                    else if (x == 1)
                        // Delivery Location
                        sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                    else if (x == 2)
                        // Supplier Reference
                        sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                    else if (x == 3)
                        // Machine
                        sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                    else if (x == 4)
                        // Customer
                        sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                    else if (x == 5)
                        // Program
                        sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                    else if (x == 6)
                        // Material
                        sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                    else if (x == 7)
                        // Coating
                        sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                    else if (x == 8)
                        // Size
                        sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                    else if (x == 9)
                        sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                    else if (x == 10)
                        sheetCell.value = returned.GetValue(x).ToString().Trim().ToUpper();
                    

                    sheetCell.row = row;
                    sheetCell.column = (startingCol + x).ToString();

                    excelFile.addCell = sheetCell;
                }
            }

            // Close the datareader
            returned.Close();
            // Close the connection to the database
            cn.Close();

            wbSheets.Add(excelFile);
            // createFile = True
            //string f = _hostingEnvironment.WebRootPath;
            //string content = _hostingEnvironment.ContentRootPath;
            //string folder = Server.MapPath("~/Exported_Files/PO/") + "POs_" + Strings.Format(poDate, "yyyy-MM-dd");
            //string filePath = folder + "/" + poFacility + "_" + Strings.Replace(Strings.Replace(Strings.Replace(vendor, " ", "_"), ",", ""), ".", "") + "_" + Strings.Format(poDate, "yyyy-MM-dd") + "_ & " + poNum + ".xlsx";
            //poCount += 1;

            //updatePOs_Number(poFacility, vendor, poDate, poNum);

            create_Document("C:/Users/wchisholm/Desktop/test", wbSheets);
        }
    }
}
