using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EDI_API.Excel_Creator
{
    public class eCell
    {
            private string text = "";
            private string textFormat = "";
            private string formulaString = "";
            private int styleIndex = 0;
            private string cellFormat = "DEFAULT";
            private string columnChar = "A";
            private int columnNum = 1;
            private int rowNum = 1;
            private bool mergedCells = false;
            private string mergedAddr = "";
            private string mergedColumn = "";
            private int mergedColumnNum = 1;
            private int mergedRowNum = 1;
            private string verticalAlign = "TOP";
            private string horizontalAlign = "LEFT";

            public string value
            {
                get
                {
                    return text;
                }
                set
                {
                    text = value;
                }
            }

            public uint styleID
            {
                get
                {
                    return (uint)styleIndex;
                }
                set
                {
                    styleIndex = (int)value;
                }
            }

            public int row
            {
                get
                {
                    return rowNum;
                }
                set
                {
                    rowNum = value;
                }
            }

            public string column
            {
                get
                {
                    return columnChar;
                }
                set
                {
                    columnChar = ConvertToLetter(System.Convert.ToInt32(value));
                    columnNum = int.Parse(value);
                }
            }

            public int colNum
            {
                get
                {
                    return columnNum;
                }
            }

            public string address
            {
                get
                {
                    return columnChar + rowNum;
                }
            }

            public bool merged
            {
                get
                {
                    return mergedCells;
                }
                set
                {
                    mergedCells = value;
                }
            }

            public string mergedCol
            {
                get
                {
                    return mergedColumn;
                }
                set
                {
                    mergedColumn = ConvertToLetter(System.Convert.ToInt32(value));
                    mergedColumnNum = int.Parse(value);
                }
            }

            public int mergedColNum
            {
                get
                {
                    return mergedColumnNum;
                }
            }

            public int mergedRow
            {
                get
                {
                    return mergedRowNum;
                }
                set
                {
                    mergedRowNum = value;
                }
            }

            public string mergedAddress
            {
                get
                {
                    return mergedColumn + mergedRowNum;
                }
            }

            public string formula
            {
                get
                {
                    return formulaString;
                }
                set
                {
                    formulaString = value;
                }
            }

            public string cellType
            {
                get
                {
                    return cellFormat;
                }
                set
                {
                    cellFormat = value.ToUpper();
                }
            }

            public string ConvertToLetter(int columnNum)
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
        }
    }
