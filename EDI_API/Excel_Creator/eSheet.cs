using DocumentFormat.OpenXml.Spreadsheet;
using EDI_API.Excel_Creator;
using System;
using System.Collections.Generic;
using System.Linq;

public class eSheet
{
    private List<eCell_Font> sheet_fonts = new List<eCell_Font>();
    private List<eCell_FillColor> sheet_fillColors = new List<eCell_FillColor>();
    private List<eCell_Borders> sheet_borders = new List<eCell_Borders>();
    private List<eCell_TextFormat> sheet_textFormats = new List<eCell_TextFormat>();
    private List<eCell_Style> sheet_styles = new List<eCell_Style>();
    private List<eCell> sheet_cells = new List<eCell>();
    private List<eColumn> sheet_columns = new List<eColumn>();
    private int maxColumn;
    private List<eImage> sheet_Images = new List<eImage>();
    private string sheetName = "DEFAULT";

    public eSheet()
    {
        eCell_Font cellFont = new eCell_Font();
        sheet_fonts.Add(cellFont);

        eCell_FillColor cellFill = new eCell_FillColor();
        sheet_fillColors.Add(cellFill);

        eCell_FillColor defaultFill = new eCell_FillColor();
        defaultFill.pattern = "Gray125";
        sheet_fillColors.Add(defaultFill);

        eCell_Borders cellBorder = new eCell_Borders();
        sheet_borders.Add(cellBorder);

        eCell_TextFormat cellTextFormat = new eCell_TextFormat();
        sheet_textFormats.Add(cellTextFormat);

        eCell_Style cellStyle = new eCell_Style();
        sheet_styles.Add(cellStyle);
    }

    public List<eCell_Font> fonts
    {
        get
        {
            return sheet_fonts;
        }
    }

    public eCell_Font addFont
    {
        set
        {
            value.indexID = sheet_fonts.Count;
            sheet_fonts.Add(value);
        }
    }

    public List<eCell_FillColor> fills
    {
        get
        {
            return sheet_fillColors;
        }
    }

    public eCell_FillColor addFill
    {
        set
        {
            value.indexID = sheet_fillColors.Count;
            sheet_fillColors.Add(value);
        }
    }

    public List<eCell_Borders> borders
    {
        get
        {
            return sheet_borders;
        }
    }

    public eCell_Borders addBorders
    {
        set
        {
            value.indexID = sheet_borders.Count;
            sheet_borders.Add(value);
        }
    }

    public List<eCell_TextFormat> textFormats
    {
        get
        {
            return sheet_textFormats;
        }
    }

    public eCell_TextFormat addTextFormat
    {
        set
        {
            value.indexID = sheet_textFormats.Count;
            sheet_textFormats.Add(value);
        }
    }

    public List<eCell_Style> styles
    {
        get
        {
            return sheet_styles;
        }
    }

    public eCell_Style addStyle
    {
        set
        {
            value.styleID = sheet_styles.Count;
            sheet_styles.Add(value);
        }
    }

    public List<eCell> cells
    {
        get
        {
            return sheet_cells;
        }
    }

    public eCell addCell
    {
        set
        {
            sheet_cells.Add(value);
        }
    }

    public List<eColumn> columns
    {
        get
        {
            return sheet_columns;
        }
    }
    public eColumn addColumn
    {
        set
        {
            sheet_columns.Add(value);
            if (maxColumn < value.Index)
                maxColumn = value.Index;
        }
    }

    public int columnLimit
    {
        get
        {
            return maxColumn;
        }
    }

    public static string findFont(eCell_Style cellStyle, List<eCell_Font> sFonts)
    {
        string font = "";
        for (int x = 0; x <= sFonts.Count - 1; x++)
        {
            if (cellStyle.fontID == sFonts[x].indexID)
                font = sFonts[x].name;
        }
        return font;
    }

    public static int findFontSize(eCell_Style cellStyle, List<eCell_Font> sFonts)
    {
        int size = 0;
        for (int x = 0; x <= sFonts.Count - 1; x++)
        {
            if (cellStyle.fontID == sFonts[x].indexID)
                size = (int)sFonts[x].size;
        }
        return size;
    }

    private static eCell_Style findStyle(int styleIndex, List<eCell_Style> sStyles)
    {
        eCell_Style sStyle = new eCell_Style();
        for (var x = 0; x <= sStyles.Count - 1; x++)
        {
            if (styleIndex == sStyles[x].styleID)
            {
                sStyle = sStyles[x];
                break;
            }
        }
        return sStyle;
    }

    public List<eImage> images
    {
        get
        {
            return sheet_Images;
        }
    }

    public eImage addImage
    {
        set
        {
            value.name = "Image " + (sheet_Images.Count + 1);
            value.ID = sheet_Images.Count + 1;
            sheet_Images.Add(value);
        }
    }

    public string name
    {
        get
        {
            return sheetName;
        }
        set
        {
            sheetName = value;
        }
    }
}
