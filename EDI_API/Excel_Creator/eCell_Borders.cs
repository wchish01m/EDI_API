using System.Collections.Generic;
using Microsoft.VisualBasic;

public class eCell_Borders
{
    private string border = "DEFAULT";
    private string border_Style = "DEFAULT";
    private string border_Color = "DEFAULT";
    private string lBorder = "DEFAULT";
    private string rBorder = "DEFAULT";
    private string tBorder = "DEFAULT";
    private string bBorder = "DEFAULT";
    private string dBorder = "DEFAULT";
    private string lBorder_Style = "DEFAULT";
    private string rBorder_Style = "DEFAULT";
    private string tBorder_Style = "DEFAULT";
    private string bBorder_Style = "DEFAULT";
    private string dBorder_Style = "DEFAULT";
    private string lBorder_Color = "DEFAULT";
    private string rBorder_Color = "DEFAULT";
    private string tBorder_Color = "DEFAULT";
    private string bBorder_Color = "DEFAULT";
    private string dBorder_Color = "DEFAULT";
    private string borderName = "DEFAULT";
    private int index = 0;

    public string setBorder
    {
        get
        {
            return border;
        }
        set
        {
            border = Strings.UCase(value);
            lBorder = Strings.UCase(value);
            rBorder = Strings.UCase(value);
            tBorder = Strings.UCase(value);
            bBorder = Strings.UCase(value);
        }
    }

    public string leftBorder
    {
        get
        {
            return lBorder;
        }
        set
        {
            lBorder = Strings.UCase(value);
            border = "VARIED";
        }
    }

    public string rightBorder
    {
        get
        {
            return rBorder;
        }
        set
        {
            rBorder = Strings.UCase(value);
            border = "VARIED";
        }
    }

    public string topBorder
    {
        get
        {
            return tBorder;
        }
        set
        {
            tBorder = Strings.UCase(value);
            border = "VARIED";
        }
    }

    public string bottomBorder
    {
        get
        {
            return bBorder;
        }
        set
        {
            bBorder = Strings.UCase(value);
            border = "VARIED";
        }
    }

    public string setBorderStyle
    {
        get
        {
            return border_Style;
        }
        set
        {
            border_Style = Strings.UCase(value);
            lBorder_Style = Strings.UCase(value);
            rBorder_Style = Strings.UCase(value);
            tBorder_Style = Strings.UCase(value);
            bBorder_Style = Strings.UCase(value);
        }
    }

    public string leftBorderStyle
    {
        get
        {
            return lBorder_Style;
        }
        set
        {
            lBorder_Style = Strings.UCase(value);
            border_Style = "VARIED";
        }
    }

    public string rightBorderStyle
    {
        get
        {
            return rBorder_Style;
        }
        set
        {
            rBorder_Style = Strings.UCase(value);
            border_Style = "VARIED";
        }
    }

    public string topBorderStyle
    {
        get
        {
            return tBorder_Style;
        }
        set
        {
            tBorder_Style = Strings.UCase(value);
            border_Style = "VARIED";
        }
    }

    public string bottomBorderStyle
    {
        get
        {
            return bBorder_Style;
        }
        set
        {
            bBorder_Style = Strings.UCase(value);
            border_Style = "VARIED";
        }
    }

    public int indexID
    {
        get
        {
            return index;
        }
        set
        {
            index = Strings.UCase((char)value);
        }
    }

    public string name
    {
        get
        {
            return borderName;
        }
        set
        {
            borderName = Strings.UCase(value);
        }
    }

    public static int getBorderByName(string name, List<eCell_Borders> eBorders)
    {
        int borderIndex = 0;
        for (int x = 0; x <= eBorders.Count - 1; x++)
        {
            if (Strings.UCase(name) == Strings.UCase(eBorders[x].name))
            {
                borderIndex = x;
                break;
            }
        }
        return borderIndex;
    }
}
