using Microsoft.VisualBasic;
using System.Collections.Generic;

public class eCell_Style
{
    private int numberFormatIndex = 0;
    private int fontIndex = 0;
    private int fillIndex = 0;
    private int borderIndex = 0;
    private int styleIndex = 0;
    private string hAlignment = "LEFT";
    private string vAlignment = "TOP";
    private string styleName = "DEFAULT";

    public int numberFormatID
    {
        get
        {
            return numberFormatIndex;
        }
        set
        {
            numberFormatIndex = value;
        }
    }

    public int fontID
    {
        get
        {
            return fontIndex;
        }
        set
        {
            fontIndex = value;
        }
    }

    public int fillID
    {
        get
        {
            return fillIndex;
        }
        set
        {
            fillIndex = value;
        }
    }

    public int borderID
    {
        get
        {
            return borderIndex;
        }
        set
        {
            borderIndex = value;
        }
    }

    public int styleID
    {
        get
        {
            return styleIndex;
        }
        set
        {
            styleIndex = value;
        }
    }

    public string name
    {
        get
        {
            return styleName;
        }
        set
        {
            styleName = value;
        }
    }

    public string horizontalAlignment
    {
        get
        {
            return hAlignment;
        }
        set
        {
            hAlignment = value;
        }
    }

    public string verticalAlignment
    {
        get
        {
            return vAlignment;
        }
        set
        {
            vAlignment = value;
        }
    }

    public static int findStyle(string name, List<eCell_Style> styles)
    {
        int index = 0;
        for (int x = 0; x <= styles.Count - 1; x++)
        {
            if (Strings.UCase(name) == Strings.UCase(styles[x].styleName))
            {
                index = x;
                break;
            }
        }

        return index;
    }
}
