using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using System.Drawing;

public class eCell_Font
{
    private string font = "Calibri";
    private double fontSize = 12;
    private string fontColor = "00000000"; // Black in hexadecimal colors
    private string colorName = "Black";
    private bool fontBold = false;
    private string boldName = "";
    private string fullName;
    private int index = 0;

    public string name
    {
        get
        {
            return font;
        }
        set
        {
            font = value;
        }
    }

    public double size
    {
        get
        {
            return fontSize;
        }
        set
        {
            fontSize = value;
        }
    }

    public string clr
    {
        get
        {
            return fontColor;
        }
        set
        {
            fontColor = HexValue_Color(Color.FromName(value));
            colorName = value;
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
            index = value;
        }
    }

    public bool bold
    {
        get
        {
            return fontBold;
        }
        set
        {
            fontBold = value;
            if (value == true)
                boldName = "Bold";
            else
                boldName = "";
        }
    }

    public string fontFullName
    {
        get
        {
            return Strings.Trim(name + " " + fontSize + " " + colorName + " " + boldName);
        }
    }

    private string HexValue_Color(Color colorName)
    {
        string r = colorName.R.ToString("X");
        string G = colorName.G.ToString("X");
        string B = colorName.B.ToString("X");
        if (r.Length == 1)
            r = 0 + r;
        if (G.Length == 1)
            G = 0 + G;
        if (B.Length == 1)
            B = 0 + B;
        return "00" + r + G + B;
    }

    public static int getFontByName(string name, List<eCell_Font> eFonts)
    {
        int fontIndex = 0;
        for (int x = 0; x <= eFonts.Count - 1; x++)
        {
            if (Strings.UCase(name) == Strings.UCase(eFonts[x].fontFullName))
            {
                fontIndex = x;
                break;
            }
        }
        return fontIndex;
    }
}
