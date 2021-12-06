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

public class eCell_FillColor
{
    private string fillName = "NONE";
    private string fillColor = "NONE";
    private string patternFill = "NONE";
    private int index = 0;

    public string clr
    {
        get
        {
            return fillColor;
        }
        set
        {
            // MsgBox(HexValue_Color(Color.FromName(value)))
            fillColor = HexValue_Color(Color.FromName(value));
            fillName = Strings.UCase(value);
        }
    }

    public void setRGB_clr(int R, int G, int B)
    {
        fillColor = RGB_HexValue_Color(R, G, B);
        fillName = "RGB(" + R + ", " + G + ", " + B + ")";
    }


    public string pattern
    {
        get
        {
            return patternFill;
        }
        set
        {
            patternFill = Strings.UCase(value);
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

    public string name
    {
        get
        {
            return fillName;
        }
    }

    public static int getFillByName(string name, List<eCell_FillColor> eFills)
    {
        int fillIndex = 0;
        for (int x = 0; x <= eFills.Count - 1; x++)
        {
            if (Strings.UCase(name) == Strings.UCase(eFills[x].name))
            {
                fillIndex = x;
                break;
            }
        }
        return fillIndex;
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

    private string RGB_HexValue_Color(int R, int G, int B)
    {
        string rHex = Conversion.Hex(R);
        string gHex = Conversion.Hex(G);
        string bHex = Conversion.Hex(B);
        if (rHex.Length == 1)
            rHex = 0 + R.ToString();
        if (gHex.Length == 1)
            gHex = 0 + G.ToString();
        if (bHex.Length == 1)
            bHex = 0 + B.ToString();
        return "00" + rHex + gHex + bHex;
    }
}
