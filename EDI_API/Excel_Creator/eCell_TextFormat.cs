using Microsoft.VisualBasic;
using System.Collections.Generic;

//'#########################################################################################################
//'####   Default number formats for cells
//'####
//'#########################################################################################################
//'ID  FORMAT CODE
//'0   General
//'1   0
//'2   0.00
//'3   #,##0
//'4   #,##0.00
//'9   0%
//'10  0.00%
//'11  0.00E+00
//'12  # ?/?
//'13  # ??/??
//'14  d/m/yyyy
//'15  d-mmm-yy
//'16  d-mmm
//'17  mmm-yy
//'18  h:mm tt
//'19  h:mm:ss tt
//'20  H:mm
//'21  H:mm : ss
//'22  m/d/yyyy H:mm
//'37  #,##0 ;(#,##0)
//'38  #,##0 ;[Red](#,##0)
//'39  #,##0.00;(#,##0.00)
//'40  #,##0.00;[Red](#,##0.00)
//'45  mm:ss
//'46  [h]:mm : ss
//'47  mmss.0
//'48  ##0.0E+0
//'49  @
//'#########################################################################################################

public class eCell_TextFormat
{
    private string numberFormat;
    private string formatName;
    private int index = 0;

    public string numFormat
    {
        get
        {
            return numberFormat;
        }
        set
        {
            numberFormat = value;
        }
    }

    public string name
    {
        get
        {
            return formatName;
        }
        set
        {
            formatName = value;
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
            index = (value + 164);
        }
    }

    public static int getTextFormatByName(string name, List<eCell_TextFormat> eTextFormat)
    {
        int textFormatIndex = 0;
        for (int x = 0; x <= eTextFormat.Count - 1; x++)
        {
            if (Strings.UCase(name) == Strings.UCase(eTextFormat[x].name))
            {
                textFormatIndex = eTextFormat[x].indexID;
                break;
            }
        }
        return textFormatIndex;
    }
}
