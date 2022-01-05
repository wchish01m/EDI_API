public class eColumn
{
    private int columnNum = 1;
    private int columnWidth = 10;

    public int Index
    {
        get
        {
            return columnNum;
        }
        set
        {
            columnNum = value;
        }
    }

    public int Width
    {
        get
        {
            return columnWidth;
        }
        set
        {
            columnWidth = value;
        }
    }
}
