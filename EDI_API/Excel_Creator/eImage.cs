
public class eImage
{
    private string filePath;
    private string imageName;
    private int imageID;
    private int xAnchor = 0;
    private int yAnchor = 0;
    private decimal widthDimension;
    private decimal heightDimension;

    public string path
    {
        get
        {
            return filePath;
        }
        set
        {
            filePath = value;
        }
    }

    public string name
    {
        get
        {
            return imageName;
        }
        set
        {
            imageName = value;
        }
    }

    public int xCord
    {
        get
        {
            return xAnchor;
        }
        set
        {
            xAnchor = value;
        }
    }

    public int yCord
    {
        get
        {
            return yAnchor;
        }
        set
        {
            yAnchor = value;
        }
    }

    public decimal width
    {
        get
        {
            return widthDimension;
        }
        set
        {
            widthDimension = value;
        }
    }

    public decimal height
    {
        get
        {
            return heightDimension;
        }
        set
        {
            heightDimension = value;
        }
    }

    public int ID
    {
        get
        {
            return imageID;
        }
        set
        {
            imageID = value;
        }
    }
}
