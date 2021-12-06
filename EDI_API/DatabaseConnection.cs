namespace EDI_API
{
    class DatabaseConnection
    {
        /*This function will return TAC-AL's connection string.*/
        public static string GetALCS()
        {
            return "Data Source=192.168.4.185\\srv05; Initial Catalog=EDI; User ID=sqlweb; Password=$ystem$ql2019";
        }

        /*This function will return IT's connection string.*/
        public static string GetITCS()
        {
            return "Data Source=192.168.4.4; Initial Catalog=IT; User ID=TopreWeb; Password=@topre123";
        }
    }
}
