using EDI_API.Models;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;

namespace EDI_API
{
    public class Queries
    {
        /**
         * This function executes a query that is used to display
         * information to the Shipping_Results page in the EDI_Interface.
         */
        public static List<GroupedParts> GetResults(string facility, string startDate, string endDate)
        {
            List<GroupedParts> partsList = new();

            String myQuery;
            SqlConnection connection;
            SqlCommand command;

            SqlDataReader returnedReader;

            connection = new SqlConnection(DatabaseConnection.GetALCS());

            myQuery = @"SELECT ctpcode, cpartsnum, SUM(nshipqty) [ntotalshipped], dshipdate, 
                        SUBSTRING(
                                        ( 
                                            SELECT '|' + ctpcode + '^' + cradleypn + '^' + cpartsnum + '^' + CAST(nshipqty AS VARCHAR(50)) + '^' + cshippernum + '^' + creferencenum + 
                                                      '^' + ctradingpartner + '^' + ccustpo + '^' + cdestination + '^' + CAST(dshiptime AS VARCHAR(50)) + '^' + CAST(dasntime AS VARCHAR(50)) AS [text()] 
                                            FROM Shipping_Results b
                                            WHERE a.cfacility = b.cfacility AND a.ctpcode = b.ctpcode AND a.cpartsnum = b.cpartsnum AND a.dshipdate = b.dshipdate
                                            ORDER BY b.dshiptime 
                                            FOR XML PATH('')
                                        ), 2, 10000) [cshipdetails]
                        FROM Shipping_Results a
                        WHERE cfacility = '" + facility + "' AND dshipdate >= '" + startDate + "' AND dshipdate <= '" + endDate + "' " +
                        "GROUP BY cfacility, ctpcode, cpartsnum, dshipdate " +
                        "ORDER BY dshipdate, ctpcode, cpartsnum";

            connection.Open();

            //Console.WriteLine(myQuery);

            command = new SqlCommand(myQuery, connection);

            returnedReader = command.ExecuteReader();

            while (returnedReader.Read())
            {
                GroupedParts parts = new();
                parts.tpCode = returnedReader[0].ToString();
                parts.partsNum = returnedReader[1].ToString();
                parts.sumQty = returnedReader[2].ToString();
                parts.shipDate = returnedReader[3].ToString();
                string[] partsInfo = returnedReader[4].ToString().Split('|');

                foreach (var partInfo in partsInfo)
                {
                    ShippingResults sd = new();
                    string[] details = partInfo.Split('^');
                    sd.tpCode = details[0];
                    sd.radleyPartNum = details[1];
                    sd.partNum = details[2];
                    sd.shipQty = details[3];
                    sd.shipperNum = details[4];
                    sd.referenceNum = details[5];
                    sd.tradingPartner = details[6];
                    sd.customerPO = details[7];
                    sd.destination = details[8];
                    sd.shipTime = details[9];
                    sd.asnTime = details[10];
                    parts.shipDetails.Add(sd);
                }
                partsList.Add(parts);
            }
            returnedReader.Close();
            return partsList;
        }
        /**
         * This function executes a query that is used to display
         * information to the Shipping_Result_Details page in the EDI_Interface.
         */
        public static List<ShippingResultDetails> GetDetailsBySearch(string partNum, string tpCode, string shipperNum, string referenceNum, string custSerial,
                                               string topSerial, string facility, string startDate, string endDate)
        {
            List<ShippingResultDetails> detailList = new();

            string myQuery;
            SqlConnection connection;
            SqlCommand command;

            SqlDataReader returnedReader;

            connection = new SqlConnection(DatabaseConnection.GetALCS());

            //myQuery = @"SELECT a.cpartsnum, cserialnumber, cpartloc, cshift, cpkgqty, cprintedby, dproddate, ddatecreated, dtimestamp,
            //            cfacility, ctpcode, cshippernum, creferencenum, b.cpartsnum, cshipserial, cserial
            //            FROM
            //            ( 
	           //             SELECT cpartsnum, cserialnumber, cpartloc, cshift, cpkgqty, cprintedby, cmfgloc, dproddate, ddatecreated, dtimestamp
	           //             FROM lsLabels
            //            )a
            //            LEFT JOIN
            //            (
            //                SELECT a.cfacility, a.ctpcode, a.cpartsnum, a.cshippernum, a.creferencenum, dshipdate, dshiptime, cshipserial, cserial
            //                FROM EDI.dbo.Shipping_Results a, EDI.dbo.Shipping_Result_Details b
            //                WHERE a.cfacility = b.cfacility AND a.cshippernum = b.cshippernum AND a.creferencenum = b.creferencenum AND a.ctpcode = b.ctpcode AND a.cpartsnum = b.cpartsnum
            //            )b ON cserialnumber = cserial
            //            WHERE cmfgloc = '" + facility + "' ";

            myQuery = @"SELECT a.cpartsnum, cserialnumber, cpartloc, cshift, cpkgqty, cprintedby, dproddate, ddatecreated, dtimestamp,
                       cfacility, ctpcode, cshippernum, creferencenum, b.cpartsnum, cshipserial, cserial
                       FROM OPENQUERY(srv04, 'SELECT cpartsnum, cserialnumber, cpartloc, cshift, cpkgqty, cprintedby, cmfgloc, dproddate, ddatecreated, dtimestamp
					   FROM Topre_Labeling.dbo.lsLabels
                       WHERE cmfgloc = ''" + facility + "'' ";

            // Handle partNum
            if (partNum != null)
            {
                myQuery += "AND cpartsnum = ''" + partNum + "'' ";
            }

            // Handle custSerial
            if (custSerial != null)
            {
                myQuery += "AND cserialnumber IN(SELECT cserialnumber FROM srv04.Topre_Labeling.dbo.lsLabels WHERE cbatchid =(SELECT cbatchid FROM srv04.Topre_Labeling.dbo.lsLabels WHERE cserialnumber = (SELECT cserial FROM [EDI].[dbo].Shipping_Result_Details WHERE cshipserial = ''" + custSerial + "''))) ";
            }

            // Handle topSerial
            if (topSerial != null)
            {
                myQuery += "AND cserialnumber IN(SELECT cserialnumber FROM srv04.Topre_Labeling.dbo.lsLabels WHERE cbatchid =(SELECT cbatchid FROM srv04.Topre_Labeling.dbo.lsLabels WHERE cserialnumber =  ''" + topSerial + "'')) ";
            }

            // Handle startDate and endDate
            if (startDate != null && endDate != null)
            {
                myQuery += "AND dproddate >= ''" + startDate + "'' AND dproddate <= ''" + endDate + "'' ";
            }

            myQuery += @" ')a
                       LEFT JOIN
                       (
                           SELECT a.cfacility, a.ctpcode, a.cpartsnum, a.cshippernum, a.creferencenum, dshipdate, dshiptime, cshipserial, cserial
                           FROM Shipping_Results a, Shipping_Result_Details b
                           WHERE a.cfacility = b.cfacility AND a.cshippernum = b.cshippernum AND a.creferencenum = b.creferencenum AND a.ctpcode = b.ctpcode AND a.cpartsnum = b.cpartsnum
                       )b ON cserialnumber = cserial
                       WHERE cmfgloc = '" + facility + "' ";

            // Handle shipperNum
            if (shipperNum != null)
            {
                myQuery += "AND cshippernum = '" + shipperNum + "' ";
            }

            // Handle referenceNum
            if (referenceNum != null)
            {
                myQuery += "AND creferencenum = '" + referenceNum + "' ";
            }

            // Handle tradingpartner
            if (tpCode != null)
            {
                myQuery += "AND ctpcode = '" + tpCode + "' ";
            }

            myQuery += "ORDER BY cserialnumber";

            Console.WriteLine(myQuery);

            connection.Open();

            command = new SqlCommand(myQuery, connection);

            returnedReader = command.ExecuteReader();

            while (returnedReader.Read())
            {
                ShippingResultDetails details = new();
                details.partsNum = returnedReader[0].ToString();
                details.serialNumber = returnedReader[1].ToString();
                details.partLoc = returnedReader[2].ToString();
                details.shift = returnedReader[3].ToString();
                details.pkgQty = returnedReader[4].ToString();
                details.printedBy = returnedReader[5].ToString();
                details.prodDate = returnedReader[6].ToString();
                details.dateCreated = returnedReader[7].ToString();
                details.timestamp = returnedReader[8].ToString();
                details.facility = returnedReader[9].ToString();
                details.tpCode = returnedReader[10].ToString();
                details.shipperNum = returnedReader[11].ToString();
                details.referenceNum = returnedReader[12].ToString();
                details.partsNum2 = returnedReader[13].ToString();
                details.shipSerial = returnedReader[14].ToString();
                details.serial = returnedReader[15].ToString();
                detailList.Add(details);
            }
            returnedReader.Close();
            connection.Close();
            return detailList;
        }

        /**
         * This function executes a query that is used to display
         * information to the Sales_Requirements page in the EDI_Interface.
         */
        public static List<SalesRequirements> GetSalesRequirements(string facility)
        {
            List<SalesRequirements> salesRequirementsList = new();

            string myQuery;
            SqlConnection connection;
            SqlCommand command;

            SqlDataReader returnedReader;

            connection = new SqlConnection(DatabaseConnection.GetALCS());

            myQuery = @"SELECT a.cfacility, a.ctradingpartner, a.cpartsnum, a.cdestination, a.cdockcode, a.creferencenum, a.crecordtype, 
a.dreqdate, a.dreqtime, a.nreqqty, a.nnetqty , b.creleasenum, b.dreleasedate, ntotalshipped, dlatestdateshipped, 
CASE 
	WHEN a.nnetqty = ntotalshipped AND ntotalshipped = a.nreqqty THEN 'SHIPPING DOCUMENTS NOT PRINTED' 
	WHEN a.nnetqty = 0 AND ntotalshipped = 0 THEN 'ASN NOT TRANSMITTED'
	WHEN a.nreqqty = a.nnetqty AND ntotalshipped = 0 THEN 'NOT SHIPPED'
	ELSE 'ASN SENT' END [cshipperstatus]
FROM Sales_Requirements a, Sales_Requirements_Origination b
LEFT JOIN
( 
	SELECT cfacility, ctradingpartner, cdestination, cpartsnum, creferencenum, nshipqty, SUM(nshipqty) AS ntotalshipped, MAX(dshipdate) AS dlatestdateshipped
	FROM Shipping_Results 
	GROUP BY cfacility, ctradingpartner, cdestination, cpartsnum, creferencenum, nshipqty
)c
ON b.cfacility = c.cfacility AND b.ctradingpartner = c.ctradingpartner AND b.cdestination = c.cdestination AND b.cpartsnum = c.cpartsnum AND b.creferencenum = c.creferencenum
WHERE a.cfacility = b.cfacility AND a.ctradingpartner = b.ctradingpartner AND a.cdestination = b.cdestination AND a.cpartsnum = b.cpartsnum AND a.creferencenum = b.creferencenum
AND a.ctpcode <> 't1' AND a.ctpcode <> 't2' AND a.ctpcode <> 't3' AND a.ctpcode <> 't4' AND a.nnetqty + ntotalshipped != a.nreqqty AND a.cfacility = '" + facility + "' ";

            connection.Open();

            //Console.WriteLine(myQuery);

            command = new SqlCommand(myQuery, connection);

            returnedReader = command.ExecuteReader();

            while (returnedReader.Read())
            {
                SalesRequirements salesRequirements = new();
                salesRequirements.facility = returnedReader[0].ToString();
                salesRequirements.tradingpartner = returnedReader[1].ToString();
                salesRequirements.partsnum = returnedReader[2].ToString();
                salesRequirements.destination = returnedReader[3].ToString();
                salesRequirements.dockcode = returnedReader[4].ToString();
                salesRequirements.referencenum = returnedReader[5].ToString();
                salesRequirements.recordtype = returnedReader[6].ToString();
                salesRequirements.reqdate = returnedReader[7].ToString();
                salesRequirements.reqtime = returnedReader[8].ToString();
                salesRequirements.reqqty = returnedReader[9].ToString();
                salesRequirements.netqty = returnedReader[10].ToString();
                salesRequirements.releasenum = returnedReader[11].ToString();
                salesRequirements.releasedate = returnedReader[12].ToString();
                salesRequirements.totalshipped = returnedReader[13].ToString();
                salesRequirements.latestdateshipped = returnedReader[14].ToString();
                salesRequirements.asnstatus = returnedReader[15].ToString();
                salesRequirementsList.Add(salesRequirements);
            }
            returnedReader.Close();
            connection.Close();
            return salesRequirementsList;
        }
    }
}
