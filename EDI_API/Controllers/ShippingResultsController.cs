using EDI_API.Models;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace EDI_API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ShippingResultsController : ControllerBase
    {
        [HttpGet("GetByRange/{facility}/{startDate}/{endDate}")]
        public List<GroupedParts> GetResultsInRange(string facility, string startDate, string endDate)
        {
            return Queries.GetResults(facility, startDate, endDate);
        }

        [HttpGet("BuildFile/{facility}/{startDate}/{endDate}")]
        public string BuildFile(string facility, string startDate, string endDate)
        {
            var timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            Sheet_Creation.BuildSpreadSheet(facility, startDate, endDate, timestamp);

            string filePath = @"~\File_Export\".Replace(@"~\", "") + "shipping_results_" + timestamp + ".xlsx";
            return filePath;
        }

        [HttpGet("Download/{filePath}")]
        public async Task<ActionResult> Download(string filePath)
        {
            byte[] bytes = await System.IO.File.ReadAllBytesAsync(filePath);
            System.IO.File.Delete(filePath);
            return File(bytes, "application/octet-stream", Path.GetFileName(filePath));
        }
    }
}
