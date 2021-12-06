using EDI_API.Excel_Creator;
using EDI_API.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
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

        [HttpGet("BuildSpreadsheet")]
        public void getSheet()
        {
            Sheet_Creation.BuildSpreadSheet();
        }
    }
}
