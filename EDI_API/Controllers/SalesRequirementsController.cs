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
    public class SalesRequirementsController : ControllerBase
    {
        [HttpGet("GetSalesInfo/{facility}")]
        public List<SalesRequirements> GetResultsInRange(string facility)
        {
            return Queries.GetSalesRequirements(facility);
        }
    }
}
