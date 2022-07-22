using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelToJSONLib;
using System.IO;

namespace ExcelToJSON.Controllers
{
    [Route("[controller]")]
    [ApiController]
    public class ConvertExcelToJsonController : ControllerBase
    {
        [HttpGet]
        public string Get()
        {
            return "result get ConvertExcelToJson";
        }

        [HttpPost]
        public IActionResult Post()
        {
            try
            {
                // запуск библиотеки
                ExcelToJsonConverter converter = new ExcelToJsonConverter();
                string sw = converter.JsonConver();

                return StatusCode(200, sw);
            }
            catch (Exception err)
            {
                return StatusCode(500, err.Message);
            }
        }
    }
}
