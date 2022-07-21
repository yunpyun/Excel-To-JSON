using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelToJSONLib;

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

                return StatusCode(200, "Конвертация завершена");
            }
            catch (Exception err)
            {
                return StatusCode(500, err.Message);
            }
        }
    }
}
