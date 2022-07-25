using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
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
        public async Task<IActionResult> Post()
        {
            IFormFile file = Request.Form.Files.FirstOrDefault();

            if (file != null)
            {
                try
                {
                    var filePath = Path.GetTempPath();
                    var fileName = filePath + file.FileName;

                    using (var stream = System.IO.File.Create(fileName))
                    {
                        await file.CopyToAsync(stream);
                    }

                    // запуск библиотеки
                    ExcelToJsonConverter converter = new ExcelToJsonConverter();
                    string sw = converter.JsonConvert(fileName);

                    return StatusCode(200, sw);
                }
                catch (Exception err)
                {
                    return StatusCode(500, err.Message);
                }
            }
            else
            {
                return NotFound("Входные файлы не найдены");
            }
        }
    }
}
