using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelToJSONLib;
using System.IO;
using System.Net.Http;

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
        public async Task<IActionResult> Post(string url)
        {
            IFormFile file = Request.Form.Files.FirstOrDefault();

            using var client = new HttpClient();

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

                    var data = new System.Net.Http.StringContent(sw);

                    var response = await client.PostAsync(url, data);

                    string result = response.Content.ReadAsStringAsync().Result;

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
