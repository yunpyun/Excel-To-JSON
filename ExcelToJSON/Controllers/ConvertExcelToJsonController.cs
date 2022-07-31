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
    [ApiController]
    public class ConvertExcelToJsonController : ControllerBase
    {
        [HttpGet]
        [Route("[controller]")]
        public string Get()
        {
            return "result get ConvertExcelToJson";
        }

        [HttpPost]
        [Route("[controller]/poststring")]
        public string PostString()
        {
            HttpRequest context = Request;
            string result = Request.Form["data"];
            return "result poststring" + " " + context.ContentType.ToString() + " " + result;
        }

        [HttpPost]
        [Route("[controller]")]
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

                    FormUrlEncodedContent str = new FormUrlEncodedContent(new[]
                    {
                        new KeyValuePair<string, string>("data", sw),
                    });

                    HttpResponseMessage response = await client.PostAsync(url, str);

                    string result = response.Content.ReadAsStringAsync().Result;

                    return StatusCode(200, result);
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
