using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelToJSONLib;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;

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
        public async Task<string> PostString()
        {
            string data = null;
            using (var streamReader = new StreamReader(Request.Body))
            {
                data = await streamReader.ReadToEndAsync();
            }

            return "result poststring" + ":\n" + data;
        }

        [HttpPost]
        [Route("[controller]")]
        public async Task<IActionResult> Post()
        {
            IFormFile file = Request.Form.Files.FirstOrDefault();
            string url = Request.Form["url"];
            string token = Environment.GetEnvironmentVariable("TOKEN_ELMA365");

            HttpClientHandler clientHandler = new HttpClientHandler();
            clientHandler.ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => { return true; };

            using var client = new HttpClient(clientHandler);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

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

                    StringContent str = new StringContent(sw);

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
