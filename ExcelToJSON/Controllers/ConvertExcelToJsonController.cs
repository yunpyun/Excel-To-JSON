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
        /// <summary>
        /// Метод для тестирования доступа к сервису
        /// </summary>
        /// <returns>Строка со статичным значением</returns>
        [HttpGet]
        [Route("[controller]")]
        public string Get()
        {
            return "result get ConvertExcelToJson";
        }


        /// <summary>
        /// Метод для тестирования, его можно вызвать из метода Post, указав его url
        /// </summary>
        /// <returns>Строка, которую передал Post</returns>
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

        /// <summary>
        /// Метод, вызываемый в активити, куда передается файл и url активити. Асинхронно вызывает метод Convert и продолжает выполнение
        /// </summary>
        /// <returns>Статус выполнения 200</returns>
        [HttpPost]
        [Route("[controller]")]
        public async Task<IActionResult> Post()
        {
            IFormFile file = Request.Form.Files.FirstOrDefault();
            string url = Request.Form["url"];

            // вызов асинхронного метода
            Convert(url, file);

            return StatusCode(200);
        }

        /// <summary>
        /// Асинхронная функция для параллельного выполнения данной функции и вызвавшего ее метода Post
        /// </summary>
        /// <param name="url">Url активити, которое нужно вызвать и вернуть результат</param>
        /// <param name="file">Исходный excel</param>
        /// <returns>Выполняет вызов активити и отправляет на него результат выполнения конвертирования</returns>
        private async Task Convert(string url, IFormFile file)
        {
            // переменная окружения для использования токена стенда ELMA365
            string token = Environment.GetEnvironmentVariable("TOKEN_ELMA365");

            // для SSL
            HttpClientHandler clientHandler = new HttpClientHandler
            {
                ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => { return true; }
            };

            // создание объекта для его последующего вызова и отправки ему результата
            using var client = new HttpClient(clientHandler);

            // авторизация с учетом токена
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            if (file != null)
            {
                try
                {
                    var filePath = Path.GetTempPath();
                    var fileName = filePath + Guid.NewGuid().ToString() + "-" + file.FileName;

                    using (var stream = System.IO.File.Create(fileName))
                    {
                        file.CopyTo(stream);
                    }

                    // запуск библиотеки
                    ExcelToJsonConverter converter = new ExcelToJsonConverter();
                    string sw = converter.JsonConvert(fileName);

                    // формирование объекта строки для отправки результата клиенту
                    StringContent str = new StringContent(sw);

                    // удаление из временной папки входного excel-файла
                    FileInfo fileInf = new FileInfo(fileName);

                    if (fileInf.Exists)
                    {
                        fileInf.Delete();
                    }

                    // отправка строкового результата клиенту
                    HttpResponseMessage response = await client.PostAsync(url, str);
                }
                catch (Exception err)
                {
                    // формирование и отправка клиенту строки с текстом ошибки
                    StringContent errMess = new StringContent("Ошибка выполнения: " + err.Message);
                    await client.PostAsync(url, errMess);
                }
            }
            else
            {
                // формирование и отправка клиенту строки с текстом ошибки
                StringContent notFound = new StringContent("Входные файлы не найдены");
                await client.PostAsync(url, notFound);
            }
        }
    }
}
