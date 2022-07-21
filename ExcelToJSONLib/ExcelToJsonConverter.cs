using System;
using System.IO;
using ExcelDataReader;
using System.Text;
using Newtonsoft.Json;

namespace ExcelToJSONLib
{
    public class ExcelToJsonConverter
    {
        public void JsonConver()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            string inFilePath = @"C:\Users\Выймова Елена\Documents\InputExcel.xlsx";
            var outFilePath = @"C:\Users\Выймова Елена\Documents\OutputJson.json";

            using (var inFile = File.Open(inFilePath, FileMode.Open, FileAccess.Read))
            using (var outFile = File.CreateText(outFilePath))
            {
                using (var reader = ExcelReaderFactory.CreateReader(inFile, new ExcelReaderConfiguration()
                { FallbackEncoding = Encoding.GetEncoding(1252) }))
                using (var writer = new JsonTextWriter(outFile))
                {
                    writer.Formatting = Formatting.Indented;
                    writer.WriteStartArray();
                    reader.Read();

                    do
                    {
                        while (reader.Read())
                        {
                            writer.WriteStartObject();

                            writer.WritePropertyName("Блок");
                            writer.WriteValue(reader.GetString(0));

                            writer.WritePropertyName("Организация");
                            writer.WriteValue(reader.GetString(1));

                            writer.WritePropertyName("Подразделение БП");
                            writer.WriteValue(reader.GetString(2));

                            writer.WritePropertyName("Дирекция");
                            writer.WriteValue(reader.GetString(3));

                            writer.WritePropertyName("Департамент");
                            writer.WriteValue(reader.GetString(4));

                            writer.WritePropertyName("Функция");
                            writer.WriteValue(reader.GetString(5));

                            writer.WritePropertyName("Управление/Служба");
                            writer.WriteValue(reader.GetString(6));

                            writer.WritePropertyName("Направление");
                            writer.WriteValue(reader.GetString(7));

                            writer.WritePropertyName("Отдел");
                            writer.WriteValue(reader.GetString(8));

                            writer.WritePropertyName("Сектор");
                            writer.WriteValue(reader.GetString(9));

                            writer.WritePropertyName("Должность");
                            writer.WriteValue(reader.GetString(10));

                            writer.WritePropertyName("Сотрудник");
                            writer.WriteValue(reader.GetString(11));

                            writer.WritePropertyName("Сотрудник_Код");
                            writer.WriteValue(reader.GetString(12));

                            writer.WritePropertyName("ОрганизацияКод");
                            writer.WriteValue(reader.GetString(13));

                            writer.WritePropertyName("Код Родителя");
                            writer.WriteValue(reader.GetString(14));

                            writer.WritePropertyName("Подразделение_Код");
                            writer.WriteValue(reader.GetString(15));

                            writer.WritePropertyName("Конечное подразделение");
                            writer.WriteValue(reader.GetString(16));

                            writer.WritePropertyName("Почта");
                            writer.WriteValue(reader.GetString(17));

                            writer.WriteEndObject();
                        }
                    } 
                    while (reader.NextResult());

                    writer.WriteEndArray();
                }
            }

        }
    }
}
