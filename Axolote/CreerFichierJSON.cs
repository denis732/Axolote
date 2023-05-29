using Newtonsoft.Json;
using System.Text;
using Axolote.Models;
using System;
using System.IO;
using ExcelDataReader;
using OfficeOpenXml;


namespace EspaceJSON
{
    class Program
    {
        public static object CodePagesEncodingProvider { get; private set; }

        public static void ConvertirExcelEnJson(string args)
        {
            //Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var inFilePath = args;
            var outFilePath = inFilePath +  "Json";
            
            using (var inFile = File.Open(inFilePath, FileMode.Open, FileAccess.Read))
            using (var outFile = File.CreateText(outFilePath))
            {
                using (var reader = ExcelReaderFactory.CreateReader(inFile, new ExcelReaderConfiguration()
                { FallbackEncoding = Encoding.GetEncoding(1252) }))
                using (var writer = new JsonTextWriter(outFile))
                {
                    writer.Formatting = Formatting.Indented; //I likes it tidy
                    writer.WriteStartArray();
                    reader.Read(); //SKIP FIRST ROW, it's TITLES.
                    do
                    {
                        while (reader.Read())
                        {
                            //peek ahead? Bail before we start anything so we don't get an empty object
                            var varSCTID = reader.GetString(1);
                            if (string.IsNullOrEmpty(varSCTID)) break;

                            writer.WriteStartObject();
                            writer.WritePropertyName("Source");
                            writer.WriteValue(reader.GetString(0));

                            writer.WritePropertyName("Code");
                            writer.WriteValue(varSCTID);

                            writer.WritePropertyName("Valeur");
                            writer.WriteValue(reader.GetString(2));

                            writer.WritePropertyName("AffichageAnglais");
                            writer.WriteValue(reader.GetString(3));

                            writer.WritePropertyName("AffichageFrancais");
                            writer.WriteValue(reader.GetString(4));

                            /*
                            <iframe src="https://channel9.msdn.com/Shows/Azure-Friday/Erich-Gamma-introduces-us-to-Visual-Studio-Online-integrated-with-the-Windows-Azure-Portal-Part-1/player" width="960" height="540" allowFullScreen frameBorder="0"></iframe>
                             */

                            writer.WriteEndObject();
                        }
                    } while (reader.NextResult());
                    writer.WriteEndArray();
                }
            }
        }
    }
}

class Fichier
{

}

//class PointConverter : JsonConverter
//{
//    public override bool CanConvert(Type objectType)
//    {
//        return objectType == typeof(FichierJson);
//    }

//    public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
//    {
//        JObject jo = JObject.Load(reader);
//        FichierJson FJson = new FichierJson
//        {
//            X = (int)jo["x"],
//            Y = (int)jo["y"]
//        };
//        return FJson;
//    }

//    public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
//    {
//        FichierJson point = (FichierJson)value;
//        JObject jo = new JObject
//        {
//            {"x", point.X },
//            {"y", point.Y }
//        };
//        jo.WriteTo(writer);
//    }
//}

//var p = new Product("Product A", new DateTime(2021, 12, 28),
//    new string[] { "small" });

//var json = JsonConvert.SerializeObject(p);
