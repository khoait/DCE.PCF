using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;

namespace DCE.PCF.Plugins.Utils
{
    public static class JsonConvert
    {
        public static string Serialize<T>(T obj)
        {
            using (var stream = new MemoryStream())
            {
                GetSerializer<T>().WriteObject(stream, obj);
                return Encoding.UTF8.GetString(stream.ToArray());
            }
        }

        public static T Deserialize<T>(string json)
        {
            using (var stream = new MemoryStream())
            {
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write(json);
                    writer.Flush();
                    stream.Position = 0;
                    return (T)GetSerializer<T>().ReadObject(stream);
                }
            }
        }

        private static DataContractJsonSerializer GetSerializer<T>()
        {
            var settings = new DataContractJsonSerializerSettings
            {
                UseSimpleDictionaryFormat = true
            };

            return new DataContractJsonSerializer(typeof(T), settings);
        }
    }
}
