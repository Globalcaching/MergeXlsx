using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace MergeXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            Settings settings;

            var serializer = new XmlSerializer(typeof(Settings));
            using (var stream = new StringReader(File.ReadAllText(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "settings.xml"))))
            using (var reader = XmlReader.Create(stream))
            {
                settings = (Settings)serializer.Deserialize(reader);
            }

            var converter = new Converter(settings);
            converter.Execute();
        }
    }
}
