using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace MergeXlsx
{
    [XmlType("settings")]
    public class Settings
    {
        [XmlElement("application")]
        public ApplicationSettings ApplicationSettings { get; set; }

        [XmlArray("sheets")]
        [XmlArrayItem("sheet")]
        public List<Sheet> Sheets { get; set; }
    }

    public class ApplicationSettings
    {
        [XmlElement("locations")]
        public FileLocations FileLocations { get; set; }
    }

    public class FileLocations
    {
        [XmlElement("source")]
        public string Source { get; set; }

        [XmlElement("destination")]
        public string Destination { get; set; }
    }

    public class Sheet
    {
        public Sheet()
        {
        }

        [XmlElement("name")]
        public string Name { get; set; }

        [XmlElement("headerrow")]
        public int HeaderRow { get; set; }

        [XmlElement("elements")]
        public Elements Elements { get; set; }

        public int ItemCount { get; set; }
    }

    public class Elements
    {
        public Elements()
        {
        }

        [XmlArray("columns")]
        [XmlArrayItem("column")]
        public List<Column> Columns { get; set; }
    }

    public class Column
    {
        public Column()
        {
        }

        [XmlElement("header")]
        public string Header { get; set; }

        [XmlElement("headercol")]
        public int HeaderCol { get; set; }

        [XmlElement("width")]
        public int Width { get; set; }
    }

}
