using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeXlsx
{
    public class Converter
    {
        private Settings _settings;
        private List<string> _sourceFiles;
        private IWorkbook _workbook;

        public Converter(Settings settings)
        {
            _settings = settings;
            _sourceFiles = new List<string>();
        }

        public void Execute()
        {
            //set destination
            if (File.Exists(_settings.ApplicationSettings.FileLocations.Destination))
            {
                File.Delete(_settings.ApplicationSettings.FileLocations.Destination);
            }

            //get source files
            if (File.Exists(_settings.ApplicationSettings.FileLocations.Source))
            {
                _sourceFiles.Add(_settings.ApplicationSettings.FileLocations.Source);
            }
            else if (Directory.Exists(_settings.ApplicationSettings.FileLocations.Source))
            {
                GetSourceFiles(_settings.ApplicationSettings.FileLocations.Source);
            }


            _workbook = new XSSFWorkbook();

            //create sheets
            foreach (var sheet in _settings.Sheets)
            {
                if (_workbook.GetSheet(sheet.Name) == null)
                {
                    _workbook.CreateSheet(sheet.Name);
                }
            }

            //add columns
            foreach (var sheet in _settings.Sheets)
            {
                var wbsheet = _workbook.GetSheet(sheet.Name);
                sheet.ItemCount = 0;
                foreach (var column in sheet.Elements.Columns)
                {
                    ICell cell = wbsheet.GetOrCreateCell(sheet.HeaderRow, column.HeaderCol);
                    cell.SetCellValue(column.Header);
                    wbsheet.SetColumnWidth(column.HeaderCol, column.Width);
                }
            }

            var dateStyle = _workbook.CreateCellStyle();
            dateStyle.DataFormat = _workbook.CreateDataFormat().GetFormat("dd-MM-yyyy");

            //var txtStyle = _workbook.CreateCellStyle();
            //txtStyle.WrapText = true;

            //run through all documents
            for (int i = 0; i < _sourceFiles.Count; i++)
            {
                Console.Write(string.Format("{0}...", Path.GetFileName(_sourceFiles[i])));
                var srcworkbook = WorkbookFactory.Create(_sourceFiles[i]);
                foreach (var sheet in _settings.Sheets)
                {
                    var wbsheet = _workbook.GetSheet(sheet.Name);
                    var srcwbsheet = srcworkbook.GetSheet(sheet.Name);
                    if (srcwbsheet != null)
                    {
                        //find first row by searching first column header name
                        var allheaders = (from a in sheet.Elements.Columns select a.Header.ToLower()).ToArray();
                        var indexes = new int?[allheaders.Length];
                        var rowenum = srcwbsheet.GetRowEnumerator();
                        while (rowenum.MoveNext())
                        {
                            var r = rowenum.Current as XSSFRow;
                            IRow ActiveRow = null;
                            var searchHeader = (from a in indexes where a != null select a).FirstOrDefault() == null;
                            if (!searchHeader)
                            {
                                ActiveRow = wbsheet.GetOrCreateRow(sheet.ItemCount + sheet.HeaderRow + 1);
                                sheet.ItemCount++;
                            }
                            foreach (var c in r.Cells)
                            {
                                if (searchHeader)
                                {
                                    var index = Array.IndexOf(allheaders, (c.StringCellValue ?? "").ToLower());
                                    if (index >= 0)
                                    {
                                        indexes[index] = c.ColumnIndex;
                                    }
                                }
                                else
                                {
                                    var col = Array.IndexOf(indexes, c.ColumnIndex);
                                    if (col >= 0)
                                    {
                                        var cell = ActiveRow.GetOrCreateCol(sheet.Elements.Columns[col].HeaderCol);
                                        switch (c.CellType)
                                        {
                                            case CellType.Numeric:
                                                try
                                                {
                                                    cell.SetCellValue(c.DateCellValue);
                                                    cell.CellStyle = dateStyle;
                                                }
                                                catch
                                                {
                                                    try
                                                    {
                                                        cell.SetCellValue(c.NumericCellValue);
                                                    }
                                                    catch
                                                    {
                                                    }
                                                }
                                                break;
                                            case CellType.Boolean:
                                                cell.SetCellValue(c.BooleanCellValue);
                                                break;
                                            case CellType.String:
                                            default:
                                                cell.SetCellValue(c.StringCellValue);
                                                break;
                                        }
                                        cell.SetCellType(c.CellType);
                                    }
                                }
                            }
                        }
                    }
                }
                Console.WriteLine("gedaan");
            }

            using (FileStream stream = new FileStream(_settings.ApplicationSettings.FileLocations.Destination, FileMode.Create, FileAccess.Write))
            {
                _workbook.Write(stream);
            }
        }

        public void GetSourceFiles(string path)
        {
            _sourceFiles.AddRange(Directory.GetFiles(path, "*.xlsx"));
            var dirs = Directory.GetDirectories(path);
            foreach (var d in dirs)
            {
                GetSourceFiles(d);
            }
        }

    }
}
