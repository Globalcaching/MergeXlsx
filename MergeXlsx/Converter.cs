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

            var dateStyle = _workbook.CreateCellStyle();
            dateStyle.DataFormat = _workbook.CreateDataFormat().GetFormat("dd-MM-yyyy");

            //add columns
            foreach (var sheet in _settings.Sheets)
            {
                var wbsheet = _workbook.GetSheet(sheet.Name);
                sheet.ItemCount = 0;
                foreach (var column in sheet.Elements.Columns)
                {
                    ICell cell = wbsheet.GetOrCreateCell(sheet.HeaderRow, column.HeaderCol);
                    cell.SetCellValue(column.Header);
                    cell.CellStyle = dateStyle;
                    if (column.Width != 0)
                    {
                        wbsheet.SetColumnWidth(column.HeaderCol, column.Width);
                    }
                    wbsheet.GetRow(sheet.HeaderRow).RowStyle = dateStyle;
                }
            }

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
                                ActiveRow.Height = r.Height;
                            }
                            foreach (var c in r.Cells)
                            {
                                if (searchHeader)
                                {
                                    var index = Array.IndexOf(allheaders, (c.StringCellValue ?? "").ToLower());
                                    if (index >= 0)
                                    {
                                        indexes[index] = c.ColumnIndex;
                                        if (sheet.ItemCount == 0)
                                        {
                                            if (sheet.Elements.Columns[index].Width == 0)
                                            {
                                                wbsheet.SetColumnWidth(sheet.Elements.Columns[index].HeaderCol, Math.Max(srcwbsheet.GetColumnWidth(c.ColumnIndex), wbsheet.GetColumnWidth(sheet.Elements.Columns[index].HeaderCol)));
                                            }
                                            if (r.RowStyle != null)
                                            {
                                                wbsheet.GetRow(sheet.HeaderRow).RowStyle.CloneStyleFrom(r.RowStyle);
                                            }
                                            wbsheet.GetRow(sheet.HeaderRow).GetCell(sheet.Elements.Columns[index].HeaderCol).CellStyle.CloneStyleFrom(c.CellStyle);
                                        }
                                    }
                                }
                                else
                                {
                                    var col = Array.IndexOf(indexes, c.ColumnIndex);
                                    if (col >= 0)
                                    {
                                        var cell = ActiveRow.GetOrCreateCol(sheet.Elements.Columns[col].HeaderCol);
                                        if (r.RowStyle != null)
                                        {
                                            ActiveRow.RowStyle = dateStyle;
                                            ActiveRow.RowStyle.CloneStyleFrom(r.RowStyle);
                                        }
                                        cell.CellStyle = dateStyle;
                                        switch (c.CellType)
                                        {
                                            case CellType.Numeric:
                                                try
                                                {
                                                    cell.SetCellValue(c.DateCellValue);
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
                                        cell.CellStyle.CloneStyleFrom(c.CellStyle);
                                        if (sheet.Elements.Columns[col].Width == 0)
                                        {
                                            wbsheet.SetColumnWidth(sheet.Elements.Columns[col].HeaderCol, Math.Max(srcwbsheet.GetColumnWidth(c.ColumnIndex), wbsheet.GetColumnWidth(sheet.Elements.Columns[col].HeaderCol)));
                                        }
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
