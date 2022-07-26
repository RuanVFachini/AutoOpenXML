using System.Collections.Generic;
using System.IO;
using System.Linq;
using AutoOpenXml.Exceptions;
using AutoOpenXml.Extensions;
using ClosedXML.Excel;

namespace AutoOpenXml
{
    public class ExportManagerBuilder<T> : WritterManager<T>  where T : new()
    {
        internal IXLWorkbook Workbook { get; set; }
        
        public ExportManagerBuilder<T> Init()
        {
            Workbook = new XLWorkbook();
            TableStyle = CommonManager.ExtractTableStyle<T>();

            BaseConfiguration();

            return this;
        }

        private void BaseConfiguration()
        {
            Properties = CommonManager.ExtractReferenceMapedProperties<T>();
            var worksheetName = CommonManager.ExtractWorksheetName<T>();
            Workbook.Worksheets.Add(worksheetName);
            Workbook.Worksheets.TryGetWorksheet(worksheetName, out var activeWorkshet);
            ActiveWorksheet = activeWorkshet;
        }

        public ExportManagerBuilder<T> UseWorkBook(IXLWorkbook workbook)
        {
            Workbook = workbook;

            BaseConfiguration();

            return this;
        }

        public ExportManagerBuilder<T> SetData(IList<T> data)
        {
            if (data.Count == 0) throw new EmptyObjectListToExportException();

            Data = data.ToList();
            return this;
        }

        public MemoryStream StartExportProcess()
        {
            Properties.SortFields();
            ProcessData();
            WriteFile();

            MemoryStream stream = new MemoryStream();
            Workbook.SaveAs(stream);
            return stream;
        }
    }
}
