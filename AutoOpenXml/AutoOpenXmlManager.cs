using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;

namespace AutoOpenXml
{
    public class AutoOpenXmlManager<T> : BaseWritterManager<T>  where T : new()
    {
        public AutoOpenXmlManager<T> Init()
        {
            Workbook = new XLWorkbook();

            Properties = ExtractReferenceMapedProperties();

            var worksheetName = ExtractWorksheetName();

            Workbook.Worksheets.Add(worksheetName);
            Workbook.Worksheets.TryGetWorksheet(worksheetName, out var activeWorkshet);
            ActiveWorksheet = activeWorkshet;

            return this;
        }

        public AutoOpenXmlManager<T> SetData(List<T> data)
        {
            Data = data;
            return this;
        }

        internal string ExtractWorksheetName()
        {
            return (string) (new T())
                .GetType()
                .CustomAttributes
                .First(x => x.AttributeType == typeof(ExportWorkSheetAttribute))
                .ConstructorArguments[0].Value;
        }

        internal List<PropertyInfo> ExtractReferenceMapedProperties()
        {
            return (new T())
                .GetType()
                .GetProperties()
                .Where(X => X.HasValidAttributeField())
                .ToList();
        }

        public MemoryStream StartProcess()
        {
            SortFields();
            ProcessData();
            WriteFile();

            MemoryStream stream = new MemoryStream();
            Workbook.SaveAs(stream);
            return stream;
        }
    }
}
