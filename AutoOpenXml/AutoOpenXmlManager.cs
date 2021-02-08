using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using AutoOpenXml.Exceptions;
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
            if (data.Count == 0) throw new EmptyObjectListToExportException();

            Data = data;
            return this;
        }

        internal string ExtractWorksheetName()
        {
            var validClass = (new T())
                .GetType()
                .CustomAttributes
                .FirstOrDefault(x => x.AttributeType == typeof(ExportWorkSheetAttribute));

            if (validClass != null) return (string) validClass.ConstructorArguments[0].Value;

            throw new MissingExportWorkSheetAttributeException();
        }

        internal List<PropertyInfo> ExtractReferenceMapedProperties()
        {
            var properties = (new T())
                .GetType()
                .GetProperties()
                .Where(X => X.HasValidAttributeField())
                .ToList();

            if (properties != null) return properties;

            throw new MissingExportColumnAttributeException();
        }

        public MemoryStream StartExportProcess()
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
