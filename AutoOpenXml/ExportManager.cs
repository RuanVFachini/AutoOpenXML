﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using AutoOpenXml.Exceptions;
using AutoOpenXml.Extensions;
using ClosedXML.Excel;

namespace AutoOpenXml
{
    public class ExportManager<T> : WritterManager<T>  where T : new()
    {
        internal IXLWorkbook Workbook { get; set; }
        
        public ExportManager<T> Init()
        {
            Workbook = new XLWorkbook();

            Properties = CommonManager.ExtractReferenceMapedProperties<T>();

            var worksheetName = CommonManager.ExtractWorksheetName<T>();

            Workbook.Worksheets.Add(worksheetName);
            Workbook.Worksheets.TryGetWorksheet(worksheetName, out var activeWorkshet);
            ActiveWorksheet = activeWorkshet;

            return this;
        }

        public ExportManager<T> SetData(List<T> data)
        {
            if (data.Count == 0) throw new EmptyObjectListToExportException();

            Data = data;
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
