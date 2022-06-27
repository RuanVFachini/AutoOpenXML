using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;

namespace AutoOpenXml
{
    public class MultiExportManagerBuilder
    {
        internal IXLWorkbook Workbook { get; set; }

        private IList<Action> _queue = new List<Action>();

        public MultiExportManagerBuilder Init()
        {
            Workbook = new XLWorkbook();
            return this;
        }
        
        public MultiExportManagerBuilder SetData<T>(IList<T> data)
            where T : new()
        {
            _queue.Add(GenerateExportWorshet(data));
            return this;
        }

        private Action GenerateExportWorshet<T>(IList<T> data)
            where T : new()
        {
            return () =>
            {
                new ExportManagerBuilder<T>()
                    .Init()
                    .UseWorkBook(Workbook)
                    .SetData(data)
                    .StartExportProcess();
            };
        }

        public MemoryStream StartExportProcess()
        {
            foreach (var action in _queue)
            {
                action.Invoke();
            }
            
            MemoryStream stream = new MemoryStream();
            Workbook.SaveAs(stream);
            return stream;
        }

    }
}