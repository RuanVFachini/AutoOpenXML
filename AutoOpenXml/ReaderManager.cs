using AutoOpenXml.Models;
using ClosedXML.Excel;

namespace AutoOpenXml
{
    public class ReaderManager<T> where T : new()
    {
        internal IXLWorksheet ActiveWorksheet { get; set; }
        private int CurrentRowIndex { get; set; } = 1;

        internal bool NextRow()
        {
            if (IsLastRow()) return false;
            CurrentRowIndex++;
            return true;
            
        }

        private bool IsLastRow()
        {
            return ActiveWorksheet.LastRowUsed().RowNumber() - CurrentRowIndex == 0;
        }

        internal CellRead TryReadColumnInfo(int index)
        {
            var cell = ActiveWorksheet.Cell(CurrentRowIndex, index);
            return new CellRead()
            {
                Type = cell.DataType,
                Value = (cell.HasFormula ? cell.CachedValue : cell.Value) ?? null
            };
        }
    }
}
