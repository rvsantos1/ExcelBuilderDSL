using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelBuilderDSL.Excel
{

    public interface ISheetBuilder
    {
        ISheetBuilder WithName(string name);
        ISheetBuilder WithNewLine();
        ISheetBuilder WithColumnValue(object value);
        SheetClass Build();
    }

    public class SheetClass
    {
        public string Name { get; set; }
        public Worksheet Sheet { get; set; }
    }

    public class SheetBuilder : ISheetBuilder
    {

        private Queue<Action> actions;
        private uint rowNumber = 0;
        private uint columnNumber = 1;
        private SheetData _dataSheet;

        private SheetBuilder()
        {
            _dataSheet = new SheetData();
            actions = new Queue<Action>();
        }

        private string Name { get; set; }

        public static ISheetBuilder Builder()
        {
            return new SheetBuilder();
        }

        public SheetClass Build()
        {
            while (actions.Count > 0)
            {
                actions.Dequeue()();
            }
            return new SheetClass
            {
                Name = this.Name,
                Sheet = new Worksheet(_dataSheet)
            };
        }

        public ISheetBuilder WithColumnValue(object value)
        {
            actions.Enqueue(() =>
            {

                var row = new Row() { RowIndex = rowNumber };
                var cell = new Cell()
                {
                    CellReference = Util.GetCharColumn(columnNumber) + rowNumber,
                    CellValue = new CellValue(value.ToString()),
                    DataType = new EnumValue<CellValues>(CellValues.String)
                };
                row.Append(cell);
                _dataSheet.Append(row);
                columnNumber++;
            });
            return this;
        }

        public ISheetBuilder WithName(string name)
        {
            this.Name = name;
            return this;
        }

        public ISheetBuilder WithNewLine()
        {
            Action action = () =>
            {
                rowNumber++;
                columnNumber = 1;
            };
            actions.Enqueue(action);
            return this;
        }
    }
}
