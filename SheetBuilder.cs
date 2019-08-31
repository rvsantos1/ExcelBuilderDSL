using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelBuilderDSL.Excel
{

    public interface ISheetBuilder
    {
    	ISheetBuilder WithName(string name);
    	ISheetBuilder WithNewLine();
    	ISheetBuilder WithColumnValue();
        SheetClass Build();
    }

    public class SheetClass{
        public string Name { get; set; }
        public Worksheet Sheet { get; set; }
    }

    public class SheerBuilder : ISheetBuilder
    {

        private SheerBuilder()
        {
            
        }

        private string Name { get; set; }

        public static ISheetBuilder Builder()
        {
            return new SheerBuilder();
        }

        public SheetClass Build()
        {
            SheetData dataSheet = new SheetData();

            return new SheetClass{
                    Name = this.Name,
                    Sheet = new Worksheet(dataSheet)
            } ;
        }

        public ISheetBuilder WithColumnValue()
        {
            throw new System.NotImplementedException();
        }

        public ISheetBuilder WithName(string name)
        {
            this.Name = name;
            return this;
        }

        public ISheetBuilder WithNewLine()
        {
            throw new System.NotImplementedException();
        }
    }




}
