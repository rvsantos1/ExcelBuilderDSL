using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelBuilderDSL.Excel
{

    public interface IExcelBuilder
    {
        IExcelBuilder WithPath(string path);
        IExcelBuilder WithOverrideFile();
        IExcelBuilder WithNonOverrideFile();
        IExcelBuilder WithSheet(SheetClass sheet);
        void Build();
    }
    public class ExcelBuilder : IExcelBuilder
    {
        private string Path { get; set; }
        public bool OverrideFile { get; private set; }
        private Queue Sheets { get; set; }

        private ExcelBuilder()
        {
            OverrideFile = true;
            Sheets = new Queue();
        }

        public static IExcelBuilder Builder()
        {
            return new ExcelBuilder();
        }


        public IExcelBuilder WithPath(string path)
        {
            this.Path = path;
            return this;
        }

        public IExcelBuilder WithOverrideFile()
        {
            this.OverrideFile = true;
            return this;
        }

        public IExcelBuilder WithNonOverrideFile()
        {
            this.OverrideFile = false;
            return this;
        }

        public IExcelBuilder Save()
        {
            return this;
        }

        public IExcelBuilder Close()
        {
            return this;
        }

        public IExcelBuilder WithSheet(SheetClass sheet)
        {
            Sheets.Enqueue(sheet);
            return this;
        }


        public void Build()
        {
            //Check If File Exist
            if (File.Exists(Path))
            {
                if (OverrideFile)
                    File.Delete(Path);
                else
                    throw new FileLoadException("File already exist, use WithOverride method for override exist file");
            }

            //Abre o documento para uso
            using (var spreadsheetDocument = SpreadsheetDocument.
                    Create(Path, SpreadsheetDocumentType.Workbook))
            {

                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add Sheets to the Workbook.
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                    AppendChild<Sheets>(new Sheets());

                ///----------------///
                ///WorkSheet Logics///
                ///----------------///    
                
                uint i =1;
                while ( Sheets.Count > 0)
                {
                    var st = Sheets.Dequeue() as SheetClass;
                    WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = st.Sheet;
                    sheets.Append(new Sheet() { Id = spreadsheetDocument.WorkbookPart.
                     GetIdOfPart(worksheetPart), SheetId = i, Name = st.Name });
                     i++;
                }



                workbookpart.Workbook.Save();



            }



        }

    }
}