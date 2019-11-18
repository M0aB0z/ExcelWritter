using System;

namespace ExcelWritter
{
    class Program
    {
        static void Main()
        {
            var def = new WorkBookDefinition
            {
                ColumnsNames = new[] { "ColA", "ColB", "ColC" },
                Rows = new[]{
                    new []{ "A1","B1","C1"},
                    new []{ "A2","B2","C2"},
                },
                SheetName = "MySheetName",
                FileName = @"C:\temp\Sample.xlsx"
            };

            var engine = new ExcelHelper();
            engine.Write(def);
        }
    }
}
