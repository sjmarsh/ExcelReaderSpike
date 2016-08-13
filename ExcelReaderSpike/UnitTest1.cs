using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using OfficeOpenXml;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelReaderSpike
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var fileInfo = new FileInfo(@".\Data\spike.xlsx");
            var excel = new ExcelPackage(fileInfo);

            var sheet = excel.Workbook.Worksheets.First();

            var spikes = new ExcelReader().ReadIntoList<Spike>(sheet, true);

            foreach(var spike in spikes)
            {
                Console.WriteLine(spike);
            }
        }
    }
    
    public class Spike
    {
        [ExcelColumn("A")]
        public int ColumnA { get; set; }

        [ExcelColumn("B")]
        public string ColumnB { get; set; }

        [ExcelColumn("C")]
        public DateTime ColumnC { get; set; }

        public string DontMapThisColumnBecauseItDoesntHaveAnAttribute { get; set; }

        public override string ToString()
        {
            return string.Format("ColumnA: {0}, ColumnB: {1}, ColumnC: {2}", ColumnA, ColumnB, ColumnC);
        }
    }
}
