﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using OfficeOpenXml.LoadFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.LoadFunctions
{
    [TestClass]
    public class LoadFromCollectionDictionaryPropertyTests
    {
        [EpplusTable]
        public class TestClass
        {
            [EpplusTableColumn(Order = 3)] 
            public string Name { get; set; }

            [EPPlusDictionaryColumn(Order = 2, HeadersKey = "1")]
            public Dictionary<string, object> Columns { get; set; }
        }

        public class TestClass2 : TestClass
        {
            [EPPlusDictionaryColumn(Order = 1, HeadersKey = "2")]
            public Dictionary<string, object> Columns2 { get; set; }
        }

        public class MyKeysProvider : IDictionaryKeysProvider
        {
            public IEnumerable<string> GetKeys(string key)
            {
                switch(key)
                {
                    case "1":
                        return new string[] { "A", "B", "C" };
                    case "2":
                        return new string[] { "C", "D", "E" };
                    default:
                        return Enumerable.Empty<string>();
                }
            }
        }

        [TestMethod]
        public void ShouldReadColumnsAndValuesFromDictionaryProperty()
        {
            var item1 = new TestClass
            {
                Name = "test 1",
                Columns = new Dictionary<string, object> { { "A", 1 }, { "B", 2 } }
            };
            var items = new List<TestClass> { item1 };
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].LoadFromCollection(items, c =>
                {
                    c.PrintHeaders = true;
                    c.KeysProvider = new MyKeysProvider();
                });
                Assert.AreEqual("C", sheet.Cells["A1"].Value);
                Assert.AreEqual("D", sheet.Cells["B1"].Value);
                Assert.AreEqual("E", sheet.Cells["C1"].Value);
                Assert.AreEqual("Name", sheet.Cells["D1"].Value);
                Assert.AreEqual(1, sheet.Cells["A2"].Value);
                Assert.AreEqual(2, sheet.Cells["B2"].Value);
                Assert.IsNull(sheet.Cells["C2"].Value);
                Assert.AreEqual("test 1", sheet.Cells["D2"].Value);
            }
        }

        [TestMethod]
        public void ShouldReadColumnsAndValuesFromDictionaryProperty2()
        {
            var item1 = new TestClass2
            {
                Name = "test 1",
                Columns = new Dictionary<string, object> { { "A", 3 } },
                Columns2 = new Dictionary<string, object> { { "C", 1 }, { "D", 2 } }
            };
            var items = new List<TestClass2> { item1 };
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].LoadFromCollection(items, c =>
                {
                    c.PrintHeaders = true;
                    c.KeysProvider = new MyKeysProvider();
                });
                Assert.AreEqual("C", sheet.Cells["A1"].Value);
                Assert.AreEqual("D", sheet.Cells["B1"].Value);
                Assert.AreEqual("E", sheet.Cells["C1"].Value);
                Assert.AreEqual("A", sheet.Cells["D1"].Value);
                Assert.AreEqual("B", sheet.Cells["E1"].Value);
                Assert.AreEqual("C", sheet.Cells["F1"].Value);
                Assert.AreEqual("Name", sheet.Cells["G1"].Value);
                Assert.AreEqual(1, sheet.Cells["A2"].Value);
                Assert.AreEqual(2, sheet.Cells["B2"].Value);
                Assert.IsNull(sheet.Cells["C2"].Value);
                Assert.AreEqual(2, sheet.Cells["D2"].Value);
                Assert.IsNull(sheet.Cells["E2"].Value);
                Assert.AreEqual("test 1", sheet.Cells["G2"].Value);
            }
        }
    }
}
