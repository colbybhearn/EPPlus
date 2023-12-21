﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using Castle.DynamicProxy;
namespace EPPlusTest.Table.PivotTable.Calculation
{
    [TestClass]
    public class PivotTableCalculationFilterTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelTable _tbl1, _tbl2;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("PivotTableCalculationFilters.xlsx", true);
            var ws = _pck.Workbook.Worksheets.Add("Data1");
            var r = LoadItemData(ws);
            _tbl1 = ws.Tables.Add(r, "Table1");
            ws = _pck.Workbook.Worksheets.Add("Data2");
            r = LoadItemData(ws);
            _tbl2 = ws.Tables.Add(r, "Table2");
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void FilterPageFieldSingleItemNoGrouping()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotWithPageFieldSingle");
            var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTablePageFieldSingle");
            var pf = pt.PageFields.Add(pt.Fields[0]);
            pf.MultipleItemSelectionAllowed = false;
            pt.CacheDefinition.Refresh();
            var df = pt.DataFields.Add(pt.Fields["Price"]);
            pf.Items.SelectSingleItem(0);

            pt.Calculate();
            Assert.AreEqual(270.6, pt.CalculatedItems[0][Array.Empty<int>()]);
        }
        [TestMethod]
        public void FilterPageFieldMultipleItems()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotWithPageFieldMulti");
            var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTablePageFieldMulti");
            var pf = pt.PageFields.Add(pt.Fields[0]);
            pf.MultipleItemSelectionAllowed = true;
            pt.CacheDefinition.Refresh();
            var df = pt.DataFields.Add(pt.Fields["Price"]);
            pf.Items[3].Hidden = true;
            pf.Items[4].Hidden = true;
            pf.Items[5].Hidden = true;

            pt.Calculate();
            Assert.AreEqual(391.92, pt.CalculatedItems[0][Array.Empty<int>()]);
        }
        [TestMethod]
        public void FilterPageFieldCustomCaptionFilterEquals()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterCapEquals");
            var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableCustomCapEquals");
            var rf = pt.RowFields.Add(pt.Fields[1]);
            pt.CacheDefinition.Refresh();
            var df = pt.DataFields.Add(pt.Fields["Price"]);            
            rf.Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionEqual, "Groceries");
            pt.Calculate();
            Assert.AreEqual(8.4, pt.CalculatedItems[0][[1]]);
            Assert.AreEqual(8.4, pt.CalculatedItems[0][[-1]]);
        }
        [TestMethod]
        public void FilterPageFieldCustomCaptionFilterNotEquals()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterCapNotEquals");
            var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableCustomCapNotEquals");
            var rf = pt.RowFields.Add(pt.Fields[1]);
            pt.CacheDefinition.Refresh();
            var df = pt.DataFields.Add(pt.Fields["Price"]);
            rf.Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionNotEqual, "Groceries");
            pt.Calculate();
            Assert.AreEqual(437.12, pt.CalculatedItems[0][[0]]);
            Assert.AreEqual(437.12, pt.CalculatedItems[0][[-1]]);
        }
        [TestMethod]
        public void FilterPageFieldCustomCaptionFilterGreaterThan()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterCapGreater");
            var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableCustomCapGreater");
            var rf = pt.RowFields.Add(pt.Fields[1]);
            pt.CacheDefinition.Refresh();
            var df = pt.DataFields.Add(pt.Fields["Price"]);
            rf.Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionGreaterThan, "Groceries");
            pt.Calculate();
            Assert.AreEqual(437.12, pt.CalculatedItems[0][[0]]);
            Assert.AreEqual(437.12, pt.CalculatedItems[0][[-1]]);
        }
        [TestMethod]
        public void FilterPageFieldCustomCaptionFilterGreaterEqualThan()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterCapGreaterEq");
            var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableCustomCapGreaterEq");
            var rf = pt.RowFields.Add(pt.Fields[1]);
            pt.CacheDefinition.Refresh();
            var df = pt.DataFields.Add(pt.Fields["Price"]);
            rf.Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionGreaterThanOrEqual, "Groceries");
            pt.Calculate();
            Assert.AreEqual(437.12, pt.CalculatedItems[0][[0]]);
            Assert.AreEqual(8.4, pt.CalculatedItems[0][[1]]);
            Assert.AreEqual(445.52, pt.CalculatedItems[0][[-1]]);
        }
        [TestMethod]
        public void FilterPageFieldCustomCaptionFilterLessThan()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterCapLess");
            var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableCustomCapLess");
            var rf = pt.RowFields.Add(pt.Fields[1]);
            pt.CacheDefinition.Refresh();
            var df = pt.DataFields.Add(pt.Fields["Price"]);
            rf.Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionLessThan, "Hardware");
            pt.Calculate();
            Assert.AreEqual(8.4, pt.CalculatedItems[0][[1]]);
            Assert.AreEqual(8.4, pt.CalculatedItems[0][[-1]]);
        }
        [TestMethod]
        public void FilterPageFieldCustomCaptionFilterLessEqualThan()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterCapLessEq");
            var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableCustomCapLessEq");
            var rf = pt.RowFields.Add(pt.Fields[1]);
            pt.CacheDefinition.Refresh();
            var df = pt.DataFields.Add(pt.Fields["Price"]);
            rf.Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionLessThanOrEqual, "Hardware");
            pt.Calculate();
            Assert.AreEqual(437.12, pt.CalculatedItems[0][[0]]);
            Assert.AreEqual(8.4, pt.CalculatedItems[0][[1]]);
            Assert.AreEqual(445.52, pt.CalculatedItems[0][[-1]]);
        }
        [TestMethod]
        public void FilterPageFieldCustomCaptionFilterBetween()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterCapBetween");
            var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableCapBetween");
            var rf = pt.RowFields.Add(pt.Fields[0]);
            pt.CacheDefinition.Refresh();
            var df = pt.DataFields.Add(pt.Fields["Price"]);
            rf.Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionBetween, "B", "D");
            pt.Calculate();
            Assert.AreEqual(7.2, pt.CalculatedItems[0][[5]]);
            Assert.AreEqual(270.6, pt.CalculatedItems[0][[0]]);
            Assert.AreEqual(277.8, pt.CalculatedItems[0][[-1]]);
        }
        [TestMethod]
        public void FilterPageFieldCustomCaptionFilterNotBetween()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterCapNotBetween");
            var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableCapNotBetween");
            var rf = pt.RowFields.Add(pt.Fields[0]);
            pt.CacheDefinition.Refresh();
            var df = pt.DataFields.Add(pt.Fields["Price"]);
            rf.Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionNotBetween, "B", "D");
            pt.Calculate();
            Assert.AreEqual(88.2, pt.CalculatedItems[0][[1]]);
            Assert.AreEqual(33.12, pt.CalculatedItems[0][[2]]);
            Assert.AreEqual(45.2, pt.CalculatedItems[0][[3]]);
            Assert.AreEqual(1.2, pt.CalculatedItems[0][[4]]);
            Assert.AreEqual(167.72, pt.CalculatedItems[0][[-1]]);
        }
        [TestMethod]
        public void FilterPageFieldCustomCaptionFilterContains()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterCapContains");
            var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableCapContains");
            var rf = pt.RowFields.Add(pt.Fields[1]);
            pt.CacheDefinition.Refresh();
            var df = pt.DataFields.Add(pt.Fields["Price"]);
            rf.Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionContains, "oCer");
            pt.Calculate();
            Assert.AreEqual(8.4, pt.CalculatedItems[0][[1]]);
            Assert.AreEqual(8.4, pt.CalculatedItems[0][[-1]]);
        }
        [TestMethod]
        public void FilterPageFieldCustomCaptionFilterNotContains()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterCapNotContains");
            var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableCapNotContains");
            var rf = pt.RowFields.Add(pt.Fields[1]);
            pt.CacheDefinition.Refresh();
            var df = pt.DataFields.Add(pt.Fields["Price"]);
            rf.Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionNotContains, "wAre");
            pt.Calculate();
            Assert.AreEqual(8.4, pt.CalculatedItems[0][[1]]);
            Assert.AreEqual(8.4, pt.CalculatedItems[0][[-1]]);
        }
        [TestMethod]
        public void FilterPageFieldCustomCaptionFilterBeginsWith()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterCapBeginsWith");
            var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableCapBeginsWith");
            var rf = pt.RowFields.Add(pt.Fields[1]);
            pt.CacheDefinition.Refresh();
            var df = pt.DataFields.Add(pt.Fields["Price"]);
            rf.Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionBeginsWith, "HarD");
            pt.Calculate();
            Assert.AreEqual(437.12, pt.CalculatedItems[0][[0]]);
            Assert.AreEqual(437.12, pt.CalculatedItems[0][[-1]]);
        }
        [TestMethod]
        public void FilterPageFieldCustomCaptionFilterNotBeginsWith()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterCapNotBeginsWith");
            var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableCapNotBeginsWith");
            var rf = pt.RowFields.Add(pt.Fields[1]);
            pt.CacheDefinition.Refresh();
            var df = pt.DataFields.Add(pt.Fields["Price"]);
            rf.Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionNotBeginsWith, "HarD");
            pt.Calculate();
            Assert.AreEqual(8.4, pt.CalculatedItems[0][[1]]);
            Assert.AreEqual(8.4, pt.CalculatedItems[0][[-1]]);
        }
        [TestMethod]
        public void FilterPageFieldCustomCaptionFilterEndsWithWith()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterCapEndsWith");
            var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableCapEndsWith");
            var rf = pt.RowFields.Add(pt.Fields[1]);
            pt.CacheDefinition.Refresh();
            var df = pt.DataFields.Add(pt.Fields["Price"]);
            rf.Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionEndsWith, "wAre");
            pt.Calculate();
            Assert.AreEqual(437.12, pt.CalculatedItems[0][[0]]);
            Assert.AreEqual(437.12, pt.CalculatedItems[0][[-1]]);
        }
        [TestMethod]
        public void FilterPageFieldCustomCaptionFilterNotEndsWith()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterCapNotEndsWith");
            var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableCapNotEndsWith");
            var rf = pt.RowFields.Add(pt.Fields[1]);
            pt.CacheDefinition.Refresh();
            var df = pt.DataFields.Add(pt.Fields["Price"]);
            rf.Filters.AddCaptionFilter(ePivotTableCaptionFilterType.CaptionNotEndsWith, "Ware");
            pt.Calculate();
            Assert.AreEqual(8.4, pt.CalculatedItems[0][[1]]);
            Assert.AreEqual(8.4, pt.CalculatedItems[0][[-1]]);
        }
        [TestMethod]
        public void FilterPageFieldCustomDateEqual()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterDateEqual");
            var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableDateEqual");
            var rf = pt.RowFields.Add(pt.Fields[4]);
            pt.CacheDefinition.Refresh();
            var df = pt.DataFields.Add(pt.Fields["Price"]);
            rf.Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateEqual, new DateTime(2010, 01, 31));
            pt.Calculate();
            Assert.AreEqual(85.2, pt.CalculatedItems[0][[0]]);
            Assert.AreEqual(85.2, pt.CalculatedItems[0][[-1]]);
        }
        [TestMethod]
        public void FilterPageFieldCustomDateBetween()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterDateBetween");
            var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableDateBetween");
            var rf = pt.RowFields.Add(pt.Fields[4]);
            pt.CacheDefinition.Refresh();
            var df = pt.DataFields.Add(pt.Fields["Price"]);
            rf.Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateBetween, new DateTime(2010, 01, 01), new DateTime(2010, 5, 30));
            pt.Calculate();
            Assert.AreEqual(85.2, pt.CalculatedItems[0][[0]]);
            Assert.AreEqual(181.4, pt.CalculatedItems[0][[-1]]);
        }
		[TestMethod]
		public void FilterPageFieldCustomDateNotBetween()
		{
			var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterDateNotBetween");
			var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableDateNotBetween");
			var rf = pt.RowFields.Add(pt.Fields[4]);
			pt.CacheDefinition.Refresh();
			var df = pt.DataFields.Add(pt.Fields["Price"]);
			rf.Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateNotBetween, new DateTime(2010, 01, 01), new DateTime(2010, 5, 30));
			pt.Calculate();
			Assert.AreEqual(173.2, pt.CalculatedItems[0][[4]]);
			Assert.AreEqual(264.12, pt.CalculatedItems[0][[-1]]);
		}
		[TestMethod]
		public void FilterPageFieldCustomDateOlderThan()
		{
			var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterDateDateOlderThanOrEqual");
			var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableDateOlderThan");
			var rf = pt.RowFields.Add(pt.Fields[4]);
			pt.CacheDefinition.Refresh();
			var df = pt.DataFields.Add(pt.Fields["Price"]);
			rf.Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateOlderThan, new DateTime(2010, 5, 31));
			pt.Calculate();
			Assert.AreEqual(85.2, pt.CalculatedItems[0][[0]]);
			Assert.AreEqual(181.4, pt.CalculatedItems[0][[-1]]);
		}
		[TestMethod]
		public void FilterPageFieldCustomDateDateOlderThanOrEqual()
		{
			var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterDateDateOlderThanOrEqual");
			var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableDateDateOlderThanOrEqual");
			var rf = pt.RowFields.Add(pt.Fields[4]);
			pt.CacheDefinition.Refresh();
			var df = pt.DataFields.Add(pt.Fields["Price"]);
			rf.Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateOlderThanOrEqual, new DateTime(2010, 5, 31));
			pt.Calculate();
			Assert.AreEqual(85.2, pt.CalculatedItems[0][[0]]);
			Assert.AreEqual(354.6, pt.CalculatedItems[0][[-1]]);
		}
		[TestMethod]
		public void FilterPageFieldCustomDateNewerThan()
		{
			var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterDateDateNewerThan");
			var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableDateNewerThan");
			var rf = pt.RowFields.Add(pt.Fields[4]);
			pt.CacheDefinition.Refresh();
			var df = pt.DataFields.Add(pt.Fields["Price"]);
			rf.Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateNewerThan, new DateTime(2010, 5, 31));
			pt.Calculate();
			Assert.AreEqual(4.2, pt.CalculatedItems[0][[5]]);
			Assert.AreEqual(90.92, pt.CalculatedItems[0][[-1]]);
		}
		[TestMethod]
		public void FilterPageFieldCustomDateNewerThanOrEqual()
		{
			var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterDateNewerThanOrEqual");
			var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableDateNewerThanOrEqual");
			var rf = pt.RowFields.Add(pt.Fields[4]);
			pt.CacheDefinition.Refresh();
			var df = pt.DataFields.Add(pt.Fields["Price"]);
			rf.Filters.AddDateValueFilter(ePivotTableDateValueFilterType.DateNewerThanOrEqual, new DateTime(2010, 5, 31));
			pt.Calculate();
			Assert.AreEqual(173.2, pt.CalculatedItems[0][[4]]);
			Assert.AreEqual(264.12, pt.CalculatedItems[0][[-1]]);
		}
		[TestMethod]
		public void FilterPageFieldCustomValueEqual()
		{
			var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterValueEqualThan");
			var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableValueEqualThan");
			var rf = pt.RowFields.Add(pt.Fields[4]);
			pt.CacheDefinition.Refresh();
			var df = pt.DataFields.Add(pt.Fields["Price"]);
			rf.Filters.AddValueFilter(ePivotTableValueFilterType.ValueEqual, df, 173.2);
			pt.Calculate();
			Assert.AreEqual(173.2, pt.CalculatedItems[0][[4]]);
			Assert.AreEqual(173.2, pt.CalculatedItems[0][[-1]]);
		}
		[TestMethod]
		public void FilterPageFieldCustomValueNotEqual()
		{
			var ws = _pck.Workbook.Worksheets.Add("PivotCustomFilterValueEqualThan");
			var pt = ws.PivotTables.Add(ws.Cells["C3"], _tbl1, "PivotTableValueEqualThan");
			var rf = pt.RowFields.Add(pt.Fields[0]);
			pt.CacheDefinition.Refresh();
			var df = pt.DataFields.Add(pt.Fields["Price"]);
			rf.Filters.AddValueFilter(ePivotTableValueFilterType.ValueEqual, df, 173.2);
			pt.Calculate();
			Assert.AreEqual(12.2, pt.CalculatedItems[0][[4]]);
			Assert.AreEqual(360.32, pt.CalculatedItems[0][[-1]]);
		}
	}
}