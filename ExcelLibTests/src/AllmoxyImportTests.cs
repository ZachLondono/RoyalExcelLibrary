using Microsoft.VisualStudio.TestTools.UnitTesting;
using RoyalExcelLibrary;
using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Providers;
using RoyalExcelLibrary.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLibTests {

	[TestClass]
	public class AllmoxyImportTests {

		[TestMethod]
		public void TestImport() {

			string filepath = "C:\\Users\\Zachary Londono\\source\\repos\\RoyalExcelLibrary\\ExcelLibTests\\Test Data\\TestExport.xml";

			AllmoxyOrderProvider provider = new AllmoxyOrderProvider(filepath);

			Order order = provider.LoadCurrentOrder();

			Assert.AreEqual("Export Order", order.Job.Name);

		}

	}

	public class DrawerBoxServiceTest {

		[TestMethod]
		public void TestFraction() {
			string fraction = HelperFuncs.FractionalImperialDim(1.125 * 25.4);
			Assert.AreEqual("1 1/8", fraction);
		}

	}

}
