using FluentAssertions;
using NUnit.Framework;
using RoyalExcelLibrary.ExcelUI.ExportFormat;
using RoyalExcelLibrary.ExcelUI.Models.Products;
using RoyalExcelLibrary.ExcelUI.Providers;
using System.Linq;

namespace ExcelLibrary.Tests.Unit {
    internal class AllmoxyProviderTests {
    
        private AllmoxyOrderProvider _sut { get; set; }
        private readonly string _basePath = "C:\\Users\\Zachary Londono\\source\\repos\\RoyalExcelLibrary\\tests\\ExcelLibrary.Tests\\Unit\\TestData\\";

        [SetUp]
        public void Setup() {
            _sut = new AllmoxyOrderProvider();
        }

        [Test]
        [TestCase("AllmoxyTest1.xml", "2019", "942.94", "62.47", "0", 22, "80 GEORGE ST", "", "PATERSON", "NJ", "07503")]
        [TestCase("AllmoxyTest2.xml", "2040", "515.88", "0", "0", 11, "7248 Camino Degrazia", "unit 293", "san diego", "CA", "92111")]
        [TestCase("AllmoxyTest3.xml", "2039", "871.91", "0", "0", 20, "Pickup", "", "", "", "")]
        [TestCase("AllmoxyTest4.xml", "2038", "1,846.96", "122.36", "0", 26, "Pickup", "", "", "", "")]
        [TestCase("AllmoxyTest5.xml", "2037", "299.22", "19.82", "0", 4, "Pickup", "", "", "", "")]
        public void Should_LoadOrder_WhenFileIsValidOrder(string filePath, 
                                                            string expectedNumber,
                                                            decimal expectedSubTotal,
                                                            decimal expectedTax,
                                                            decimal expectedShipping,
                                                            int expectedProdCount,
                                                            string expectedAddressLine1,
                                                            string expectedAddressLine2,
                                                            string expectedCity,
                                                            string expectedState,
                                                            string expectedZip) {
            _sut.FilePath = _basePath + filePath;

            var order = _sut.LoadCurrentOrder();

            order.Should().NotBeNull();
            order.Number.Should().Be(expectedNumber);
            order.SubTotal.Should().Be(expectedSubTotal);
            order.Tax.Should().Be(expectedTax);
            order.ShippingCost.Should().Be(expectedShipping);
            order.Products.Sum(p => p.Qty).Should().Be(expectedProdCount);
            order.Products.Count(p => (p as DrawerBox).SideMaterial == RoyalExcelLibrary.ExcelUI.Models.MaterialType.Unknown).Should().Be(0);

            order.Customer.Address.Should().BeEquivalentTo(new {
                Line1 = expectedAddressLine1,
                Line2 = expectedAddressLine2,
                City = expectedCity,
                State = expectedState,
                Zip = expectedZip
            });

        }

    }

}
