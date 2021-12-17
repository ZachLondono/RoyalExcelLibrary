using FluentAssertions;
using NUnit.Framework;
using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using RoyalExcelLibrary.Providers;
using System.Linq;

namespace ExcelLibrary.Tests.Unit {
    internal class HafeleProviderTests {

        private HafeleDBOrderProvider _sut { get; set; }
        private readonly string _basePath = "C:\\Users\\Zachary Londono\\source\\repos\\RoyalExcelLibrary\\tests\\ExcelLibrary.Tests\\Unit\\TestData\\";

        [SetUp]
        public void Setup() {
            _sut = new HafeleDBOrderProvider();
        }

        [Test]
        [TestCase("HafeleTest1.xlsx", "23700076", "79.97", 3, "450 Huyler Street", "Suite 107", "South Hackensack", "NJ", "07606")]
        [TestCase("HafeleTest2.xlsx", "23700304", "326.08", 11, "1301 Tech Court", "", "Westminster", "Maryland", "21157")]
        [TestCase("HafeleTest3.xlsx", "23698752", "44.76", 1, "45 Saw Mill River Road", "", "Yonkers", "NY", "10701")]
        [TestCase("HafeleTest4.xlsx", "23698874", "439.57", 12, "11 Atlantic Avenue", "", "S. Yarmouth", "MA", "02664")]
        [TestCase("HafeleTest5.xlsx", "23670019", "492.29", 6, "3 Woodland Avenue", "", "Westhampton Beach", "NY", "11978")]
        [TestCase("HafeleTest6.xlsx", "123456", "154.35", 2, "123 Address Street", "", "Bound Brook", "NJ", "08876")]
        [TestCase("HafeleTest7.xlsx", "123456", "154.35", 2, "123 Address Street", "", "Bound Brook", "NJ", "08876")] // Metric test
        public void Should_LoadOrder_WhenFileIsValidOrder(string filePath,
                                                            string expectedNumber,
                                                            decimal expectedSubTotal,
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
            order.SubTotal.Should().Match(s => (s - expectedSubTotal <= 0.05M));
            order.Tax.Should().Be(0M);
            order.ShippingCost.Should().Be(0M);
            order.Products.Sum(p => p.Qty).Should().Be(expectedProdCount);

            order.Customer.Address.Should().BeEquivalentTo(new {
                Line1 = expectedAddressLine1,
                Line2 = expectedAddressLine2,
                City = expectedCity,
                State = expectedState,
                Zip = expectedZip
            });

        }

        [Test]
        public void Should_ConvertToInches_WhenUnitsSetToInches() {

            // Inches
            _sut.FilePath = _basePath + "HafeleTest6.xlsx";
            Order imperialOrder = _sut.LoadCurrentOrder();

            // MM
            _sut.FilePath = _basePath + "HafeleTest7.xlsx";
            Order metricOrder = _sut.LoadCurrentOrder();

            var im_products = imperialOrder.Products.ToList();
            var mm_products = metricOrder.Products.ToList();

            im_products.Count().Should().Be(mm_products.Count());

            for (int i = 0; i < im_products.Count(); i++) {
                var a = im_products[i] as DrawerBox;
                var b = mm_products[i] as DrawerBox;

                a.Qty.Should().Be(b.Qty);
                a.UnitPrice.Should().Be(b.UnitPrice);
                a.Height.Should().Match(h => (h - b.Height < 1));
                a.Width.Should().Be(b.Width);
                a.Depth.Should().Be(b.Depth);

            }

        }

    }

}
