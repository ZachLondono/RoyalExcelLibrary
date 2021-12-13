using FluentAssertions;
using NUnit.Framework;
using RoyalExcelLibrary.Models.Products;
using RoyalExcelLibrary.Providers;
using System.Linq;

namespace ExcelLibrary.Tests.Unit {
    internal class RichelieuProviderTests {

        private RichelieuExcelDBOrderProvider _sut { get; set; }
        private readonly string _basePath = "C:\\Users\\Zachary Londono\\source\\repos\\RoyalExcelLibrary\\tests\\ExcelLibrary.Tests\\Unit\\TestData\\";

        [SetUp]
        public void Setup() {
            _sut = new RichelieuExcelDBOrderProvider();
        }


        [Test]
        [TestCase("RichelieuTest1.xml", "EA5045A", "ORDER-830 CORLEIA", "99.30", "0", "0", 2, "612 U.S. ROUTE 9", "", "WEST CREEK", "New Jersey", "08092")]
        [TestCase("RichelieuTest2.xml", "EA5979A", "ORDER-832 HORVAT", "57.93", "0", "0", 3, "612 U.S. ROUTE 9", "", "WEST CREEK", "New Jersey", "08092")]
        [TestCase("RichelieuTest3.xml", "EA6004A", "ORDER-834 HORVAT", "59.01", "0", "0", 3, "612 U.S. ROUTE 9", "", "WEST CREEK", "New Jersey", "08092")]
        [TestCase("RichelieuTest4.xml", "EA6255A", "ORDER-836 HORVAT", "272.18", "0", "0", 9, "612 U.S. ROUTE 9", "", "WEST CREEK", "New Jersey", "08092")] 
        [TestCase("RichelieuTest5.xml", "EA6335A", "J-23889 Order", "178.05", "0", "0", 6, "50 SCHOOLHOUSE RD", "", "SOUDERTON", "Pennsylvania", "18964")]
        [TestCase("RichelieuTest6.xml", "EA9980A", "Dorsey", "726.86", "0", "0", 22, "1909 E Westmoreland Street", "", "Philadelphia", "Pennsylvania", "19134")]
        [TestCase("RichelieuTest7.xml", "EB7989A", "dawn pantry", "281.30", "0", "0", 4, "42 frost circle", "", "middletown", "New Jersey", "07748")]
        public void Should_LoadOrder_WhenFileIsValidOrder(string filePath,
                                                            string expectedNumber,
                                                            string expectedJobName,
                                                            decimal expectedSubTotal,
                                                            decimal expectedTax,
                                                            decimal expectedShipping,
                                                            int expectedProdCount,
                                                            string expectedAddressLine1,
                                                            string expectedAddressLine2,
                                                            string expectedCity,
                                                            string expectedState,
                                                            string expectedZip) {
            // Arrange
            //Load xml from file
            _sut.XMLContent = System.IO.File.ReadAllText(_basePath + filePath);

            // Act
            var order = _sut.LoadCurrentOrder();

            // Assert
            order.Should().NotBeNull();
            order.Number.Should().Be(expectedNumber);
            order.Job.Name.Should().Be(expectedJobName);
            // Richelieu order total doesn't always match the sum of the individual items in the order, but should always be within 1 cent
            order.SubTotal.Should().Match(s => expectedSubTotal - s <= 0.03M);
            order.Tax.Should().Be(expectedTax);
            order.ShippingCost.Should().Be(expectedShipping);
            order.Products.Sum(p => p.Qty).Should().Be(expectedProdCount);
            order.Products.Count(p => (p as DrawerBox).SideMaterial == RoyalExcelLibrary.Models.MaterialType.Unknown).Should().Be(0);

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
