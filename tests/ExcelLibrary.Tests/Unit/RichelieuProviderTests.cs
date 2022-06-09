using FluentAssertions;
using NUnit.Framework;
using RoyalExcelLibrary.ExcelUI.Models;
using RoyalExcelLibrary.ExcelUI.Models.Options;
using RoyalExcelLibrary.ExcelUI.Models.Products;
using RoyalExcelLibrary.ExcelUI.Providers;
using System.Linq;
using static RoyalExcelLibrary.ExcelUI.Providers.RichelieuExcelDBOrderProvider;

namespace ExcelLibrary.Tests.Unit {
    internal class RichelieuProviderTests {

        private RichelieuExcelDBOrderProvider _sut { get; set; }
        private readonly string _basePath = @"C:\Users\Zachary Londono\source\repos\RoyalExcelLibraryV2\tests\ExcelLibrary.Tests\Unit\TestData\";

        [SetUp]
        public void Setup() {
            _sut = new RichelieuExcelDBOrderProvider();
        }

        [Test]
        // Different notch tests
        [TestCase("RCT08114INNNHRR3", UndermountNotch.No_Notch, "BirchFJ", "Plywood 1/4\"", false, false, true)]
        [TestCase("RCT08114IWHNHRR3", UndermountNotch.Wide_Notch, "BirchFJ", "Plywood 1/4\"", false, false, true)]
        [TestCase("RCT08114IFBNHRR3", UndermountNotch.Front_Back, "BirchFJ", "Plywood 1/4\"", false, false, true)]
        // Different scoop/front tests
        [TestCase("RCT08114IFBNHNR3", UndermountNotch.Front_Back, "BirchFJ", "Plywood 1/4\"", false, true, true)]
        [TestCase("RCT08114IFBNH1R3", UndermountNotch.Front_Back, "BirchFJ", "Plywood 1/4\"", true, true, true)]
        [TestCase("RCT08114IFBNH2R3", UndermountNotch.Front_Back, "BirchFJ", "Plywood 1/4\"", true, true, true)]
        [TestCase("RCT08114IFBNH3R3", UndermountNotch.Front_Back, "BirchFJ", "Plywood 1/4\"", true, true, true)]
        [TestCase("RCT08114IFBNH4R3", UndermountNotch.Front_Back, "BirchFJ", "Plywood 1/4\"", true, true, true)]
        // Rush vs No Rush Tests
        [TestCase("RCT08114IFBNH4R3", UndermountNotch.Front_Back, "BirchFJ", "Plywood 1/4\"", true, true, true)]
        [TestCase("RCT08114IFBNH4R0", UndermountNotch.Front_Back, "BirchFJ", "Plywood 1/4\"", true, true, false)]
        // Economy vs Hybrid Tests
        [TestCase("RCT09114IFBNH1R3", UndermountNotch.Front_Back, "Hybrid", "Plywood 1/4\"", true, true, true)]
        [TestCase("RCT08114IFBNH1R3", UndermountNotch.Front_Back, "BirchFJ", "Plywood 1/4\"", true, true, true)]
        public void Should_ParseSkuToDrawerBox(string sku, UndermountNotch expectedNotch, string expectedSideMaterial, string expectedBottomMaterial, bool scoopFront, bool clearFront, bool rush) {
            var config = _sut.ParseSku(sku);
            config.Notch.Should().Be(expectedNotch);
            config.BoxMaterial.Should().Be(expectedSideMaterial);
            config.BotMaterial.Should().Be(expectedBottomMaterial);
            config.ScoopFront.Should().Be(scoopFront);
            config.PullOutFront.Should().Be(clearFront);
            config.Rush.Should().Be(rush);
        }

        [Test]
        [TestCase("RichelieuTest1.xml", "EA5045A", "ORDER-830 CORLEIA", "99.30", "0", "0", 2, "612 U.S. ROUTE 9", "", "WEST CREEK", "New Jersey", "08092")]
        [TestCase("RichelieuTest2.xml", "EA5979A", "ORDER-832 HORVAT", "57.93", "0", "0", 3, "612 U.S. ROUTE 9", "", "WEST CREEK", "New Jersey", "08092")]
        [TestCase("RichelieuTest3.xml", "EA6004A", "ORDER-834 HORVAT", "59.01", "0", "0", 3, "612 U.S. ROUTE 9", "", "WEST CREEK", "New Jersey", "08092")]
        [TestCase("RichelieuTest4.xml", "EA6255A", "ORDER-836 HORVAT", "272.18", "0", "0", 9, "612 U.S. ROUTE 9", "", "WEST CREEK", "New Jersey", "08092")] 
        [TestCase("RichelieuTest5.xml", "EA6335A", "J-23889 Order", "178.05", "0", "0", 6, "50 SCHOOLHOUSE RD", "", "SOUDERTON", "Pennsylvania", "18964")]
        [TestCase("RichelieuTest6.xml", "EA9980A", "Dorsey", "726.86", "0", "0", 22, "1909 E Westmoreland Street", "", "Philadelphia", "Pennsylvania", "19134")]
        [TestCase("RichelieuTest7.xml", "EB7989A", "dawn pantry", "281.30", "0", "0", 4, "42 frost circle", "", "middletown", "New Jersey", "07748")]
        [TestCase("RichelieuTest8.xml", "D89989A", "ORDER-6769", "179.07", "0", "0", 3, "3001 IRWIN DR STE C", "", "MOUNT LAUREL", "New Jersey", "08054")]
        [TestCase("RichelieuTest9.xml", "EC4232A", "KAAS", "768.11", "0", "0", 24, "802 MACOPIN ROAD", "", "WEST MILFORD", "New Jersey", "07480")]
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

            order.Products.Count(p => (p as DrawerBox).NotchOption == UndermountNotch.Unknown).Should().Be(0);

            // Check that the materials where properly read
            //order.Products.Count(p => (p as DrawerBox).SideMaterial == MaterialType.Unknown).Should().Be(0);
            //order.Products.Count(p => (p as DrawerBox).BottomMaterial == MaterialType.Unknown).Should().Be(0);

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
