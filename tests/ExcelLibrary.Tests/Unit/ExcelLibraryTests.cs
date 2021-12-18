using FluentAssertions;
using NUnit.Framework;
using RoyalExcelLibrary.ExcelUI.Providers;
using System;
using System.Collections.Generic;

namespace ExcelLibrary.Tests.Unit {

    public class ExcelLibraryTests {

        [Test]
        [TestCase("319.04", "9.72")]
        [TestCase("1969.32", "58.40")]
        [TestCase("536.18", "16.12")]
        [TestCase("320.9", "9.76")]
        public void Should_CalculateStripeFee(decimal totalCharge, decimal expectedFee) {
            decimal result = RoyalExcelLibrary.ExcelUI.ExcelLibrary.CalculateStripeFee(totalCharge);
            Assert.AreEqual(expectedFee, result);
        }

        [Test]
        [TestCase("allmoxy", typeof(AllmoxyOrderProvider))]
        [TestCase("ot", typeof(OTDBOrderProvider))]
        [TestCase("hafele", typeof(HafeleDBOrderProvider))]
        [TestCase("richelieu", typeof(RichelieuExcelDBOrderProvider))]
        [TestCase("loaded", typeof(UniversalDBOrderProvider))]
        [TestCase("aLlMoxY", typeof(AllmoxyOrderProvider))]
        [TestCase("oT", typeof(OTDBOrderProvider))]
        [TestCase("hAfELe", typeof(HafeleDBOrderProvider))]
        [TestCase("riChElIeU", typeof(RichelieuExcelDBOrderProvider))]
        [TestCase("LoAdeD", typeof(UniversalDBOrderProvider))]
        public void Should_ReturnIOrderProvider_FromValidProviderName(string providerName, Type expectedType) {
            IOrderProvider orderProvider = RoyalExcelLibrary.ExcelUI.ExcelLibrary.GetProviderByName(providerName);
            orderProvider.Should().Match(p => p.GetType() == expectedType);
        }

        [Test]
        public void Should_ThrowException_WhenProviderNameInvalid() {
            Action result = () => RoyalExcelLibrary.ExcelUI.ExcelLibrary.GetProviderByName("DoesNotExist"); 
            result.Should().Throw<ArgumentException>();
        }

    }

}
