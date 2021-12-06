using FluentAssertions;
using NUnit.Framework;
using RoyalExcelLibrary.Providers;
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
            decimal result = RoyalExcelLibrary.ExcelLibrary.CalculateStripeFee(totalCharge);
            Assert.AreEqual(expectedFee, result);
        }

        [Test]
        [TestCase("319.04", "0", "9.72", "19.82", "0.13", "37.64")]
        [TestCase("536.18", "0", "16.12", "0", "0.13", "67.61")]
        [TestCase("1969.32", "50", "58.4", "122.36", "0.13", "226.01")]
        public void Should_CalculateOTCommission(decimal totalCharge, decimal shippingCost, decimal tax, decimal stripeFee, decimal commissionRate, decimal expectedComission) {
            decimal result = RoyalExcelLibrary.ExcelLibrary.CalculateCommissionPayment(totalCharge, shippingCost, tax, stripeFee, commissionRate);
            Assert.AreEqual(expectedComission, result);
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
            IOrderProvider orderProvider = RoyalExcelLibrary.ExcelLibrary.GetProviderByName(providerName);
            orderProvider.Should().Match(p => p.GetType() == expectedType);
        }

        [Test]
        public void Should_ThrowException_WhenProviderNameInvalid() {
            Action result = () => RoyalExcelLibrary.ExcelLibrary.GetProviderByName("DoesNotExist"); 
            result.Should().Throw<ArgumentException>();
        }

    }

}
