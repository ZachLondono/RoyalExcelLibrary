using FluentAssertions;
using NUnit.Framework;

namespace ExcelLibrary.Tests.Unit {
    public class HelperFuncTest {
        [Test]
        [TestCase(38.1, "1 1/2")]
        [TestCase(25.4, "1")]
        [TestCase(26.9875, "1 1/16")]
        [TestCase(12.7, "1/2")]
        public void Should_ConvertMetricDoubleToInchesString(double metric, string expected) {
            string result = RoyalExcelLibrary.ExcelUI.HelperFuncs.FractionalImperialDim(metric);
            result.Should().Be(expected);
        }

        [Test]
        [TestCase(25.00, "31/32")]     // 63/64
        [TestCase(25.40, "1")]         // 1
        [TestCase(25.80, "1 1/32")]    // 1 1/64
        [TestCase(26.19, "1 1/32")]    // 1 1/32
        [TestCase(26.59, "1 1/32")]    // 1 3/64
        [TestCase(26.99, "1 1/16")]    // 1 1/16
        [TestCase(27.38, "1 1/16")]    // 1 5/64
        [TestCase(27.78, "1 3/32")]    // 1 3/32
        [TestCase(28.18, "1 1/8")]     // 1 7/64
        [TestCase(28.58, "1 1/8")]     // 1 1/8
        [TestCase(28.97, "1 1/8")]     // 1 9/64
        [TestCase(29.37, "1 5/32")]    // 1 5/32
        [TestCase(29.77, "1 3/16")]    // 1 11/64
        public void Should_RoundToNearest16nd_WhenMetricIsNotExactlyA16nd(double metric, string expected) {
            string result = RoyalExcelLibrary.ExcelUI.HelperFuncs.FractionalImperialDim(metric);
            result.Should().Be(expected);
        }

        [Test]
        [TestCase("1 1/2", 1.5)]
        [TestCase("5 1/32", 5.03125)]
        [TestCase("0", 0)]
        [TestCase("1", 1)]
        [TestCase("-1", -1)]
        public void Should_ConvertStringToDouble_WhenValidString(string value, double expected) {
            double result = RoyalExcelLibrary.ExcelUI.HelperFuncs.ConvertToDouble(value);
            result.Should().Be(expected);
        }

    }

}
