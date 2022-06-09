using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelLibrary.Tests.Unit {
    
    internal class RegexTesting {

        [Test]
        [TestCase("Fixed Divider 1", "1")]
        [TestCase("fixed divider 1", "1")]
        [TestCase("Fixed Divider 2", "2")]
        [TestCase("Fixed Divider 3", "3")]
        [TestCase("Fixed Divider 10", "10")]
        [TestCase("Fixed Divider 11", "11")]
        public void TestRegex(string insertOption, string expected) {

            Regex rx = new Regex(@"(?<=Fixed\sDivider\s)[0-9]+", RegexOptions.IgnoreCase);

            MatchCollection matches = rx.Matches(insertOption);

            Assert.NotZero(matches.Count);
            Assert.AreEqual(expected, matches[0].Value);

        }

    }

}
