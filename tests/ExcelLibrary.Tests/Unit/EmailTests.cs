using FluentAssertions;
using NUnit.Framework;
using RoyalExcelLibrary.ExcelUI.ExportFormat;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLibrary.Tests.Unit {
    public class EmailTests {


        [Test]
        public void Should_InsertIntoHtmlBody() {

            string result = OutlookEmailExport.InsertIntoExistingBody("<html><body someattrb='abc'><p>Existing Content</p></body></html>", "Hello World");

            result.Should().Be("<html><body someattrb='abc'><span>Hello World</span><p>Existing Content</p></body></html>");

        }

    }
}
