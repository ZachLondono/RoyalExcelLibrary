using ClosedXML.Excel;
using NUnit.Framework;
using RoyalExcelLibrary.ExcelUI.src.FluentWorkbookValidation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLibrary.Tests.Unit {
    internal class ValidationTests {

        private readonly string path = @"C:\Users\Zachary Londono\source\repos\RoyalExcelLibraryV2\tests\ExcelLibrary.Tests\Unit\TestData\HafeleTest1.xlsx";

        [Test]
        public void Should_Be_Valid() {

            using (XLWorkbook workbook = new XLWorkbook(path)) {

                var validator = new WkbkValidator(workbook);

                validator.WkbkRule()
                    .HasSheet("Order Sheet");

                Assert.DoesNotThrow(() => validator.Validate());

            }

        }

        [Test]
        public void Should_Be_Valid_Worksheet() {

            using (XLWorkbook workbook = new XLWorkbook(path)) {

                var validator = new WkbkValidator(workbook);

                validator.WkbkRule()
                    .ForSheet("Order Sheet")
                    .HasRange("A1");

                Assert.DoesNotThrow(() => validator.Validate());

            }

        }

        [Test]
        public void Should_Be_Valid_Range() {

            using (XLWorkbook workbook = new XLWorkbook(path)) {

                var validator = new WkbkValidator(workbook);

                validator.WkbkRule()
                    .ForSheet("Order Sheet")
                    .ForRange("A15")
                    .NotEmpty();

                Assert.DoesNotThrow(() => validator.Validate());

            }

        }

        [Test]
        public void Should_Be_Valid_Range_MergedCell() {

            using (XLWorkbook workbook = new XLWorkbook(path)) {

                var validator = new WkbkValidator(workbook);

                validator.WkbkRule()
                    .ForSheet("Order Sheet")
                    .ForRange("N11")
                    .NotEmpty();

                Assert.DoesNotThrow(() => validator.Validate());

            }

        }

        [Test]
        public void Should_Be_Valid_Range_DoubleValue() {

            using (XLWorkbook workbook = new XLWorkbook(path)) {

                var validator = new WkbkValidator(workbook);

                validator.WkbkRule()
                    .ForSheet("Order Sheet")
                    .ForRange("A16")
                    .NotEmpty()
                    .ContainsDouble();

                Assert.DoesNotThrow(() => validator.Validate());

            }

        }

        [Test]
        public void Should_Be_Valid_TwoValidators() {

            using (XLWorkbook workbook = new XLWorkbook(path)) {

                var validator = new WkbkValidator(workbook);

                validator.WkbkRule()
                    .HasSheet("Order Sheet");

                validator.WkbkRule()
                    .HasSheet("Slides");

                Assert.DoesNotThrow(() => validator.Validate());

            }

        }

        [Test]
        public void Should_Be_Valid_ChainedValidation() {

            using (XLWorkbook workbook = new XLWorkbook(path)) {

                var validator = new WkbkValidator(workbook);

                validator.WkbkRule()
                    .HasSheet("Order Sheet")
                    .HasSheet("Slides");

                Assert.DoesNotThrow(() => validator.Validate());

            }

        }

        [Test]
        public void Should_Be_Valid_ChainedTypes() {

            using (XLWorkbook workbook = new XLWorkbook(path)) {

                var validator = new WkbkValidator(workbook);

                validator.WkbkRule()
                    .HasSheet("Order Sheet")
                    .ForSheet("Order Sheet")
                    .HasRange("F16")
                    .ForRange("F16")
                    .NotEmpty()
                    .ContainsDouble();

                Assert.DoesNotThrow(() => validator.Validate());

            }

        }

        [Test]
        public void Should_Throw_Exception() {

            using (XLWorkbook workbook = new XLWorkbook(path)) {

                var validator = new WkbkValidator(workbook);

                validator.WkbkRule()
                    .HasSheet("Does Not Exist");

                Assert.Throws(typeof(Exception), () => validator.Validate());
            }

        }

        [Test]
        public void Should_Throw_Exception_Worksheet() {

            using (XLWorkbook workbook = new XLWorkbook(path)) {

                var validator = new WkbkValidator(workbook);

                validator.WkbkRule()
                    .ForSheet("Order Sheet")
                    .HasRange("Does Not Exist");

                Assert.Throws(typeof(Exception), () => validator.Validate());

            }

        }

        [Test]
        public void Should_Throw_Exception_Range() {

            using (XLWorkbook workbook = new XLWorkbook(path)) {

                var validator = new WkbkValidator(workbook);

                validator.WkbkRule()
                    .ForSheet("Order Sheet")
                    .ForRange("A1")
                    .NotEmpty();

                Assert.Throws(typeof(Exception), () => validator.Validate());

            }

        }

        [Test]
        public void Should_Throw_Exception_Worksheet_WithMessage() {

            using (XLWorkbook workbook = new XLWorkbook(path)) {

                var validator = new WkbkValidator(workbook);

                string message = "Test Message";

                validator.WkbkRule()
                    .ForSheet("Order Sheet")
                    .HasRange("Does Not Exist")
                    .WithMessage(message);

                var exception = Assert.Throws(typeof(Exception), () => validator.Validate());

                Assert.AreEqual(message, exception.Message);

            }

        }

        [Test]
        public void Should_Throw_Exception_ChainedValidators() {

            using (XLWorkbook workbook = new XLWorkbook(path)) {

                var validator = new WkbkValidator(workbook);

                validator.WkbkRule()
                    .HasSheet("Order Sheet")
                    .HasSheet("Does Not Exist");

                Assert.Throws(typeof(Exception), () => validator.Validate());
            }

        }

        [Test]
        public void Should_Throw_Exception_TwoValidators() {

            using (XLWorkbook workbook = new XLWorkbook(path)) {

                var validator = new WkbkValidator(workbook);

                validator.WkbkRule()
                    .HasSheet("Order Sheet");

                validator.WkbkRule()
                    .HasSheet("Does Not Exist");

                Assert.Throws(typeof(Exception), () => validator.Validate());

            }

        }

        [Test]
        public void Should_Throw_Exception_With_Message() {

            using (XLWorkbook workbook = new XLWorkbook(path)) {

                var validator = new WkbkValidator(workbook);

                string message = "Test Message";

                validator.WkbkRule()
                    .HasSheet("Does Not Exist")
                    .WithMessage(message);

                var exception = Assert.Throws(typeof(Exception), () => validator.Validate());

                Assert.AreEqual(message, exception.Message);

            }

        }

        [Test]
        public void Should_Throw_Exception_ChainedTypes() {

            using (XLWorkbook workbook = new XLWorkbook(path)) {

                var validator = new WkbkValidator(workbook);
                
                string message = "Test Message";

                validator.WkbkRule()
                    .HasSheet("Order Sheet")
                    .ForSheet("Order Sheet")
                    .HasRange("A1")
                    .ForRange("A1")
                    .NotEmpty()
                    .ContainsDouble()
                    .WithMessage(message);

                var exception = Assert.Throws(typeof(Exception), () => validator.Validate());

                Assert.AreEqual(message, exception.Message);

            }

        }

    }
}
