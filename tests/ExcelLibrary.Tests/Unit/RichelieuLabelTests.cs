using FluentAssertions;
using NUnit.Framework;
using RoyalExcelLibrary.ExcelUI.ExportFormat;
using RoyalExcelLibrary.ExcelUI.ExportFormat.Labels;
using RoyalExcelLibrary.ExcelUI.Models;
using RoyalExcelLibrary.ExcelUI.Models.Products;
using RoyalExcelLibrary.ExcelUI.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLibrary.Tests.Unit {
    internal class RichelieuLabelTests {

        [Test]
        public void Should_CreateLabels_WhenOrderIsValid() {

            // Arrange
            Job job = new Job {
                CreationDate = DateTime.Now,
                GrossRevenue = 0,
                JobSource = "Test",
                Name = "ABC"
            };

            RichelieuOrder order = new RichelieuOrder(job);
            order.Number = "OrderNumber";
            order.RichelieuNumber = "123";
            order.ClientFirstName = "FirstName";
            order.ClientLastName = "LastName";
            order.Customer = new Company {
                Name = "CustomerName",
                Address = new Address {
                    Line1 = "1",
                    Line2 = "2",
                    City = "3",
                    State = "4",
                    Zip = "5"
                }
            };
            order.AddProduct(new DrawerBox {
                Height = 3 * 25.4,
                Width = 2 * 25.4,
                Depth = 25.4,
                Qty = 1,
                ProductDescription = "Drawer Box Description",
                LineNumber = 1,
                Note = "Drawer Box Note"
            });
            RichelieuLabelExport sut = new RichelieuLabelExport();

            // Act / Assert
            sut.PrintLables(order, new MockLabelFactory());


        }

        class MockLabelFactory : ILabelServiceFactory {
            public ILabelService CreateService(string template) {
                return MockRichelieuLabelService.GetInstance();
            }
        }

        class MockRichelieuLabelService : ILabelService {

            public List<Tuple<Label, int>> labels = new List<Tuple<Label, int>>();

            private static MockRichelieuLabelService _instance = null;

            public static MockRichelieuLabelService GetInstance() {
                if (_instance is null) _instance = new MockRichelieuLabelService();
                return _instance;
            }

            public void AddLabel(Label label, int qty) {
                labels.Add(new Tuple<Label, int>(label, qty));
            }

            public Label CreateLabel() {

                Label label = new Label {
                    LabelFields = new Dictionary<string, LabelField> {
                        { "JOB", new LabelField { Type = LabelFieldType.Text, Value = null } },
                        { "PO", new LabelField { Type = LabelFieldType.Text, Value = null } },
                        { "SIZE", new LabelField { Type = LabelFieldType.Text, Value = null } },
                        { "QTY", new LabelField { Type = LabelFieldType.Text, Value = null } },
                        { "DESC", new LabelField { Type = LabelFieldType.Text, Value = null } },
                        { "ORDER", new LabelField { Type = LabelFieldType.Text, Value = null } },
                        { "NOTE", new LabelField { Type = LabelFieldType.Text, Value = null } },
                        { "TEXT", new LabelField { Type = LabelFieldType.Text, Value = null } },
                        { "TEXT_1", new LabelField { Type = LabelFieldType.Text, Value = null } },
                        { "TEXT_2", new LabelField { Type = LabelFieldType.Text, Value = null } },
                        { "ADDRESS", new LabelField { Type = LabelFieldType.Text, Value = null } },
                    }
                };

                return label;
            }


            public int _timesPrinted = 0;

            public void PrintLabels() {

                _timesPrinted++;

                if (_timesPrinted == 1) {
                    labels.Count().Should().Be(1);
                    Tuple<Label, int> boxLabelTuple = labels[0];
                    boxLabelTuple.Item2.Should().Be(1);
                    Label boxLabel = boxLabelTuple.Item1;
                    boxLabel.LabelFields["JOB"].Value.Should().Be("ABC");
                    boxLabel.LabelFields["PO"].Value.Should().Be("OrderNumber");
                    boxLabel.LabelFields["SIZE"].Value.Should().Be("3\"Hx2\"Wx1\"D");
                    boxLabel.LabelFields["QTY"].Value.Should().Be(1);
                    boxLabel.LabelFields["DESC"].Value.Should().Be("Drawer Box Description");
                    boxLabel.LabelFields["ORDER"].Value.Should().Be("123 : 1");
                    boxLabel.LabelFields["NOTE"].Value.Should().Be("Drawer Box Note");
                } else {
                    labels.Count().Should().Be(2);
                    Tuple<Label, int> shippingLabelTuple = labels[1];
                    shippingLabelTuple.Item2.Should().Be(1);
                }

            }

        }

    }
}
