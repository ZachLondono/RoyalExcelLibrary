using System;
using Microsoft.Extensions.Hosting;
using RoyalExcelLibrary.Application;
using Microsoft.Extensions.DependencyInjection;
using MediatR;
using System.Threading.Tasks;
using RoyalExcelLibrary.Application.Features.Configuration.Export;
using RoyalExcelLibrary.Application.Features.Configuration;
using System.Diagnostics;
using ExcelDna.Integration;
using Microsoft.Extensions.Logging;
using RoyalExcelLibrary.Application.Common;
using System.Collections.Generic;
using RoyalExcelLibrary.Application.Features.Options.Materials;
using RoyalExcelLibrary.Application.Features.Product;
using RoyalExcelLibrary.Application.Features.Product.Commands;
using System.Windows.Forms;
using RoyalExcelLibrary.Application.Features.Product.Query;
using RoyalExcelLibrary.Application.Features.Order;

namespace RoyalExcelLibrary.ExcelUI.src {
    public class RoyalAddIn : IExcelAddIn {

        private IHost _host;
        private static ILogger<RoyalAddIn> _logger;
        private static ISender _sender;
        
        public static AppConfiguration Configuration { get; private set; }

        public void AutoOpen() {

            Debug.WriteLine("Opening RoyalAddIn");

            var settings = System.Configuration.ConfigurationManager.ConnectionStrings;
            var dbConfig = new DatabaseConfiguration {
                AppConfigConnectionString = settings["AppConfigConnectionString"].ConnectionString ?? "",
                JobConnectionString = settings["JobConnectionString"].ConnectionString ?? ""
            };

            _host = Host.CreateDefaultBuilder()
                .ConfigureServices((context, services) => {
                    services.AddApplication(dbConfig); // Load the RoyalExcelLibrary.Application class library
                    services.AddLogging();
                })
                .Build();

            _logger = _host.Services.GetService<ILogger<RoyalAddIn>>();

            _sender = _host.Services.GetService<ISender>();

            LoadAppConfiguration("A");

        }

        public void AutoClose() { }

        public void LoadAppConfiguration(string profile) {
            
            try {

                Task<AppConfiguration> configTask = _sender.Send(new AppConfigurationQuery(profile));

                Configuration = configTask.Result;

            } catch (Exception e) {
                _logger.LogError("Error reading configuration:\n" + e.ToString());
            }

        }

        public static IEnumerable<Material> GetMaterials() {

            try {

                Task<IEnumerable<Material>> materialTask = _sender.Send(new AvailableMaterialQuery());

                return materialTask.Result;

            } catch (Exception e) {
                _logger.LogError("Error reading Material:\n" + e.ToString());
            }

            return null;

        }

        public static ExportOptions.Configuration CreateExportTemplate(string name, string path, int copies) {
            
            try {

                Task<ExportOptions.Configuration> export = _sender.Send(new CreateExportCommand(name, path, copies));
                return export.Result;

            } catch (Exception e) {
                _logger.LogError("Error creating Export Template:\n" + e.ToString());
            }

            return null;

        }

        public static Order StoreOrder(Order order) {

            try {

                Task<Order> task = _sender.Send(new StoreOrderCommand(order));
                return task.Result;

            } catch (Exception e) {
                _logger.LogError("Error storing order:\n" + e.ToString());
                MessageBox.Show(e.ToString(), "Exception");
            }

            return null;

        }

        public static Order QueryOrder(int orderId) {

            try {

                Task<Order> task = _sender.Send(new OrderQuery(orderId));
                return task.Result;

            } catch (Exception e) {
                _logger.LogError("Error reading order:\n" + e.ToString());
                MessageBox.Show(e.ToString(), "Exception");
            }

            return null;

        }

        public static DrawerBox StoreDrawerBox(DrawerBox drawerBox, int jobId) {

            try {

                Task<DrawerBox> task = _sender.Send(new StoreDrawerBoxCommand(drawerBox, jobId));
                return task.Result;

            } catch (Exception e) {
                _logger.LogError("Error storing drawerbox:\n" + e.ToString());
                MessageBox.Show(e.ToString(), "Exception");
            }

            return null;

        }

        public static DrawerBox QueryDrawerBox(int boxId) {

            try {

                Task<DrawerBox> task = _sender.Send(new DrawerBoxQuery(boxId));
                return task.Result;

            } catch (Exception e) {
                _logger.LogError("Error querying drawerbox:\n" + e.ToString());
                MessageBox.Show(e.ToString(), "Exception");
            }

            return null;

        }

    }

}

