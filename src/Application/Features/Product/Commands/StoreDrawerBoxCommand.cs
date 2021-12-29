using MediatR;
using Dapper;
using RoyalExcelLibrary.Application.Common;
using System;
using System.Data.OleDb;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace RoyalExcelLibrary.Application.Features.Product.Commands {
    public class StoreDrawerBoxCommand : IRequest<DrawerBox> {
        public DrawerBox Box { get; set; }
        public int OrderId { get; set; }
        public StoreDrawerBoxCommand(DrawerBox box, int orderId) {
            Box = box;
            OrderId = orderId;
        }
    }

    internal class StoreDrawerBoxCommandHandler : IRequestHandler<StoreDrawerBoxCommand, DrawerBox> {

        private readonly DatabaseConfiguration _dbConfig;
        private readonly ILogger<StoreDrawerBoxCommandHandler> _logger;

        public StoreDrawerBoxCommandHandler(DatabaseConfiguration dbConfig, ILogger<StoreDrawerBoxCommandHandler> logger) {
            _dbConfig = dbConfig;
            _logger = logger;
        }

        public Task<DrawerBox> Handle(StoreDrawerBoxCommand request, CancellationToken cancellationToken) {

            _logger.LogInformation("Storing drawerbox {@DrawerBox}", request.Box);

            using (var connection = new OleDbConnection(_dbConfig.JobConnectionString)) {

                connection.Open();

                // Insert the drawerbox 
                int rows = connection.Execute(
                    sql: @"INSERT INTO [DrawerBoxes] ([Qty], [Height], [Width], [Depth], [BoxMaterial], [BottomMaterial], [OrderId])
                            VALUES (@Qty, @Height, @Width, @Depth, @BoxMaterial, @BottomMaterial, @OrderId);",
                    param: new {
                        request.Box.Qty,
                        request.Box.Height,
                        request.Box.Width,
                        request.Box.Depth,
                        request.OrderId,
                        BoxMaterial = request.Box.BoxMaterial.TypeId,
                        BottomMaterial = request.Box.BottomMaterial.TypeId
                    });


                // Check that the drawerbox was inserted, and get it's Id
                if (rows > 0) {
                    var query = connection.CreateCommand();
                    query.CommandText = "SELECT @@IDENTITY FROM DrawerBoxes;";
                    var reader = query.ExecuteReader();
                    reader.Read();
                    request.Box.Id = reader.GetInt32(0);
                    _logger.LogInformation("New drawerbox stored with ID: {@ID}", request.Box.Id);
                } else {
                    request.Box.Id = -1; // The new drawerbox was not inserted
                    return Task.FromResult(request.Box);
                }

                // Insert the drawerbox's extra options
                foreach (var extra in request.Box.Extras) {
                    string category = extra.Key;
                    string option = extra.Value;

                    try {

                        connection.Execute(
                            sql: @"INSERT INTO [DrawerBoxExtras] ([Category], [Option], [ProductId])
                                VALUES (@Category, @Option, @ProductId);",
                            param: new {
                                Category = category,
                                Option = option,
                                ProductId = request.Box.Id
                            });
                        
                        _logger.LogInformation("Drawer box extra stored: {@Category} -> {@Option}", category, option);

                    } catch (Exception e) {
                        _logger.LogError("Error storing drawerbox extra: {@Category} -> {@Option}\n{@Exception}", category, option, e);
                    }
                }

                connection.Close();

            }

            return Task.FromResult(request.Box);

        }
    }

}
