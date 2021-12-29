using MediatR;
using Dapper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using RoyalExcelLibrary.Application.Common;
using System.Data.OleDb;
using RoyalExcelLibrary.Application.Features.Product;
using RoyalExcelLibrary.Application.Features.Product.Commands;
using Microsoft.Extensions.Logging;

namespace RoyalExcelLibrary.Application.Features.Order {
    public class StoreOrderCommand : IRequest<Order> {
        public Order Order { get; set; }
        public StoreOrderCommand(Order order) {
            Order = order;
        }
    }

    internal class StoreOrderCommandHandler : IRequestHandler<StoreOrderCommand, Order> {
        
        private readonly DatabaseConfiguration _dbConfig;
        private readonly ISender _sender;
        private readonly ILogger<StoreOrderCommandHandler> _logger;

        public StoreOrderCommandHandler(DatabaseConfiguration dbConfig, ISender sender, ILogger<StoreOrderCommandHandler> logger) {
            _dbConfig = dbConfig;
            _sender = sender;
            _logger = logger;
        }

        public Task<Order> Handle(StoreOrderCommand request, CancellationToken cancellationToken) {

            _logger.LogInformation("Storing new order: {@Order}", request.Order);

            using (var connection = new OleDbConnection(_dbConfig.JobConnectionString)) {

                connection.Open();

                // Insert the order
                var command = connection.CreateCommand();
                command.CommandText = @"INSERT INTO [Orders] ([Customer], [OrderName], [OrderNumber], [OrderDate])
                                        VALUES (@Customer, @OrderName, @OrderNumber, @OrderDate);";

                command.Parameters.Add("@Customer", OleDbType.VarChar).Value = request.Order.Customer;
                command.Parameters.Add("@OrderName", OleDbType.VarChar).Value = request.Order.OrderName;
                command.Parameters.Add("@OrderNumber", OleDbType.VarChar).Value = request.Order.OrderNumber;
                command.Parameters.Add("@OrderDate", OleDbType.Date).Value = request.Order.OrderDate;

                int rows = command.ExecuteNonQuery();

                // Check that the orders was inserted, and get it's Id
                if (rows > 0) {
                    var query = connection.CreateCommand();
                    query.CommandText = "SELECT @@IDENTITY FROM DrawerBoxes;";
                    var reader = query.ExecuteReader();
                    reader.Read();
                    request.Order.Id = reader.GetInt32(0);
                    _logger.LogInformation("New order stored with ID: {@ID}", request.Order.Id);
                } else {
                    request.Order.Id = -1; // The new drawerbox was not inserted
                    return Task.FromResult(request.Order);
                }

                // Insert the orders's extra detail
                foreach (var extra in request.Order.OrderDetails) {
                    string key = extra.Key;
                    string detailValue = extra.Value;

                    try {
                        connection.Execute(
                            sql: @"INSERT INTO [OrderDetails] ([Key], [DetailValue], [OrderId])
                                VALUES (@Key, @DetailValue, @OrderId);",
                            param: new {
                                Key = key,
                                DetailValue = detailValue,
                                OrderId = request.Order.Id
                            });

                        _logger.LogInformation("Stored order detail: {@Key} -> {@Value}", key, detailValue);
                    } catch (Exception ex) {
                        _logger.LogError("Failed to store order detail: {@Key} -> {@Value}\n{@Exception}", key, detailValue, ex);
                    }
                }

                //TODO: insert invoice information

                connection.Close();

            }

            _logger.LogInformation("Storing products in order");
            foreach (IProduct product in request.Order.Products)
                if (product is DrawerBox)
                    product.Id = _sender.Send(new StoreDrawerBoxCommand(product as DrawerBox, request.Order.Id)).Result.Id;

            return Task.FromResult(request.Order);

        }
    }
}
