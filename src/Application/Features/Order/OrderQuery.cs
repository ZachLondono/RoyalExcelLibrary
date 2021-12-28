using Dapper;
using MediatR;
using RoyalExcelLibrary.Application.Common;
using RoyalExcelLibrary.Application.Features.Product;
using RoyalExcelLibrary.Application.Features.Product.Query;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Application.Features.Order {

    public class OrderQuery :IRequest<Order> {
        public int Id { get; set; }
        public OrderQuery(int id) {
            Id = id;
        }
    }

    internal class OrderQueryHandler : IRequestHandler<OrderQuery, Order>{
        
        private readonly DatabaseConfiguration _dbConfig;
        private readonly ISender _sender;

        public OrderQueryHandler(DatabaseConfiguration dbConfig, ISender sender) {
            _dbConfig = dbConfig;
            _sender = sender;
        }

        public Task<Order> Handle(OrderQuery request, CancellationToken cancellationToken) {

            Order order = null;
            Dictionary<string, string> details = null;
            IEnumerable<int> productIds = null;
            using (var connection = new OleDbConnection(_dbConfig.JobConnectionString)) {
                connection.Open();

                order = connection.QueryFirstOrDefault<Order>(
                                sql: @"SELECT [Customer], [OrderName], [OrderNumber], [OrderDate]
                                    FROM [Orders]
                                    WHERE [ID] = @Id;",
                                param: request);

                details = connection.Query<(string key, string val)>(
                                sql: @"SELECT [Key], [DetailValue]
                                    FROM [OrderDetails]
                                    WHERE [OrderId] = @Id;",
                                param: request)
                                .ToDictionary(t => t.key, t => t.val);

                productIds = connection.Query<int>(
                                sql: @"SELECT [ID] FROM [DrawerBoxes] WHERE [OrderId] = @Id;",
                                param: request);


                order.OrderDetails = details;

                connection.Close();
            }

            if (!(productIds is null)) {
                // Add all products to the order
                List<IProduct> products = new List<IProduct>();
                foreach (int id in productIds) {
                    DrawerBox box = _sender.Send(new DrawerBoxQuery(id)).Result;
                    products.Add(box);
                }

                order.Products = products;
            }

            return Task.FromResult(order);
        }
    }

}
