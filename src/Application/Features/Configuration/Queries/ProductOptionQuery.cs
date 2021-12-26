using RoyalExcelLibrary.Application.Common;
using Dapper;
using MediatR;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace RoyalExcelLibrary.Application.Features.Configuration {

    public class ProductOptionQuery : IRequest<ProductOptions> {
        public string Profile { get; set; }

        public ProductOptionQuery(string profile) {
            Profile = profile;
        }
    }

    internal class ProductOptionQueryHandler : IRequestHandler<ProductOptionQuery, ProductOptions> {

        private readonly DatabaseConfiguration _dbConfig;
        private readonly ILogger<QueryHandler> _logger;

        public ProductOptionQueryHandler(DatabaseConfiguration dbConfig, ILogger<QueryHandler> logger) {
            _dbConfig = dbConfig;
            _logger = logger;
        }

        public Task<ProductOptions> Handle(ProductOptionQuery request, CancellationToken cancellationToken) {

            _logger.LogInformation("Handling query for product option pricing");

            string productOptionQuery = @"SELECT [OptionCategory], [OptionName], [Price]
                                            FROM OptionPricing
                                            WHERE Profile = @Profile;";

            ProductOptions options = null;
            using (var connection = new OleDbConnection(_dbConfig.AppConfigConnectionString)) {

                connection.Open();

                var productOptions = connection.Query<(string category, string name, decimal price)>(productOptionQuery, request);
                Dictionary<string, Dictionary<string, decimal>> categories = new Dictionary<string, Dictionary<string, decimal>>();
                foreach (var (category, name, price) in productOptions) {

                    Dictionary<string, decimal> prices;
                    if (!categories.ContainsKey(category)) {
                        categories[category] = new Dictionary<string, decimal>();
                    }
                    prices = categories[category];

                    prices.Add(name, price);
                    _logger.LogInformation($"Product option price configuration loaded: [{category}] - [{name}] - [{price}]");
                }

                options = new ProductOptions(categories);

            }

            return Task.FromResult(options);
        }
    }

}
