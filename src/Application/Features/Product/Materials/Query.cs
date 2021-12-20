using Dapper;
using MediatR;
using Microsoft.Extensions.Logging;
using RoyalExcelLibrary.Application.Common;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Application.Features.Options.Materials {
    public class MaterialQuery : IRequest<IEnumerable<Material>> { }

    internal class QueryHandler : IRequestHandler<MaterialQuery, IEnumerable<Material>> {

        private readonly DatabaseConfiguration _dbConfig;
        private readonly ILogger<QueryHandler> _logger;

        public QueryHandler(DatabaseConfiguration dbConfig, ILogger<QueryHandler> logger) {
            _dbConfig = dbConfig;
            _logger = logger;
        }

        public Task<IEnumerable<Material>> Handle(MaterialQuery request, CancellationToken cancellationToken) {

            _logger.LogInformation("Handling query for Material Configuration");

            string query = @"SELECT [MaterialID] As [TypeId], [MaterialName], [CutListCode], [PriceId] As [Id], [Dimension], [Price]
                            FROM Materials
                            LEFT JOIN MaterialPrices
                            ON MaterialPrices.MaterialKey = Materials.MaterialId;";

            using (var connection = new OleDbConnection(_dbConfig.AppConfigConnectionString)) {

                connection.Open();

                var materials = connection.Query<MaterialType, Material, Material>(
                    query,
                    (matType, mat) => {
                        mat.Type = matType;
                        return mat;
                    },
                    splitOn: "Id"
                );

                return Task.FromResult(materials);

            }

        }

    }

}
