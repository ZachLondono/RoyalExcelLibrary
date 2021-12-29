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
    public class AvailableMaterialQuery : IRequest<IEnumerable<Material>> { }

    internal class QueryHandler : IRequestHandler<AvailableMaterialQuery, IEnumerable<Material>> {

        private readonly DatabaseConfiguration _dbConfig;
        private readonly ILogger<QueryHandler> _logger;

        public QueryHandler(DatabaseConfiguration dbConfig, ILogger<QueryHandler> logger) {
            _dbConfig = dbConfig;
            _logger = logger;
        }

        public Task<IEnumerable<Material>> Handle(AvailableMaterialQuery request, CancellationToken cancellationToken) {

            _logger.LogInformation("Handling query for Material Configuration");

            string query = @"SELECT MaterialType.[Id] As TypeId, MaterialType.[MaterialName], MaterialType.[CutListCode], MaterialInventory.[Id], MaterialInventory.[Dimension], MaterialInventory.[Price]
                            FROM MaterialInventory
                            LEFT JOIN MaterialType
                            ON MaterialType.Id = MaterialInventory.TypeId
                            Where MaterialInventory.Profile = @Profile;";

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

                _logger.LogInformation("Materials returned by query {@Materials}", materials);

                return Task.FromResult(materials);

            }

        }

    }

}
