using Dapper;
using MediatR;
using RoyalExcelLibrary.Application.Common;
using RoyalExcelLibrary.Application.Features.Options.Materials;
using System.Data.OleDb;
using System.Threading;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Application.Features.Product.Query {
    public class MaterialTypeQuery : IRequest<MaterialType> {
        public int Id { get; set; }
        public MaterialTypeQuery(int id) {
            Id = id;
        }
    }

    internal class MaterialTypeQueryHandler : IRequestHandler<MaterialTypeQuery, MaterialType> {

        private readonly DatabaseConfiguration _dbConfig;

        public MaterialTypeQueryHandler(DatabaseConfiguration dbConfig) {
            _dbConfig = dbConfig;
        }
        public Task<MaterialType> Handle(MaterialTypeQuery request, CancellationToken cancellationToken) {

            MaterialType matType = null;
            using (var connection = new OleDbConnection(_dbConfig.JobConnectionString)) {

                connection.Open();

                matType = connection.QueryFirstOrDefault<MaterialType>(
                    sql: @"SELECT Id AS TypeId, MaterialName, CutListCode
                            FROM MaterialType
                            WHERE MaterialType.[Id] = @Id;",
                    param: request);

                connection.Close();

            }

            return Task.FromResult(matType);

        }
    }

}
