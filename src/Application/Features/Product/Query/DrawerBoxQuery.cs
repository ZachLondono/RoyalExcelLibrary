using MediatR;
using Dapper;
using RoyalExcelLibrary.Application.Common;
using System.Linq;
using System.Data.OleDb;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using RoyalExcelLibrary.Application.Features.Options.Materials;
using Microsoft.Extensions.Logging;

namespace RoyalExcelLibrary.Application.Features.Product.Query {
    public class DrawerBoxQuery : IRequest<DrawerBox> {
        public int Id { get; set; }
        public DrawerBoxQuery(int id) {
            Id = id;
        }
    }

    internal class DrawerBoxQueryHandler : IRequestHandler<DrawerBoxQuery, DrawerBox> {

        private readonly DatabaseConfiguration _dbConfig;
        private readonly ISender _sender;
        private readonly ILogger<DrawerBoxQueryHandler> _logger;

        public DrawerBoxQueryHandler(DatabaseConfiguration dbConfig, ISender sender, ILogger<DrawerBoxQueryHandler> logger) {
            _dbConfig = dbConfig;
            _sender = sender;
            _logger = logger;
        }

        public Task<DrawerBox> Handle(DrawerBoxQuery request, CancellationToken cancellationToken) {

            _logger.LogInformation("Querying drawer box with ID: {@ID}", request.Id);

            Dictionary<string, string> extras;
            DBDTO db;
            using (var connection = new OleDbConnection(_dbConfig.JobConnectionString)) {
                connection.Open();

                // Get all the extras for the drawer box
                extras = connection.Query<(string category, string option)>(
                                sql:@"SELECT [Category], [Option]
                                    FROM [DrawerBoxExtras]
                                    WHERE [ProductId] = @Id;",
                                param:request)
                                .ToDictionary(t => t.category, t => t.option);

                _logger.LogInformation("Drawerbox extras returned: {@Extras}", extras);

                // Get all the basic drawerbox properties
                db = connection.QueryFirstOrDefault<DBDTO>(
                                sql:@"SELECT [Qty], [Height], [Width], [Depth], [BoxMaterial], [BottomMaterial], [OrderId]
                                    FROM[DrawerBoxes]
                                    WHERE[ID] = @Id;",
                                param:request);

                _logger.LogInformation("Drawerbox data returned: {@DrawerBox}", db);

                connection.Close();
            }

            DrawerBox box = null;

            if (db is null) return Task.FromResult(box);

            DrawerBoxBuilder builder = new DrawerBoxBuilder()
                    .WithQty(db.Qty)
                    .WithHeight(db.Height)
                    .WithWidth(db.Width)
                    .WithDepth(db.Depth);

            _logger.LogInformation("Querying drawerbox material box:{@BoxMaterial} bot:{@BotMaterial}", db.BoxMaterial, db.BottomMaterial);
            // Get the bottom and box materials
            MaterialType bottomMaterial = _sender.Send(new MaterialTypeQuery(db.BottomMaterial)).Result;
            MaterialType boxMaterial = _sender.Send(new MaterialTypeQuery(db.BoxMaterial)).Result;
            _logger.LogInformation("Queried material box:{@BoxMaterial} bot:{@BotMaterial}", boxMaterial, bottomMaterial);

            builder.WithBotMaterial(bottomMaterial);
            builder.WithBoxMaterial(boxMaterial);

            foreach (var extra in extras) {
                builder.WithExtra(extra.Key, extra.Value);
            }

            return Task.FromResult(builder.Build());

        }

        internal class DBDTO {
            public int Qty { get; set; }
            public int Height { get; set; }
            public int Width { get; set; }
            public int Depth { get; set; }
            public int BoxMaterial { get; set; }
            public int BottomMaterial { get; set; }
            public int OrderId { get; set; }
        }

    }
}
