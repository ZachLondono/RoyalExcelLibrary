using RoyalExcelLibrary.Application.Common;
using Dapper;
using MediatR;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace RoyalExcelLibrary.Application.Features.Configuration {
    
    public class ExportOptionQuery : IRequest<ExportOptions> {
        public string Profile { get; set; }
        public ExportOptionQuery(string profile) {
            Profile = profile;
        }
    }

    internal class ExportOptionQueryHandler : IRequestHandler<ExportOptionQuery, ExportOptions> {

        private readonly DatabaseConfiguration _dbConfig;
        private readonly ILogger<QueryHandler> _logger;

        public ExportOptionQueryHandler(DatabaseConfiguration dbConfig, ILogger<QueryHandler> logger) {
            _dbConfig = dbConfig;
            _logger = logger;
        }

        public Task<ExportOptions> Handle(ExportOptionQuery request, CancellationToken cancellationToken) {
            _logger.LogInformation("Handling query for Export Options");

            string exportQuery = @"SELECT [TemplateName], [TemplatePath], [Copies]
                            FROM ExportTemplates
                            WHERE [Profile] = @Profile;";

            ExportOptions options = null;

            using (var connection = new OleDbConnection(_dbConfig.AppConfigConnectionString)) {

                connection.Open();

                var exportSettings = connection.Query<ExportOptions.Configuration>(exportQuery, request);
                Dictionary<string, ExportOptions.Configuration> exportConfigs = new Dictionary<string, ExportOptions.Configuration>();
                foreach (var configSetting in exportSettings) {
                    if (!exportConfigs.ContainsKey(configSetting.TemplateName)) {
                        exportConfigs.Add(configSetting.TemplateName, configSetting);
                        _logger.LogInformation($"Export template loaded: {configSetting.TemplateName}");
                    } else _logger.LogWarning($"Duplicate export key found in profile '{configSetting.TemplateName}'");
                }

                options = new ExportOptions(exportConfigs);
            }

            return Task.FromResult(options);

        }
    }


}
