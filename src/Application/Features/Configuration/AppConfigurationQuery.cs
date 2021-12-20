using RoyalExcelLibrary.Application.Common;
using RoyalExcelLibrary.Application.Features.Configuration.Export;
using Dapper;
using MediatR;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace RoyalExcelLibrary.Application.Features.Configuration {
    
    /// <summary>
    /// Queries the Application Configuration Database for all of the application settings
    /// </summary>
    public class AppConfigurationQuery : IRequest<AppConfiguration> {   
        public string Profile { get; set; }

        public AppConfigurationQuery(string profile) {
            Profile = profile;
        }

    }

    internal class QueryHandler : IRequestHandler<AppConfigurationQuery, AppConfiguration> {

        private readonly DatabaseConfiguration _dbConfig;
        private readonly ILogger<QueryHandler> _logger;

        public QueryHandler(DatabaseConfiguration dbConfig, ILogger<QueryHandler> logger) {
            _dbConfig = dbConfig;
            _logger = logger;
        }

        public Task<AppConfiguration> Handle(AppConfigurationQuery request, CancellationToken cancellationToken) {

            _logger.LogInformation("Handling query for Application Configuration");

            string query = @"SELECT [ID], [TemplateName], [TemplatePath], [Copies]
                            FROM ExportTemplates
                            WHERE [Profile] = @Profile;";

            AppConfiguration config = null;
            using (var connection = new OleDbConnection(_dbConfig.AppConfigConnectionString)) {

                connection.Open();

                var configSettings = connection.Query<ExportConfiguration>(query, request);

                Dictionary<string, ExportConfiguration> configMap = new Dictionary<string, ExportConfiguration>();
                foreach (var configSetting in configSettings) {
                    if (!configMap.ContainsKey(configSetting.TemplateName)) {
                        configMap.Add(configSetting.TemplateName, configSetting);
                        _logger.LogInformation($"Export tempalte loaded: {configSetting.TemplateName}");
                    }
                }

                 config = new AppConfiguration(configMap);

            }

            return Task.FromResult(config);

        }

    }

}
