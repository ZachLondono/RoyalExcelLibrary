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

            string exportQuery = @"SELECT [TemplateName], [TemplatePath], [Copies]
                            FROM ExportTemplates
                            WHERE [Profile] = @Profile;";

            string configMapQuery = @"SELECT [ID], [Key], [Value]
                                    FROM ConfigurationMap
                                    WHERE [Profile] = @Profile";

            AppConfiguration config = null;
            using (var connection = new OleDbConnection(_dbConfig.AppConfigConnectionString)) {

                connection.Open();

                var exportSettings = connection.Query<ExportConfiguration>(exportQuery, request);

                Dictionary<string, ExportConfiguration> exportConfigs = new Dictionary<string, ExportConfiguration>();
                foreach (var configSetting in exportSettings) {
                    if (!exportConfigs.ContainsKey(configSetting.TemplateName)) {
                        exportConfigs.Add(configSetting.TemplateName, configSetting);
                        _logger.LogInformation($"Export template loaded: {configSetting.TemplateName}");
                    } else _logger.LogWarning($"Duplicate export key found in profile '{configSetting.TemplateName}'");
                }

                var configSettings = connection.Query<(string key,string value)>(configMapQuery, request);
                Dictionary<string, string> configMap = new Dictionary<string, string>();
                foreach (var configSetting in configSettings) {
                    if (!configMap.ContainsKey(configSetting.key)) {
                        configMap.Add(configSetting.key, configSetting.value);
                        _logger.LogInformation($"Configuration loaded: {configSetting.key} -> {configSetting.value}");
                    } else _logger.LogWarning($"Duplicate configuration key found in profile '{configSetting.key}'");
                }

                config = new AppConfiguration(exportConfigs, configMap);

            }

            return Task.FromResult(config);

        }

    }

}
