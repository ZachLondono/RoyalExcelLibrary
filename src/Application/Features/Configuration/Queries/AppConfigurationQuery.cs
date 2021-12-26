using RoyalExcelLibrary.Application.Common;
using Dapper;
using MediatR;
using System.Collections.Generic;
using System.Data.OleDb;
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

            string configMapQuery = @"SELECT [ID], [Key], [Value]
                                    FROM ConfigurationMap
                                    WHERE [Profile] = @Profile";

            AppConfiguration config = null;
            using (var connection = new OleDbConnection(_dbConfig.AppConfigConnectionString)) {

                connection.Open();

                var configSettings = connection.Query<(string key,string value)>(configMapQuery, request);
                Dictionary<string, string> configMap = new Dictionary<string, string>();
                foreach (var (key, value) in configSettings) {
                    if (!configMap.ContainsKey(key)) {
                        configMap.Add(key, value);
                        _logger.LogInformation($"Configuration loaded: {key} -> {value}");
                    } else _logger.LogWarning($"Duplicate configuration key found in profile '{key}'");
                }

                config = new AppConfiguration(configMap);

            }

            return Task.FromResult(config);

        }

    }

}
