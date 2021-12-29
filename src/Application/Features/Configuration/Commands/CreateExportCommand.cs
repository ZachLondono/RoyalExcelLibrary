﻿using MediatR;
using System.Threading.Tasks;
using System.Threading;
using RoyalExcelLibrary.Application.Common;
using System.Data.OleDb;
using Microsoft.Extensions.Logging;

namespace RoyalExcelLibrary.Application.Features.Configuration.Export {

    public class CreateExportCommand : IRequest<ExportOptions.Configuration> { 
    
        public string TemplateName { get; set; }
        public string TemplatePath { get; set; }
        public int Copies { get; set; }

        public CreateExportCommand(string templateName, string templatePath, int copies) {
            TemplateName = templateName;
            TemplatePath = templatePath;
            Copies = copies;
        }

    }

    internal class CreateCommandHandler : IRequestHandler<CreateExportCommand, ExportOptions.Configuration> {

        private readonly DatabaseConfiguration _dbConfig;
        private readonly ILogger<CreateCommandHandler> _logger;

        public CreateCommandHandler(DatabaseConfiguration dbConfig, ILogger<CreateCommandHandler> logger) {
            _dbConfig = dbConfig;
            _logger = logger;
        }

        public Task<ExportOptions.Configuration> Handle(CreateExportCommand request, CancellationToken cancellationToken) {

            ExportOptions.Configuration newConfig = new ExportOptions.Configuration {
                TemplateName = request.TemplateName,
                TemplatePath = request.TemplatePath,
                Copies = request.Copies
            };

            _logger.LogInformation("Storing new export configuration: {@ExportConfig}", newConfig);

            using (var connection = new OleDbConnection(_dbConfig.AppConfigConnectionString)) {

                connection.Open();

                var command = connection.CreateCommand();
                command.CommandText = @"INSERT INTO [ExportTemplates] ([TemplateName], [TemplatePath], [Copies])
                                        VALUES (@TemplateName, @TemplatePath, @Copies);";
                command.Parameters.Add(new OleDbParameter("@TemplateName", OleDbType.VarChar)).Value = newConfig.TemplateName;
                command.Parameters.Add(new OleDbParameter("@TemplatePath", OleDbType.VarChar)).Value = newConfig.TemplatePath;
                command.Parameters.Add(new OleDbParameter("@Copies", OleDbType.Integer)).Value = newConfig.Copies;

                _logger.LogInformation("Executing Insert query");
                int rows = command.ExecuteNonQuery();

                if (rows > 0) {
                    var query = connection.CreateCommand();
                    query.CommandText = "SELECT @@IDENTITY FROM ExportTemplates;";
                    var reader = query.ExecuteReader();
                    reader.Read();
                    newConfig.ID = reader.GetInt32(0);
                    _logger.LogInformation("New export config ID: {@ID}", newConfig.ID);
                } else newConfig.ID = -1; // The new tempalte was not inserted

            }

            return Task.FromResult(newConfig);

        }

    }

}

