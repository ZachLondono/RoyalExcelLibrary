using Microsoft.Extensions.DependencyInjection;
using System.Reflection;
using MediatR;
using RoyalExcelLibrary.Application.Common;
using Microsoft.Extensions.Configuration;

namespace RoyalExcelLibrary.Application {
    public static class DependencyInjection {
    
        public static IServiceCollection AddApplication(this IServiceCollection services, DatabaseConfiguration configuration) {
            
            services.AddMediatR(Assembly.GetAssembly(typeof(DependencyInjection)));

            // Connection strings to the configuration and job database is read at startup
            services.AddSingleton(configuration);

            return services;

        }
    
    }

}
