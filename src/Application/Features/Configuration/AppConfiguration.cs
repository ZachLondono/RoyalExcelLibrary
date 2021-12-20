using RoyalExcelLibrary.Application.Features.Configuration.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Application.Features.Configuration {

    public class AppConfiguration {

        public IReadOnlyDictionary<string, string> _configs { get; }

        public IReadOnlyDictionary<string, ExportConfiguration> ExportConfigs { get; }

        public AppConfiguration(IReadOnlyDictionary<string, ExportConfiguration> exportConfigs, IReadOnlyDictionary<string, string> configs) {
            ExportConfigs = exportConfigs;
            _configs = configs;
        }

        public string this[string key] {
            get {
                return _configs[key];
            }
        }

    }

}
