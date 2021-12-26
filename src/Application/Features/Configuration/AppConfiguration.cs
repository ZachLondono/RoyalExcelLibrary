using RoyalExcelLibrary.Application.Features.Configuration.Export;
using System.Collections.Generic;
using System.IO;

namespace RoyalExcelLibrary.Application.Features.Configuration {

    public class AppConfiguration {

        private readonly IReadOnlyDictionary<string, string> _configs;

        public AppConfiguration(IReadOnlyDictionary<string, string> configs) {
            _configs = configs;
        }

        public string this[string key] {
            get {
                return _configs[key];
            }
        }

    }

    public class ExportOptions {

        private IReadOnlyDictionary<string, Configuration> _exportConfigs;

        public ExportOptions(IReadOnlyDictionary<string, Configuration> exportConfigs) {
            _exportConfigs = exportConfigs;
        }
        public Configuration this[string key] {
            get {
                return _exportConfigs[key];
            }
        }
        public class Configuration {

            public int ID { get; set; }

            public string TemplateName { get; set; }

            public string TemplatePath { get; set; }

            public int Copies { get; set; }

        }

    }

    public class ProductOptions {

        private readonly Dictionary<string, Dictionary<string, decimal>> _productOptions;

        public ProductOptions(Dictionary<string, Dictionary<string, decimal>> productOptions) {
            _productOptions = productOptions;
        }

        public decimal this[string categoryName, string optionName] {
            get {
                if (_productOptions is null) throw new InvalidDataException("No product option data loaded");
                return _productOptions[categoryName][optionName];
            }
        }

        public bool ContainsCategory(string categoryName) {
            return _productOptions.ContainsKey(categoryName);
        }

        public bool ContainsOption(string categoryName, string optionName) { 
            if (!_productOptions.ContainsKey(categoryName)) return false;
            return _productOptions[categoryName].ContainsKey(optionName);
        }

    }

}
