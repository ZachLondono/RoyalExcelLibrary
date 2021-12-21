using System.Collections.Generic;

namespace RoyalExcelLibrary.Application.Features.Product {

    internal class ClosetPanel : IProduct {

        public int Qty { get; private set; }

        public string Name { get; private set; }

        public Dictionary<string, string> Parameters { get; private set; }

        public ClosetPanel(string name, Dictionary<string,string> parameters) {
            Name = name;
            Parameters = parameters;
        }

    }

}