using System.Collections.Generic;

namespace RoyalExcelLibrary.Application.Features.Product {

    internal class ClosetPanel : IProduct {

        public int Id { get; set; }
        public int Qty { get; private set; }

        public string Name { get; private set; }

        public IReadOnlyDictionary<string, string> Parameters { get; private set; }

        public ClosetPanel(int qty, string name, Dictionary<string,string> parameters) {
            Qty = qty;
            Name = name;
            Parameters = parameters;
        }

    }

}