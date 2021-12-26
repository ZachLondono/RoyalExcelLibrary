using System.Collections.Generic;

namespace RoyalExcelLibrary.Application.Features.Product {
    public interface IProduct {
        int Qty { get; }
        string Name { get; }
        //decimal Price();
    }

    public interface ICompositeProduct<T> : IProduct where T : IProduct {
        IList<T> GetParts();
    }

}
