using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Application.Features.Product {
    internal interface IProduct {

        int Qty { get; }

        string Name { get; }

        string Description { get; }

        //decimal Price();

    }

    internal interface ICompositeProduct<T> : IProduct where T : IProduct {

        IList<T> GetParts();

    }

}
