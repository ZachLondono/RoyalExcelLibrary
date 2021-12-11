using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Services {
    public interface ILabelServiceFactory {

        ILabelService CreateService(string template);

    }

    public class DymoLabelServiceFactory : ILabelServiceFactory {
        public ILabelService CreateService(string template) {
            return new DymoLabelService(template);
        }

    }

}
