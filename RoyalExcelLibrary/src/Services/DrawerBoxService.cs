using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.DAL;
using RoyalExcelLibrary.Providers;

using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Hosting;
using System.Threading.Tasks;
using System.Threading;
using RoyalExcelLibrary.DAL.Repositories;
using System.Data;
using RoyalExcelLibrary.Models.Products;

namespace RoyalExcelLibrary.Services {
    public class DrawerBoxService : IProductService {

        private readonly IJobRepository _jobRepository;
        private readonly IDrawerBoxRepository _drawerBoxRepository;
        private readonly IDbConnection _connection;

        public DrawerBoxService(IDbConnection dbConnection) {
            _connection = dbConnection;
            _jobRepository = new JobRepository(dbConnection);
            _drawerBoxRepository = new DrawerBoxRepository(dbConnection);

		}

        // <summary>
        // Stores the job in the current excel workbook in the job database and tracks the material it requires
        // </summar>
		public void StoreCurrentOrder(Order order) {

            Job job = _jobRepository.Insert(order.Job);
            order.Job.Id = job.Id;


            int count = 0;
            foreach (IProduct product in order.Products) {
                if (product is DrawerBox) {
                    DrawerBox drawerBox = (DrawerBox)product;
                    drawerBox.JobId = order.Job.Id;
                    _drawerBoxRepository.Insert(drawerBox);
                    count++;
                } 
            }

        }

		public void GenerateConfirmation() {
			throw new System.NotImplementedException();
		}

		public void GenerateInvoice() {
			throw new System.NotImplementedException();
		}

		public void ConfirmOrder() {
			throw new System.NotImplementedException();
		}

		public void PayOrder() {
			throw new System.NotImplementedException();
		}

	}

}
