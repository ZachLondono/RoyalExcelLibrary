using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.DAL.Repositories {
	public interface IRepository<T> where T : BaseRepoClass {

		T GetById(int id);

		T Insert(T entity);

		void Delete(T entity);

		void Update(T entity);

		IEnumerable<T> GetAll();

	}

	public class BaseRepoClass {

		public int Id { get; set; }

	}

}
