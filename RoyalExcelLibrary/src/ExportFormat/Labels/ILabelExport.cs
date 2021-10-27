using RoyalExcelLibrary.Models;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.ExportFormat.Labels {

	public interface ILabelExport {

		void PrintLables(Order order);

	}

}
