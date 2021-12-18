using System;

using RoyalExcelLibrary.ExcelUI.Models;

namespace RoyalExcelLibrary.ExcelUI.Providers {

	public interface IOrderProvider {

		// <summary>
		// Loads the current job from the data source
		// </summar>
		Order LoadCurrentOrder();

	}

}
