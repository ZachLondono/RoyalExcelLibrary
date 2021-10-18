using System;

using RoyalExcelLibrary.Models;

namespace RoyalExcelLibrary.Providers {

	public interface IOrderProvider {

		// <summary>
		// Loads the current job from the data source
		// </summar>
		Order LoadCurrentOrder();

	}

}
