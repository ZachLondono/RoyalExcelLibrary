using RoyalExcelLibrary.DAL.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Models {
	
	// <summary>
	// Represents an item in the inventory which can be used to build an product
	// </summary>
	public class InventoryItem : BaseRepoClass {

		public string Name { get; set; }

		public double Width { get; set; }

		public double Length { get; set; }
		
		public double Thickness { get; set; }

		public bool IsAvailable { get; set; }

		public MaterialType Material { get; set; }

	}
}
