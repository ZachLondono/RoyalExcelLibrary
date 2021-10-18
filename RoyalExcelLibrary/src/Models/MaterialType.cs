using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Models {
	public enum MaterialType {

		Unknown,

		SolidBirch,

		EconomyBirch,

		HybridBirch,

		SolidWalnut,

		Plywood1_2,

		Plywood1_4

	}

	public class MaterialFunctions {

		public static string TypeToString(MaterialType matType) {
			switch (matType) {
				case MaterialType.SolidBirch:
					return "solid_birch";
				case MaterialType.EconomyBirch:
					return "economy_birch";
				case MaterialType.HybridBirch:
					return "hybrid_birch";
				case MaterialType.Plywood1_2:
					return "plywood_1_2";
				case MaterialType.Plywood1_4:
					return "plywood_1_4";
				default:
					return "Unknown";
			}
		}

		public static MaterialType StringToType(string matType) {
			switch (matType) {
				case "solid_birch":
					return MaterialType.SolidBirch;
				case "economy_birch":
					return MaterialType.EconomyBirch;
				case "hybrid_birch":
					return MaterialType.HybridBirch;
				case "plywood_1_2":
					return MaterialType.Plywood1_2;
				case "plywood_1_4":
					return MaterialType.Plywood1_4;
				default:
					return MaterialType.Unknown;
			}
		}


	}

}
