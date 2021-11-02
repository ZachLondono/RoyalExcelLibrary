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

		WhiteOak,

		Plywood1_2,

		Plywood1_4,
		
		WhiteMela1_2,
		
		WhiteMela1_4,
		
		BlackMela1_2,

		BlackMela1_4

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
				case MaterialType.WhiteMela1_2:
					return "whitemela_1_2";
				case MaterialType.WhiteMela1_4:
					return "whitemela_1_4";
				case MaterialType.BlackMela1_2:
					return "blackmela_1_2";
				case MaterialType.BlackMela1_4:
					return "blackmela_1_4";
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
				case "whitemela_1_2":
					return MaterialType.WhiteMela1_2;
				case "whitemela_1_4":
					return MaterialType.WhiteMela1_4;
				case "blackmela_1_2":
					return MaterialType.BlackMela1_2;
				case "blackmela_1_4":
					return MaterialType.BlackMela1_4;
				default:
					return MaterialType.Unknown;
			}
		}


	}

}
