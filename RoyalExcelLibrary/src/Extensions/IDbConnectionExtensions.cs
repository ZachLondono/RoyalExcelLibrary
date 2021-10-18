using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;

namespace RoyalExcelLibrary.Extensions {
	public static class IDbConnectionExtensions {
	
		public static void AddParamWithValue(this IDbCommand command, string name, object value) {
			var param = command.CreateParameter();
			param.ParameterName = name;
			param.Value = value;
			command.Parameters.Add(param);
		}		

	}

}
