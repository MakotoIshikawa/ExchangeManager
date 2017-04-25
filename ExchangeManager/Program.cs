using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExchangeManager.Extensions;
using ExchangeManager.Interface;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager {
	public class Program {
		private static string _username = @"ishikawm@kariverification14.onmicrosoft.com";
		private static string _password = @"Ishikawam!";

		public static void Main(string[] args) {
			var service = new ExchangeOnlineManager(_username, _password);
		}
	}
}
