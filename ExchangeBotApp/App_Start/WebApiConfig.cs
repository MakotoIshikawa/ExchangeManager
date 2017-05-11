using System.Web.Http;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace ExchangeBotApp {
	public static class WebApiConfig {
		public static void Register(HttpConfiguration config) {
			// Json settings
			var with = config.Formatters.JsonFormatter.SerializerSettings;
			with.NullValueHandling = NullValueHandling.Ignore;
			with.ContractResolver = new CamelCasePropertyNamesContractResolver();
			with.Formatting = Formatting.Indented;

			JsonConvert.DefaultSettings = () => new JsonSerializerSettings {
				ContractResolver = new CamelCasePropertyNamesContractResolver(),
				Formatting = Formatting.Indented,
				NullValueHandling = NullValueHandling.Ignore,
			};

			// Web API configuration and services

			// Web API routes
			config.MapHttpAttributeRoutes();

			config.Routes.MapHttpRoute(
				name: "DefaultApi",
				routeTemplate: "api/{controller}/{id}",
				defaults: new { id = RouteParameter.Optional }
			);
		}
	}
}
