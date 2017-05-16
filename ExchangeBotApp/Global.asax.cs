using System.Web;
using System.Web.Http;

namespace ExchangeBotApp {
	/// <summary>
	/// Web アプリケーション
	/// </summary>
	public class WebApiApplication : HttpApplication {
		/// <summary>
		/// アプリケーション開始
		/// </summary>
		protected void Application_Start()
			=> GlobalConfiguration.Configure(WebApiConfig.Register);
	}
}
