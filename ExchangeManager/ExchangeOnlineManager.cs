using System;
using ExchangeManager.Interface;
using ExchangeManager.Primitives;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager {
	/// <summary>
	/// Exchange Online の情報を管理するクラスです。
	/// </summary>
	public class ExchangeOnlineManager : ExchangeManagerBase, IExchangeManager {
		#region フィールド

		#endregion

		#region コンストラクタ

		/// <summary>
		/// コンストラクタ
		/// </summary>
		/// <param name="username">ユーザー名 (メールアドレス)</param>
		/// <param name="password">パスワード</param>
		public ExchangeOnlineManager(string username, string password) : base(username) {
			this.Password = password;
			this.Service = this.CreateService();

			this.Service.AutodiscoverUrl(username, url => {
				// 検証コールバックのデフォルトは、URLを拒否することです。
				var result = false;

				var redirectionUri = new Uri(url);

				// リダイレクトURLの内容を検証します。
				// この単純な検証コールバックでは、HTTPSを使用して認証資格情報を暗号化する場合、
				// リダイレクトURLは有効と見なされます。
				if (redirectionUri.Scheme == "https") {
					result = true;
				}

				return result;
			});
		}

		#endregion

		#region プロパティ

		/// <summary>
		/// パスワード
		/// </summary>
		public string Password { get; protected set; }

		#endregion

		#region メソッド

		/// <summary>
		/// EWS の新しいインスタンスを生成します。
		/// </summary>
		/// <returns>生成した EWS のインスタンスを返します。</returns>
		protected override Ews.ExchangeService CreateService()
			=> new Ews.ExchangeService(Ews.ExchangeVersion.Exchange2013_SP1) {
				Credentials = new Ews.WebCredentials(this.UserName, this.Password),
				UseDefaultCredentials = false,
				TraceEnabled = true,
				TraceFlags = Ews.TraceFlags.All,
			};

		#endregion
	}
}
