using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExchangeManager.Extensions;
using ExchangeManager.Interface;
using ExchangeManager.Primitives;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager {
	public class ExchangeOnlineManager : ExchangeManagerBase, IExchangeManager {
		#region フィールド

		#endregion

		#region コンストラクタ

		/// <summary>
		/// コンストラクタ
		/// </summary>
		/// <param name="username">ユーザー名</param>
		/// <param name="password">パスワード</param>
		public ExchangeOnlineManager(string username, string password) : base() {
			this.Service = new ExchangeService(ExchangeVersion.Exchange2013_SP1) {
				Credentials = new WebCredentials(username, password),
				UseDefaultCredentials = false,
				TraceEnabled = true,
				TraceFlags = TraceFlags.All,
			};

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

		#endregion

		#region メソッド

		#endregion
	}
}
