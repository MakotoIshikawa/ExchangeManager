using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager {
	class Program {
		static void Main(string[] args) {
			var url = @"https://kariverification14.sharepoint.com";
			var username = @"ishikawm@kariverification14.onmicrosoft.com";
			var password = @"Ishikawam!";

			var service = new ExchangeService(ExchangeVersion.Exchange2013_SP1) {
				Credentials = new WebCredentials(username, password),
				UseDefaultCredentials = true,
				TraceEnabled = true,
				TraceFlags = TraceFlags.All,
			};

			service.AutodiscoverUrl(username, redirectionUrl => {
				// 検証コールバックのデフォルトは、URLを拒否することです。
				var result = false;

				var redirectionUri = new Uri(redirectionUrl);

				// リダイレクトURLの内容を検証します。
				// この単純な検証コールバックでは、HTTPSを使用して認証資格情報を暗号化する場合、
				// リダイレクトURLは有効と見なされます。
				if (redirectionUri.Scheme == "https") {
					result = true;
				}

				return result;
			});

			var subject = "HelloWorld";
			var text = "これは、EWS Managed APIを使用して送信した最初のメールです。";
			var isRichText = false;

			var email = new EmailMessage(service) {
				Subject = subject,
				Body = new MessageBody(isRichText ? BodyType.HTML : BodyType.Text, text),
			};
			email.ToRecipients.Add(username);

			email.Send();
		}
	}
}
