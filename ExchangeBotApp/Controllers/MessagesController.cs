using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using ExchangeBotApp.Dialogs;
using ExchangeBotApp.Extensions;
using Microsoft.Bot.Connector;

namespace ExchangeBotApp {
	[BotAuthentication]
	public class MessagesController : ApiController {
		/// <summary>
		/// POST: api/Messages
		/// ユーザーからのメッセージを受信して返信する
		/// </summary>
		public async Task<HttpResponseMessage> Post([FromBody]Activity activity)
			=> await activity.PostAsync<RootDialog>(a => {
				switch (a?.Type) {
				case ActivityTypes.DeleteUserData:
					// ここでユーザーの削除を実装する
					// ユーザーの削除を処理する場合は、実際のメッセージを返す
					break;
				case ActivityTypes.ConversationUpdate:
					// メンバーの追加や削除など、会話の状態の変更を処理する
					// 情報には[Activity.MembersAdded]、[Activity.MembersRemoved]、[Activity.Action]を使用します
					// すべてのチャンネルで利用できるわけではありません
					break;
				case ActivityTypes.ContactRelationUpdate:
					// 連絡先リストからの追加/削除の処理
					// [Activity.From] + [Activity.Action]は何が起こったかを表します
					break;
				case ActivityTypes.Typing:
					// ユーザーが入力していることを知っているハンドル
					break;
				case ActivityTypes.Ping:
					break;
				}
			});
	}
}