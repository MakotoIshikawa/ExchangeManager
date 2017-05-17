using System;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;

namespace BotLibrary.Extensions {
	/// <summary>
	/// Activity を拡張するメソッドを提供します。
	/// </summary>
	public static partial class ActivityExtension {
		#region SendAsync

		/// <summary>
		/// 非同期でダイアログの処理を開始します。
		/// </summary>
		/// <typeparam name="TDialog">Dialog</typeparam>
		/// <param name="this">Activity</param>
		public static async Task SendAsync<TDialog>(this Activity @this) where TDialog : IDialog<object>, new()
			=> await Conversation.SendAsync(@this, () => new TDialog());

		#endregion

		#region PostAsync

		/// <summary>
		/// 非同期で POST します。
		/// </summary>
		/// <typeparam name="TDialog">ダイアログの型</typeparam>
		/// <param name="activity">Activity</param>
		/// <param name="handleSystemMessage">システムメッセージを処理するメソッド</param>
		/// <returns>HTTP 応答メッセージを返します。</returns>
		public static async Task<HttpResponseMessage> PostAsync<TDialog>(this Activity activity, Action<Activity> handleSystemMessage = null) where TDialog : IDialog<object>, new() {
			var type = activity?.GetActivityType();
			if (type == ActivityTypes.Message) {
				await activity.SendAsync<TDialog>();
			} else {
				handleSystemMessage?.Invoke(activity);
			}

			var response = new HttpResponseMessage(HttpStatusCode.Accepted);
			return response;
		}

		#endregion
	}
}