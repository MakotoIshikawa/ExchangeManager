using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExtensionsLibrary.Extensions;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Luis.Models;
using Microsoft.Bot.Connector;

namespace BotLibrary.Extensions {
	/// <summary>
	/// DialogContext を拡張するメソッドを提供します。
	/// </summary>
	public static partial class DialogContextExtension {
		#region メソッド

		#region ユーザーデータ

		/// <summary>
		/// ユーザーデータから値を取得します。
		/// </summary>
		/// <typeparam name="T">値の型</typeparam>
		/// <param name="this">DialogContext</param>
		/// <param name="key">キー</param>
		/// <returns>値を返します。</returns>
		public static T GetUserData<T>(this IDialogContext @this, string key)
			=> @this.UserData.GetValueOrDefault(u => u.GetValue<T>(key));

		/// <summary>
		/// ユーザーデータに値を設定します。
		/// </summary>
		/// <typeparam name="T">値の型</typeparam>
		/// <param name="this">DialogContext</param>
		/// <param name="key">キー</param>
		/// <param name="value">値</param>
		public static void SetUserData<T>(this IDialogContext @this, string key, T value)
			=> @this.UserData.SetValue(key, value);

		#endregion

		#region PostDefaultResponseMessageAsync

		/// <summary>
		/// デフォルト時の応答メッセージを POST します。
		/// </summary>
		/// <param name="this">IDialogContext</param>
		/// <param name="result">LUIS の結果</param>
		/// <param name="dictionary">ボタンの元となる連想配列</param>
		/// <param name="ex">例外</param>
		public static async Task PostDefaultResponseMessageAsync(this IDialogContext @this, LuisResult result, Dictionary<string, string> dictionary, Exception ex = null) {
			var msg = GetDefaultResponseMessage(result, ex);
			await @this.PostButtonsAsync(msg, dictionary);
		}

		/// <summary>
		/// デフォルト時の応答メッセージを取得します。
		/// </summary>
		/// <param name="result">LUIS の結果</param>
		/// <param name="ex">例外</param>
		/// <returns>応答メッセージを返します。</returns>
		private static string GetDefaultResponseMessage(this LuisResult result, Exception ex = null) {
			var sb = new StringBuilder();
			sb.AppendLine("お問い合わせありがとうございます。");

			if (ex != null) {
				sb.AppendLine($" {ex.Message} ");
			}

			sb.AppendLine("恐れ入りますが以下のサイトでお調べください。")
			.AppendLine($@">https://www.google.co.jp/search?q={result.Query}");

			return sb.ToString();
		}

		#endregion

		#region PostButtonsAsync

		/// <summary>
		/// Button 群を POST します。
		/// </summary>
		/// <param name="this">DialogContext</param>
		/// <param name="text">メッセージ</param>
		/// <param name="dictionary">ボタン群の元となる連想配列</param>
		public static async Task PostButtonsAsync(this IDialogContext @this, string text, Dictionary<string, string> dictionary) {
			var btns = dictionary?.ToButtons();
			await @this.PostMessageAsync(text, btns);
		}

		private static Attachment ToButtons(this Dictionary<string, string> dictionary, string text = null) {
			var cd = new ThumbnailCard {
				Buttons = (
					from kv in dictionary
					select kv.ToButton()
				).ToList(),
			};
			return cd.ToAttachment();
		}

		private static CardAction ToButton(this KeyValuePair<string, string> kv) {
			return new CardAction() {
				Value = kv.Value,
				Type = "postBack",
				Title = kv.Key,
			};
		}

		#endregion

		#region PostMessageAsync

		/// <summary>
		/// 回答を非同期で POST します。
		/// </summary>
		/// <param name="context">IDialogContext</param>
		/// <param name="text">テキスト文字列</param>
		/// <param name="attachments">付属品の配列</param>
		/// <returns></returns>
		public static async Task PostMessageAsync(this IDialogContext context, string text, params Attachment[] attachments) {
			var msg = context.CreateMessage();

			if (!text.IsEmpty()) {
				msg.Text = text;
			}

			if (attachments?.Any(a => a != null) ?? false) {
				attachments.ForEach(a => msg.Attachments.Add(a));
			}

			await context.PostAsync(msg);
		}

		private static IMessageActivity CreateMessage(this IDialogContext context) {
			var msg = context.MakeMessage();
			if (msg.Attachments == null) {
				msg.Attachments = new List<Attachment>();
			}

			return msg;
		}

		#endregion

		#endregion
	}
}