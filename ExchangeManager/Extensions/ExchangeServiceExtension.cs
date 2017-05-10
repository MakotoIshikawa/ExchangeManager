using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExtensionsLibrary.Extensions;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager.Extensions {
	/// <summary>
	/// ExchangeService を拡張するメソッドを提供します。
	/// </summary>
	public static partial class ExchangeServiceExtension {
		#region メソッド

		#region メール送信

		/// <summary>
		/// メールを送信します。
		/// </summary>
		/// <param name="this">ExchangeService</param>
		/// <param name="to">宛先 (; 区切りで複数指定できます。)</param>
		/// <param name="subject">件名</param>
		/// <param name="body">本文</param>
		/// <param name="isRichText">リッチテキストかどうかを指定します。</param>
		/// <param name="setting">メールの設定をするメソッド</param>
		public static void SendMail(this Ews.ExchangeService @this, string to, string subject, string body, bool isRichText = false, Action<Ews.EmailMessage> setting = null) {
			var email = new Ews.EmailMessage(@this) {
				Subject = subject,
				Body = new Ews.MessageBody(isRichText ? Ews.BodyType.HTML : Ews.BodyType.Text, body),
			};
			email.ToRecipients.AddRange(to.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()));

			setting?.Invoke(email);

			email.Send();
		}

		/// <summary>
		/// メールを送信します。[非同期]
		/// </summary>
		/// <param name="this">ExchangeService</param>
		/// <param name="to">宛先 (; 区切りで複数指定できます。)</param>
		/// <param name="subject">件名</param>
		/// <param name="body">本文</param>
		/// <param name="isRichText">リッチテキストかどうかを指定します。</param>
		/// <param name="setting">メールの設定をするメソッド</param>
		public static async Task SendMailAsync(this Ews.ExchangeService @this, string to, string subject, string body, bool isRichText = false, Action<Ews.EmailMessage> setting = null)
			=> await Task.Run(() => @this.SendMail(to, subject, body, isRichText, setting));

		#endregion

		#region 予定取得

		/// <summary>
		/// 予定情報を取得します。
		/// </summary>
		/// <param name="this">ExchangeService</param>
		/// <param name="startDate">開始時刻</param>
		/// <param name="endDate">終了時刻</param>
		/// <param name="maxItemsReturned">取得最大数</param>
		/// <returns>予定情報を返します。</returns>
		public static Ews.FindItemsResults<Ews.Appointment> FindAppointments(this Ews.ExchangeService @this, DateTime startDate, DateTime endDate, int? maxItemsReturned = null) {
			// カレンダーフォルダオブジェクトをフォルダIDのみで初期化します。
			var calendar = Ews.CalendarFolder.Bind(@this, Ews.WellKnownFolderName.Calendar, new Ews.PropertySet());

			// 取得する予定の開始時刻と終了時刻、および予定の数を設定します。
			var cView = maxItemsReturned.HasValue
				? new Ews.CalendarView(startDate, endDate, maxItemsReturned.Value)
				: new Ews.CalendarView(startDate, endDate);

			// 返されるプロパティは、予定の件名、開始時刻、および終了時刻に制限します。
			cView.PropertySet = new Ews.PropertySet(Ews.ItemSchema.Subject, Ews.AppointmentSchema.Start, Ews.AppointmentSchema.End);

			// カレンダービューを使用して予定のコレクションを取得します。
			var appointments = calendar.FindAppointments(cView);
			return appointments;
		}

		/// <summary>
		/// 予定情報を取得します。[非同期]
		/// </summary>
		/// <param name="this">ExchangeService</param>
		/// <param name="startDate">開始時刻</param>
		/// <param name="endDate">終了時刻</param>
		/// <param name="maxItemsReturned">取得最大数</param>
		/// <returns>予定情報を返します。</returns>
		public static async Task<Ews.FindItemsResults<Ews.Appointment>> FindAppointmentsAsync(this Ews.ExchangeService @this, DateTime startDate, DateTime endDate, int? maxItemsReturned = default(int?))
			=> await Task.Run(() => @this.FindAppointments(startDate, endDate, maxItemsReturned));

		#endregion

		#region 空き時間確認

		#region GetUserAvailability (オーバーロード)

		/// <summary>
		/// 指定した時間枠内のユーザー、ルーム、リソースのセットの可用性に関する詳細情報を取得します。
		/// </summary>
		/// <param name="this">ExchangeService</param>
		/// <param name="attendees">可用性情報を取得する出席者。</param>
		/// <param name="options">返される情報を制御するオプション。</param>
		/// <param name="requestedData">要求されたデータ。(フリー/ビジーおよび/または提案)</param>
		/// <returns>各ユーザーの可用性情報が表示されます。
		/// 要求内のユーザーの順序によって、応答内の各ユーザーの可用性データの順序が決まります。</returns>
		public static Ews.GetUserAvailabilityResults GetUserAvailability(this Ews.ExchangeService @this, IEnumerable<Ews.AttendeeInfo> attendees, Ews.AvailabilityOptions options, Ews.AvailabilityData requestedData = Ews.AvailabilityData.FreeBusyAndSuggestions)
			=> @this.GetUserAvailability(attendees, options.DetailedSuggestionsWindow, requestedData, options);

		/// <summary>
		/// 指定した時間枠内のユーザー、ルーム、リソースのセットの可用性に関する詳細情報を取得します。
		/// </summary>
		/// <param name="this">ExchangeService</param>
		/// <param name="addresses">可用性情報を取得する出席者のアドレス。</param>
		/// <param name="options">返される情報を制御するオプション。</param>
		/// <param name="requestedData">要求されたデータ。(フリー/ビジーおよび/または提案)</param>
		/// <returns>各ユーザーの可用性情報が表示されます。
		/// 要求内のユーザーの順序によって、応答内の各ユーザーの可用性データの順序が決まります。</returns>
		public static Ews.GetUserAvailabilityResults GetUserAvailability(this Ews.ExchangeService @this, IEnumerable<string> addresses, Ews.AvailabilityOptions options, Ews.AvailabilityData requestedData = Ews.AvailabilityData.FreeBusyAndSuggestions)
			=> @this.GetUserAvailability(addresses.Select(a => new Ews.AttendeeInfo(a)), options, requestedData);

		#endregion

		#region GetUserAvailabilityAsync (オーバーロード)

		/// <summary>
		/// 非同期で
		/// 指定した時間枠内のユーザー、ルーム、リソースのセットの可用性に関する詳細情報を取得します。
		/// </summary>
		/// <param name="this">ExchangeService</param>
		/// <param name="attendees">可用性情報を取得する出席者。</param>
		/// <param name="options">返される情報を制御するオプション。</param>
		/// <param name="requestedData">要求されたデータ。(フリー/ビジーおよび/または提案)</param>
		/// <returns>各ユーザーの可用性情報が表示されます。
		/// 要求内のユーザーの順序によって、応答内の各ユーザーの可用性データの順序が決まります。</returns>
		public static async Task<Ews.GetUserAvailabilityResults> GetUserAvailabilityAsync(this Ews.ExchangeService @this, IEnumerable<Ews.AttendeeInfo> attendees, Ews.AvailabilityOptions options, Ews.AvailabilityData requestedData = Ews.AvailabilityData.FreeBusyAndSuggestions)
			=> await Task.Run(() => @this.GetUserAvailability(attendees, options, requestedData));

		/// <summary>
		/// 非同期で
		/// 指定した時間枠内のユーザー、ルーム、リソースのセットの可用性に関する詳細情報を取得します。
		/// </summary>
		/// <param name="this">ExchangeService</param>
		/// <param name="addresses">可用性情報を取得する出席者のアドレス。</param>
		/// <param name="options">返される情報を制御するオプション。</param>
		/// <param name="requestedData">要求されたデータ。(フリー/ビジーおよび/または提案)</param>
		/// <returns>各ユーザーの可用性情報が表示されます。
		/// 要求内のユーザーの順序によって、応答内の各ユーザーの可用性データの順序が決まります。</returns>
		public static async Task<Ews.GetUserAvailabilityResults> GetUserAvailabilityAsync(this Ews.ExchangeService @this, IEnumerable<string> addresses, Ews.AvailabilityOptions options, Ews.AvailabilityData requestedData = Ews.AvailabilityData.FreeBusyAndSuggestions)
			=> await @this.GetUserAvailabilityAsync(addresses.Select(a => new Ews.AttendeeInfo(a)), options, requestedData);

		#endregion

		#endregion

		#endregion
	}
}
