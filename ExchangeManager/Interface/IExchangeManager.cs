using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager.Interface {
	public interface IExchangeManager {
		#region プロパティ

		/// <summary>
		/// ユーザー名 (メールアドレス)
		/// </summary>
		string UserName { get; }

		/// <summary>
		/// パスワード
		/// </summary>
		string Password { get; }

		#endregion

		#region メソッド

		#region メール送信

		/// <summary>
		/// メールを送信します。
		/// </summary>
		/// <param name="to">宛先 (; 区切りで複数指定できます。)</param>
		/// <param name="subject">件名</param>
		/// <param name="body">本文</param>
		/// <param name="isRichText">リッチテキストかどうかを指定します。</param>
		/// <param name="setting">メールの設定をするメソッド</param>
		void SendMail(string to, string subject, string body, bool isRichText = false, Action<Ews.EmailMessage> setting = null);

		/// <summary>
		/// メールを送信します。[非同期]
		/// </summary>
		/// <param name="to">宛先 (; 区切りで複数指定できます。)</param>
		/// <param name="subject">件名</param>
		/// <param name="body">本文</param>
		/// <param name="isRichText">リッチテキストかどうかを指定します。</param>
		/// <param name="setting">メールの設定をするメソッド</param>
		Task SendMailAsync(string to, string subject, string body, bool isRichText = false, Action<Ews.EmailMessage> setting = null);

		#endregion

		#region 予定取得

		/// <summary>
		/// 予定情報を取得します。
		/// </summary>
		/// <param name="startDate">開始時刻</param>
		/// <param name="endDate">終了時刻</param>
		/// <param name="maxItemsReturned">取得最大数</param>
		/// <returns>予定情報を返します。</returns>
		Ews.FindItemsResults<Ews.Appointment> FindAppointments(DateTime startDate, DateTime endDate, int? maxItemsReturned = null);

		/// <summary>
		/// 予定情報を取得します。[非同期]
		/// </summary>
		/// <param name="startDate">開始時刻</param>
		/// <param name="endDate">終了時刻</param>
		/// <param name="maxItemsReturned">取得最大数</param>
		/// <returns>予定情報を返します。</returns>
		Task<Ews.FindItemsResults<Ews.Appointment>> FindAppointmentsAsync(DateTime startDate, DateTime endDate, int? maxItemsReturned = null);

		#endregion

		#region 空き時間確認

		/// <summary>
		/// 指定した時間枠内のユーザー、ルーム、リソースのセットの可用性に関する詳細情報を取得します。
		/// </summary>
		/// <param name="attendees">可用性情報を取得する出席者。</param>
		/// <param name="options">返される情報を制御するオプション。</param>
		/// <param name="requestedData">要求されたデータ。(フリー/ビジーおよび/または提案)</param>
		/// <returns>各ユーザーの可用性情報が表示されます。
		/// 要求内のユーザーの順序によって、応答内の各ユーザーの可用性データの順序が決まります。</returns>
		Ews.GetUserAvailabilityResults GetUserAvailability(IEnumerable<Ews.AttendeeInfo> attendees, Ews.AvailabilityOptions options, Ews.AvailabilityData requestedData = Ews.AvailabilityData.FreeBusyAndSuggestions);

		/// <summary>
		/// 非同期で
		/// 指定した時間枠内のユーザー、ルーム、リソースのセットの可用性に関する詳細情報を取得します。
		/// </summary>
		/// <param name="attendees">可用性情報を取得する出席者。</param>
		/// <param name="options">返される情報を制御するオプション。</param>
		/// <param name="requestedData">要求されたデータ。(フリー/ビジーおよび/または提案)</param>
		/// <returns>各ユーザーの可用性情報が表示されます。
		/// 要求内のユーザーの順序によって、応答内の各ユーザーの可用性データの順序が決まります。</returns>
		Task<Ews.GetUserAvailabilityResults> GetUserAvailabilityAsync(IEnumerable<Ews.AttendeeInfo> attendees, Ews.AvailabilityOptions options, Ews.AvailabilityData requestedData = Ews.AvailabilityData.FreeBusyAndSuggestions);

		#endregion

		#endregion
	}
}
