using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager.Interface {
	public interface IExchangeManager {
		#region プロパティ

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

		#endregion

		#region 空き時間確認

		/// <summary>
		/// 空き時間を取得します。
		/// </summary>
		/// <param name="attendees">出席者</param>
		/// <param name="startTime">開始時間</param>
		/// <param name="endTime">終了時間</param>
		/// <param name="goodSuggestionThreshold">推奨される会議時間としての資格を得るために、その期間に期間を開いておく必要がある出席者の割合を取得または設定します。1～49でなければなりません。デフォルト値は25です。</param>
		/// <param name="maximumNonWorkHoursSuggestionsPerDay">1日あたりの通常の営業時間外に推奨される会議時間の数を取得または設定します。0～48の間でなければなりません。デフォルト値は0です。</param>
		/// <param name="maximumSuggestionsPerDay">1日に返される推奨会議時間の数を取得または設定します。0～48の間でなければなりません。デフォルト値は10です。</param>
		/// <param name="meetingDuration">提案を取得する会議の所要時間を分単位で取得または設定します。30～1440でなければなりません。既定値は60です。</param>
		/// <returns>空き時間の情報を返します。</returns>
		Ews.GetUserAvailabilityResults GetUserAvailability(IEnumerable<Ews.AttendeeInfo> attendees, DateTime startTime, DateTime endTime, int goodSuggestionThreshold = 25, int maximumNonWorkHoursSuggestionsPerDay = 0, int maximumSuggestionsPerDay = 10, int meetingDuration = 60);

		#endregion

		#endregion
	}
}
