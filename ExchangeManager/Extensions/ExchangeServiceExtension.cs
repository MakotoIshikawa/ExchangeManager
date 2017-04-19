using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
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

		/// <summary>
		/// 空き時間を取得します。
		/// </summary>
		/// <param name="this">ExchangeService</param>
		/// <param name="attendees">出席者</param>
		/// <param name="startTime">開始時間</param>
		/// <param name="endTime">終了時間</param>
		/// <param name="goodSuggestionThreshold">推奨される会議時間としての資格を得るために、その期間に期間を開いておく必要がある出席者の割合を取得または設定します。1～49でなければなりません。デフォルト値は25です。</param>
		/// <param name="maximumNonWorkHoursSuggestionsPerDay">1日あたりの通常の営業時間外に推奨される会議時間の数を取得または設定します。0～48の間でなければなりません。デフォルト値は0です。</param>
		/// <param name="maximumSuggestionsPerDay">1日に返される推奨会議時間の数を取得または設定します。0～48の間でなければなりません。デフォルト値は10です。</param>
		/// <param name="meetingDuration">提案を取得する会議の所要時間を分単位で取得または設定します。30～1440でなければなりません。既定値は60です。</param>
		/// <returns>空き時間の情報を返します。</returns>
		public static Ews.GetUserAvailabilityResults GetUserAvailability(this Ews.ExchangeService @this, IEnumerable<Ews.AttendeeInfo> attendees, DateTime startTime, DateTime endTime, int goodSuggestionThreshold = 25, int maximumNonWorkHoursSuggestionsPerDay = 0, int maximumSuggestionsPerDay = 10, int meetingDuration = 60) {
			var detailedSuggestionsWindow = new Ews.TimeWindow(startTime, endTime);

			// 空き時間情報および推奨される会議時間を要求するオプションを指定します。
			var availabilityOptions = new Ews.AvailabilityOptions {
				GoodSuggestionThreshold = goodSuggestionThreshold.WithinRange(1, 49),
				MaximumNonWorkHoursSuggestionsPerDay = maximumNonWorkHoursSuggestionsPerDay.WithinRange(0, 48),
				MaximumSuggestionsPerDay = maximumSuggestionsPerDay.WithinRange(0, 48),

				// MeetingDurationのデフォルト値は60分ですが、デモンストレーションの目的で明示的に設定することに注意してください。
				MeetingDuration = meetingDuration.WithinRange(30, 1440),

				MinimumSuggestionQuality = Ews.SuggestionQuality.Good,
				DetailedSuggestionsWindow = detailedSuggestionsWindow,
				RequestedFreeBusyView = Ews.FreeBusyViewType.FreeBusy,
			};

			// 空き時間情報と推奨される会議時間のセットを返します。
			// このメソッドの結果、EWSへの GetUserAvailabilityRequest が呼び出されます。
			var results = @this.GetUserAvailability(attendees, detailedSuggestionsWindow,
				Ews.AvailabilityData.FreeBusyAndSuggestions, availabilityOptions);

			return results;
		}


		/// <summary>
		/// 空き時間を取得します。[非同期]
		/// </summary>
		/// <param name="this">ExchangeService</param>
		/// <param name="attendees">出席者</param>
		/// <param name="startTime">開始時間</param>
		/// <param name="endTime">終了時間</param>
		/// <param name="goodSuggestionThreshold">推奨される会議時間としての資格を得るために、その期間に期間を開いておく必要がある出席者の割合を取得または設定します。1～49でなければなりません。デフォルト値は25です。</param>
		/// <param name="maximumNonWorkHoursSuggestionsPerDay">1日あたりの通常の営業時間外に推奨される会議時間の数を取得または設定します。0～48の間でなければなりません。デフォルト値は0です。</param>
		/// <param name="maximumSuggestionsPerDay">1日に返される推奨会議時間の数を取得または設定します。0～48の間でなければなりません。デフォルト値は10です。</param>
		/// <param name="meetingDuration">提案を取得する会議の所要時間を分単位で取得または設定します。30～1440でなければなりません。既定値は60です。</param>
		/// <returns>空き時間の情報を返します。</returns>
		public static async Task<Ews.GetUserAvailabilityResults> GetUserAvailabilityAsync(this Ews.ExchangeService @this, IEnumerable<Ews.AttendeeInfo> attendees, DateTime startTime, DateTime endTime, int goodSuggestionThreshold = 25, int maximumNonWorkHoursSuggestionsPerDay = 0, int maximumSuggestionsPerDay = 10, int meetingDuration = 60)
			=> await Task.Run(() => @this.GetUserAvailability(attendees, startTime, endTime, goodSuggestionThreshold, maximumNonWorkHoursSuggestionsPerDay, maximumSuggestionsPerDay, meetingDuration));

		#endregion

		#endregion
	}

	/// <summary>
	/// 数値関数の拡張メソッドを提供します。
	/// </summary>
	public static partial class MathExtension {
		#region メソッド

		#region WithinRange

		/// <summary>
		/// 上限と下限を設定して、範囲内の値を取得します。
		/// </summary>
		/// <param name="value">値</param>
		/// <param name="min">下限値</param>
		/// <param name="max">上限値</param>
		/// <returns>範囲内の値を返します。</returns>
		public static short WithinRange(this short value, short min, short max) {
			return Math.Min(max, Math.Max(min, value));
		}

		/// <summary>
		/// 上限と下限を設定して、範囲内の値を取得します。
		/// </summary>
		/// <param name="value">値</param>
		/// <param name="min">下限値</param>
		/// <param name="max">上限値</param>
		/// <returns>範囲内の値を返します。</returns>
		public static ushort WithinRange(this ushort value, ushort min, ushort max) {
			return Math.Min(max, Math.Max(min, value));
		}

		/// <summary>
		/// 上限と下限を設定して、範囲内の値を取得します。
		/// </summary>
		/// <param name="value">値</param>
		/// <param name="min">下限値</param>
		/// <param name="max">上限値</param>
		/// <returns>範囲内の値を返します。</returns>
		public static int WithinRange(this int value, int min, int max) {
			return Math.Min(max, Math.Max(min, value));
		}

		/// <summary>
		/// 上限と下限を設定して、範囲内の値を取得します。
		/// </summary>
		/// <param name="value">値</param>
		/// <param name="min">下限値</param>
		/// <param name="max">上限値</param>
		/// <returns>範囲内の値を返します。</returns>
		public static uint WithinRange(this uint value, uint min, uint max) {
			return Math.Min(max, Math.Max(min, value));
		}

		/// <summary>
		/// 上限と下限を設定して、範囲内の値を取得します。
		/// </summary>
		/// <param name="value">値</param>
		/// <param name="min">下限値</param>
		/// <param name="max">上限値</param>
		/// <returns>範囲内の値を返します。</returns>
		public static long WithinRange(this long value, long min, long max) {
			return Math.Min(max, Math.Max(min, value));
		}

		/// <summary>
		/// 上限と下限を設定して、範囲内の値を取得します。
		/// </summary>
		/// <param name="value">値</param>
		/// <param name="min">下限値</param>
		/// <param name="max">上限値</param>
		/// <returns>範囲内の値を返します。</returns>
		public static ulong WithinRange(this ulong value, ulong min, ulong max) {
			return Math.Min(max, Math.Max(min, value));
		}

		/// <summary>
		/// 上限と下限を設定して、範囲内の値を取得します。
		/// </summary>
		/// <param name="value">値</param>
		/// <param name="min">下限値</param>
		/// <param name="max">上限値</param>
		/// <returns>範囲内の値を返します。</returns>
		public static float WithinRange(this float value, float min, float max) {
			return Math.Min(max, Math.Max(min, value));
		}

		/// <summary>
		/// 上限と下限を設定して、範囲内の値を取得します。
		/// </summary>
		/// <param name="value">値</param>
		/// <param name="min">下限値</param>
		/// <param name="max">上限値</param>
		/// <returns>範囲内の値を返します。</returns>
		public static double WithinRange(this double value, double min, double max) {
			return Math.Min(max, Math.Max(min, value));
		}

		/// <summary>
		/// 上限と下限を設定して、範囲内の値を取得します。
		/// </summary>
		/// <param name="value">値</param>
		/// <param name="min">下限値</param>
		/// <param name="max">上限値</param>
		/// <returns>範囲内の値を返します。</returns>
		public static decimal WithinRange(this decimal value, decimal min, decimal max) {
			return Math.Min(max, Math.Max(min, value));
		}

		#endregion

		#endregion
	}
}
