using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager.Extensions {
	public static partial class ExchangeServiceExtension {
		public static async Task SendMailAsync(this Ews.ExchangeService service, string to, string subject, string body, bool isRichText = false, Action<Ews.EmailMessage> setting = null) {
			await Task.Run(() => service.SendMail(to, subject, body, isRichText, setting));
		}

		public static void SendMail(this Ews.ExchangeService service, string to, string subject, string body, bool isRichText = false, Action<Ews.EmailMessage> setting = null) {
			var email = new Ews.EmailMessage(service) {
				Subject = subject,
				Body = new Ews.MessageBody(isRichText ? Ews.BodyType.HTML : Ews.BodyType.Text, body),
			};
			email.ToRecipients.AddRange(to.Split(';').Select(s => s.Trim()));

			setting?.Invoke(email);

			email.Send();
		}

		public static Ews.FindItemsResults<Ews.Appointment> FindAppointments(this Ews.ExchangeService service, DateTime startDate, DateTime endDate, int? maxItemsReturned = null) {
			// カレンダーフォルダオブジェクトをフォルダIDのみで初期化します。
			var calendar = Ews.CalendarFolder.Bind(service, Ews.WellKnownFolderName.Calendar, new Ews.PropertySet());

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

		public static Ews.GetUserAvailabilityResults GetUserAvailability(this Ews.ExchangeService @this, IEnumerable<Ews.AttendeeInfo> attendees, DateTime startTime, DateTime endTime, int goodSuggestionThreshold, int maximumNonWorkHoursSuggestionsPerDay, int maximumSuggestionsPerDay, int meetingDuration = 60) {
			var detailedSuggestionsWindow = new Ews.TimeWindow(startTime, endTime);

			// 空き時間情報および推奨される会議時間を要求するオプションを指定します。
			var availabilityOptions = new Ews.AvailabilityOptions {
				GoodSuggestionThreshold = goodSuggestionThreshold,
				MaximumNonWorkHoursSuggestionsPerDay = maximumNonWorkHoursSuggestionsPerDay,
				MaximumSuggestionsPerDay = maximumSuggestionsPerDay,

				// MeetingDurationのデフォルト値は60分ですが、デモンストレーションの目的で明示的に設定することに注意してください。
				MeetingDuration = meetingDuration,

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
	}
}
