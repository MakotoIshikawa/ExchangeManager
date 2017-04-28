using System;
using System.Collections.Generic;
using System.Linq;
using ExchangeManager.Interface;
using ExchangeManager.Primitives;
using ExtensionsLibrary.Extensions;
using System.Threading.Tasks;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager {
	/// <summary>
	/// Exchange Online の情報を管理するクラスです。
	/// </summary>
	public class ExchangeOnlineManager : ExchangeManagerBase, IExchangeManager {
		#region フィールド

		#endregion

		#region コンストラクタ

		/// <summary>
		/// コンストラクタ
		/// </summary>
		/// <param name="username">ユーザー名 (メールアドレス)</param>
		/// <param name="password">パスワード</param>
		public ExchangeOnlineManager(string username, string password) : base(username, password) {
			this.Service = this.CreateService(username, password);

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

		/// <summary>
		/// EWS の新しいインスタンスを生成します。
		/// </summary>
		/// <param name="username">ユーザー名 (メールアドレス)</param>
		/// <param name="password">パスワード</param>
		/// <returns>生成した EWS のインスタンスを返します。</returns>
		protected override Ews.ExchangeService CreateService(string username, string password) {
			return new Ews.ExchangeService(Ews.ExchangeVersion.Exchange2013_SP1) {
				Credentials = new Ews.WebCredentials(username, password),
				UseDefaultCredentials = false,
				TraceEnabled = true,
				TraceFlags = Ews.TraceFlags.All,
			};
		}

		#region Bind

		public Ews.Appointment Bind(string uniqueId, params Ews.PropertyDefinitionBase[] additionalProperties)
			=> this.Bind(new Ews.ItemId(uniqueId), additionalProperties);

		public Ews.Appointment Bind(Ews.ItemId id, params Ews.PropertyDefinitionBase[] additionalProperties)
			=> Bind(this.Service, id, additionalProperties);

		public async Task<Ews.Appointment> BindAsync(string uniqueId, params Ews.PropertyDefinitionBase[] additionalProperties)
			=> await this.BindAsync(new Ews.ItemId(uniqueId), additionalProperties);

		public async Task<Ews.Appointment> BindAsync(Ews.ItemId id, params Ews.PropertyDefinitionBase[] additionalProperties)
			=> await Task.Run(() => Bind(this.Service, id, additionalProperties));

		private static Ews.Appointment Bind(Ews.ExchangeService service, Ews.ItemId id, params Ews.PropertyDefinitionBase[] additionalProperties) {
			return additionalProperties?.Any() ?? false
				? Ews.Appointment.Bind(service, id, new Ews.PropertySet(additionalProperties))
				: Ews.Appointment.Bind(service, id);
		}

		#endregion

		#region 会議室一覧コレクション取得

		/// <summary>
		/// 会議室一覧のコレクションを取得します。
		/// </summary>
		public IEnumerable<Ews.EmailAddress> GetRoomLists()
			=> this.Service.GetRoomLists();

		/// <summary>
		/// 非同期で
		/// 会議室一覧のコレクションを取得します。
		/// </summary>
		public async Task<IEnumerable<Ews.EmailAddress>> GetRoomListsAsync()
			=> await Task.Run(() => this.GetRoomLists());

		#endregion

		#region 会議室コレクション取得

		/// <summary>
		/// 非同期で
		/// 会議室のコレクションを取得します。
		/// </summary>
		public IEnumerable<Ews.AttendeeInfo> GetRooms()
			=> this.GetRoomLists()?
				.SelectMany(rl => this.Service.GetRooms(rl.Address))?
				.Distinct(r => r.Address)?
				.Select(r => new Ews.AttendeeInfo {
					SmtpAddress = r.Address,
					AttendeeType = Ews.MeetingAttendeeType.Room,
				});

		/// <summary>
		/// 会議室のコレクションを取得します。
		/// </summary>
		public async Task<IEnumerable<Ews.AttendeeInfo>> GetRoomsAsync()
			=> (await this.GetRoomListsAsync())?
				.SelectMany(rl => this.Service.GetRooms(rl.Address))?
				.Distinct(r => r.Address)?
				.Select(r => new Ews.AttendeeInfo {
					SmtpAddress = r.Address,
					AttendeeType = Ews.MeetingAttendeeType.Room,
				});

		#endregion

		#region 予定作成

		public Ews.Appointment Save(string subject, DateTime start, DateTime end, Action<Ews.Appointment> setting = null)
			=> Save(this.Service, subject, start, end, setting);

		private static Ews.Appointment Save(Ews.ExchangeService service, string subject, DateTime start, DateTime end, Action<Ews.Appointment> setting) {
			// 予定を作成する予定オブジェクトのプロパティを設定します。
			var appointment = new Ews.Appointment(service) {
				Subject = subject,
				Start = start,
				End = end,
			};

			setting?.Invoke(appointment);

			var mode = (appointment.RequiredAttendees.Any())
				? Ews.SendInvitationsMode.SendToAllAndSaveCopy
				: Ews.SendInvitationsMode.SendToNone;

			// 予定をカレンダーに保存します。
			appointment.Save(mode);

			return appointment;
		}

		#endregion

		#region 予定更新

		public Ews.Appointment Update(string uniqueId, Action<Ews.Appointment> update)
			=> Update(this.Service, uniqueId, update);

		public async Task<Ews.Appointment> UpdateAsync(string uniqueId, Action<Ews.Appointment> update)
			=> await UpdateAsync(this.Service, uniqueId, update);

		private Ews.Appointment Update(Ews.ExchangeService service, string uniqueId, Action<Ews.Appointment> update) {
			var appointment = this.Bind(uniqueId,
				Ews.ItemSchema.Subject
				, Ews.ItemSchema.Body
				, Ews.AppointmentSchema.Start
				, Ews.AppointmentSchema.End
				, Ews.AppointmentSchema.Location
				, Ews.AppointmentSchema.IsMeeting
				, Ews.ItemSchema.Attachments
				, Ews.ItemSchema.Categories
				, Ews.AppointmentSchema.Duration
				, Ews.AppointmentSchema.RequiredAttendees
				, Ews.AppointmentSchema.Resources
				, Ews.ItemSchema.Id
				, Ews.AppointmentSchema.IsAllDayEvent
				, Ews.ItemSchema.IsAssociated
				, Ews.AppointmentSchema.IsCancelled
				, Ews.ItemSchema.IsDraft
				, Ews.ItemSchema.IsFromMe
				, Ews.AppointmentSchema.IsMeeting
				, Ews.AppointmentSchema.IsOnlineMeeting
				, Ews.AppointmentSchema.IsRecurring
				, Ews.ItemSchema.IsReminderSet
				, Ews.ItemSchema.IsResend
				, Ews.AppointmentSchema.IsResponseRequested
				, Ews.ItemSchema.IsSubmitted
				, Ews.ItemSchema.IsUnmodified
				, Ews.AppointmentSchema.TimeZone
			);

			return appointment.Update(update);
		}

		private async Task<Ews.Appointment> UpdateAsync(Ews.ExchangeService service, string uniqueId, Action<Ews.Appointment> update) {
			var appointment = await this.BindAsync(uniqueId,
				Ews.ItemSchema.Subject
				, Ews.ItemSchema.Body
				, Ews.AppointmentSchema.Start
				, Ews.AppointmentSchema.End
				, Ews.AppointmentSchema.Location
				, Ews.AppointmentSchema.IsMeeting
				, Ews.ItemSchema.Attachments
				, Ews.ItemSchema.Categories
				, Ews.AppointmentSchema.Duration
				, Ews.AppointmentSchema.RequiredAttendees
				, Ews.AppointmentSchema.Resources
				, Ews.ItemSchema.Id
				, Ews.AppointmentSchema.IsAllDayEvent
				, Ews.ItemSchema.IsAssociated
				, Ews.AppointmentSchema.IsCancelled
				, Ews.ItemSchema.IsDraft
				, Ews.ItemSchema.IsFromMe
				, Ews.AppointmentSchema.IsMeeting
				, Ews.AppointmentSchema.IsOnlineMeeting
				, Ews.AppointmentSchema.IsRecurring
				, Ews.ItemSchema.IsReminderSet
				, Ews.ItemSchema.IsResend
				, Ews.AppointmentSchema.IsResponseRequested
				, Ews.ItemSchema.IsSubmitted
				, Ews.ItemSchema.IsUnmodified
				, Ews.AppointmentSchema.TimeZone
			);

			return appointment.Update(update);
		}

		#endregion

		#endregion
	}

	public static class AppointmentExtension {

		public static Ews.Appointment Update(this Ews.Appointment appointment, Action<Ews.Appointment> update) {
			update?.Invoke(appointment);

			// 明示的に指定しない限り、デフォルトではSendToAllAndSaveCopyを使用します。
			// これにより、予定を会議に変換できます。
			// これを避けるには、非会議でSendToNoneを明示的に設定します。
			var mode = appointment.IsMeeting
				? Ews.SendInvitationsOrCancellationsMode.SendToAllAndSaveCopy
				: Ews.SendInvitationsOrCancellationsMode.SendToNone;

			// 更新要求をExchangeサーバーに送信します。
			appointment.Update(Ews.ConflictResolutionMode.AlwaysOverwrite, mode);

			return appointment;
		}
	}
}
