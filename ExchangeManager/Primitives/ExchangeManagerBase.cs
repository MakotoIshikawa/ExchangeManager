using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExchangeManager.Extensions;
using ExchangeManager.Interface;
using ExtensionsLibrary.Extensions;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager.Primitives {
	/// <summary>
	/// Exchange を管理する機能を提供する抽象クラスです。
	/// </summary>
	public abstract class ExchangeManagerBase : IExchangeManager {
		#region フィールド

		#endregion

		#region コンストラクタ

		/// <summary>
		/// コンストラクタ
		/// </summary>
		protected ExchangeManagerBase() {
		}

		#endregion

		#region プロパティ

		/// <summary>
		/// ExchangeService
		/// </summary>
		public Ews.ExchangeService Service { get; protected set; }

		#endregion

		#region メソッド

		#region EWS 生成

		/// <summary>
		/// EWS の新しいインスタンスを生成します。
		/// </summary>
		/// <returns>生成した EWS のインスタンスを返します。</returns>
		protected abstract Ews.ExchangeService CreateService();

		#endregion

		#region メール送信

		/// <summary>
		/// メールを送信します。
		/// </summary>
		/// <param name="to">宛先 (; 区切りで複数指定できます。)</param>
		/// <param name="subject">件名</param>
		/// <param name="body">本文</param>
		/// <param name="isRichText">リッチテキストかどうかを指定します。</param>
		/// <param name="setting">メールの設定をするメソッド</param>
		public virtual void SendMail(string to, string subject, string body, bool isRichText = false, Action<Ews.EmailMessage> setting = null)
			=> this.Service.SendMail(to, subject, body, isRichText, setting);

		/// <summary>
		/// 非同期で
		/// メールを送信します。
		/// </summary>
		/// <param name="to">宛先 (; 区切りで複数指定できます。)</param>
		/// <param name="subject">件名</param>
		/// <param name="body">本文</param>
		/// <param name="isRichText">リッチテキストかどうかを指定します。</param>
		/// <param name="setting">メールの設定をするメソッド</param>
		public virtual async Task SendMailAsync(string to, string subject, string body, bool isRichText = false, Action<Ews.EmailMessage> setting = null)
			=> await this.Service.SendMailAsync(to, subject, body, isRichText, setting);

		#endregion

		#region 予定取得

		/// <summary>
		/// 予定情報を取得します。
		/// </summary>
		/// <param name="startDate">開始時刻</param>
		/// <param name="endDate">終了時刻</param>
		/// <param name="maxItemsReturned">取得最大数</param>
		/// <returns>予定情報を返します。</returns>
		public virtual Ews.FindItemsResults<Ews.Appointment> FindAppointments(DateTime startDate, DateTime endDate, int? maxItemsReturned = null)
			=> this.Service.FindAppointments(startDate, endDate, maxItemsReturned);

		/// <summary>
		/// 非同期で
		/// 予定情報を取得します。
		/// </summary>
		/// <param name="startDate">開始時刻</param>
		/// <param name="endDate">終了時刻</param>
		/// <param name="maxItemsReturned">取得最大数</param>
		/// <returns>予定情報を返します。</returns>
		public virtual async Task<Ews.FindItemsResults<Ews.Appointment>> FindAppointmentsAsync(DateTime startDate, DateTime endDate, int? maxItemsReturned = default(int?))
			=> await this.Service.FindAppointmentsAsync(startDate, endDate, maxItemsReturned);

		#endregion

		#region 予定作成

		#region Save

		/// <summary>
		/// 予定を作成します。
		/// </summary>
		/// <param name="subject">件名</param>
		/// <param name="start">開始時刻</param>
		/// <param name="end">終了時刻</param>
		/// <param name="setting">予定の詳細設定をするメソッド</param>
		/// <returns>予定を返します。</returns>
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

		#region SaveAsync

		/// <summary>
		/// 非同期で
		/// 予定を作成します。
		/// </summary>
		/// <param name="subject">件名</param>
		/// <param name="start">開始時刻</param>
		/// <param name="end">終了時刻</param>
		/// <param name="setting">予定の詳細設定をするメソッド</param>
		/// <returns>予定を返します。</returns>
		public async Task<Ews.Appointment> SaveAsync(string subject, DateTime start, DateTime end, Action<Ews.Appointment> setting = null)
			=> await Task.Run(() => Save(this.Service, subject, start, end, setting));

		#endregion

		#endregion

		#region 予定更新

		#region Update

		/// <summary>
		/// 予定を更新します。
		/// </summary>
		/// <param name="uniqueId">予定のID</param>
		/// <param name="update">更新するメソッド</param>
		/// <returns>予定を返します。</returns>
		public Ews.Appointment Update(string uniqueId, Action<Ews.Appointment> update)
			=> Update(this.Service, uniqueId, update);

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

		#endregion

		#region UpdateAsync

		/// <summary>
		/// 非同期で
		/// 予定を更新します。
		/// </summary>
		/// <param name="uniqueId">予定のID</param>
		/// <param name="update">更新するメソッド</param>
		/// <returns>予定を返します。</returns>
		public async Task<Ews.Appointment> UpdateAsync(string uniqueId, Action<Ews.Appointment> update)
			=> await UpdateAsync(this.Service, uniqueId, update);

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

			return await appointment.UpdateAsync(update);
		}

		#endregion

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
		public virtual Ews.GetUserAvailabilityResults GetUserAvailability(IEnumerable<Ews.AttendeeInfo> attendees, Ews.AvailabilityOptions options, Ews.AvailabilityData requestedData = Ews.AvailabilityData.FreeBusyAndSuggestions)
			=> this.Service.GetUserAvailability(attendees, options, requestedData);

		/// <summary>
		/// 指定した時間枠内のユーザー、ルーム、リソースのセットの可用性に関する詳細情報を取得します。
		/// </summary>
		/// <param name="addresses">可用性情報を取得する出席者のアドレス。</param>
		/// <param name="options">返される情報を制御するオプション。</param>
		/// <param name="requestedData">要求されたデータ。(フリー/ビジーおよび/または提案)</param>
		/// <returns>各ユーザーの可用性情報が表示されます。
		/// 要求内のユーザーの順序によって、応答内の各ユーザーの可用性データの順序が決まります。</returns>
		public virtual Ews.GetUserAvailabilityResults GetUserAvailability(IEnumerable<string> addresses, Ews.AvailabilityOptions options, Ews.AvailabilityData requestedData = Ews.AvailabilityData.FreeBusyAndSuggestions)
			=> this.Service.GetUserAvailability(addresses, options, requestedData);

		/// <summary>
		/// 非同期で
		/// 指定した時間枠内のユーザー、ルーム、リソースのセットの可用性に関する詳細情報を取得します。
		/// </summary>
		/// <param name="attendees">可用性情報を取得する出席者。</param>
		/// <param name="options">返される情報を制御するオプション。</param>
		/// <param name="requestedData">要求されたデータ。(フリー/ビジーおよび/または提案)</param>
		/// <returns>各ユーザーの可用性情報が表示されます。
		/// 要求内のユーザーの順序によって、応答内の各ユーザーの可用性データの順序が決まります。</returns>
		public virtual async Task<Ews.GetUserAvailabilityResults> GetUserAvailabilityAsync(IEnumerable<Ews.AttendeeInfo> attendees, Ews.AvailabilityOptions options, Ews.AvailabilityData requestedData = Ews.AvailabilityData.FreeBusyAndSuggestions)
			=> await this.Service.GetUserAvailabilityAsync(attendees, options, requestedData);

		/// <summary>
		/// 非同期で
		/// 指定した時間枠内のユーザー、ルーム、リソースのセットの可用性に関する詳細情報を取得します。
		/// </summary>
		/// <param name="addresses">可用性情報を取得する出席者のアドレス。</param>
		/// <param name="options">返される情報を制御するオプション。</param>
		/// <param name="requestedData">要求されたデータ。(フリー/ビジーおよび/または提案)</param>
		/// <returns>各ユーザーの可用性情報が表示されます。
		/// 要求内のユーザーの順序によって、応答内の各ユーザーの可用性データの順序が決まります。</returns>
		public virtual async Task<Ews.GetUserAvailabilityResults> GetUserAvailabilityAsync(IEnumerable<string> addresses, Ews.AvailabilityOptions options, Ews.AvailabilityData requestedData = Ews.AvailabilityData.FreeBusyAndSuggestions)
			=> await this.Service.GetUserAvailabilityAsync(addresses, options, requestedData);

		#endregion

		#region Bind

		/// <summary>
		/// 既存の予定にバインドし、その最初のクラスのプロパティを読み込みます。
		/// このメソッドを呼び出すと、EWSが呼び出されます。
		/// </summary>
		/// <param name="uniqueId"></param>
		/// <param name="additionalProperties"></param>
		/// <returns></returns>
		public Ews.Appointment Bind(string uniqueId, params Ews.PropertyDefinitionBase[] additionalProperties)
			=> this.Bind(new Ews.ItemId(uniqueId), additionalProperties);

		public Ews.Appointment Bind(Ews.ItemId id, params Ews.PropertyDefinitionBase[] additionalProperties)
			=> Bind(this.Service, id, additionalProperties);

		public async Task<Ews.Appointment> BindAsync(string uniqueId, params Ews.PropertyDefinitionBase[] additionalProperties)
			=> await this.BindAsync(new Ews.ItemId(uniqueId), additionalProperties);

		public async Task<Ews.Appointment> BindAsync(Ews.ItemId id, params Ews.PropertyDefinitionBase[] additionalProperties)
			=> await Task.Run(() => Bind(this.Service, id, additionalProperties));

		private static Ews.Appointment Bind(Ews.ExchangeService service, Ews.ItemId id, params Ews.PropertyDefinitionBase[] additionalProperties)
			=> additionalProperties?.Any() ?? false
				? Ews.Appointment.Bind(service, id, new Ews.PropertySet(additionalProperties))
				: Ews.Appointment.Bind(service, id);

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
		public IEnumerable<Ews.EmailAddress> GetRooms()
			=> this.GetRoomLists()?
				.SelectMany(rl => this.Service.GetRooms(rl.Address))?
				.Distinct(r => r.Address);

		/// <summary>
		/// 会議室のコレクションを取得します。
		/// </summary>
		public async Task<IEnumerable<Ews.EmailAddress>> GetRoomsAsync()
			=> (await this.GetRoomListsAsync())?
				.SelectMany(rl => this.Service.GetRooms(rl.Address))?
				.Distinct(r => r.Address);

		/// <summary>
		/// 非同期で
		/// 会議室のコレクションを取得します。
		/// </summary>
		public IEnumerable<Ews.AttendeeInfo> GetRoomsAsAttendee()
			=> this.GetRooms()?
				.Select(r => new Ews.AttendeeInfo {
					SmtpAddress = r.Address,
					AttendeeType = Ews.MeetingAttendeeType.Room,
				});

		/// <summary>
		/// 会議室のコレクションを取得します。
		/// </summary>
		public async Task<IEnumerable<Ews.AttendeeInfo>> GetRoomsAsAttendeeAsync()
			=> (await this.GetRoomsAsync())?
				.Select(r => new Ews.AttendeeInfo {
					SmtpAddress = r.Address,
					AttendeeType = Ews.MeetingAttendeeType.Room,
				});

		#endregion

		#endregion
	}
}
