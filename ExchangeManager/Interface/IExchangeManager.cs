using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager.Interface {
	/// <summary>
	/// Exchange を管理するインターフェイスを提供します。
	/// </summary>
	public interface IExchangeManager {
		#region プロパティ

		/// <summary>
		/// Exchange Webサービスへのバインドを取得します。
		/// </summary>
		Ews.ExchangeService Service { get; }

		/// <summary>
		/// ユーザー名 (メールアドレス)
		/// </summary>
		string UserName { get; }

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
		Ews.Appointment Save(string subject, DateTime start, DateTime end, Action<Ews.Appointment> setting = null);

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
		Task<Ews.Appointment> SaveAsync(string subject, DateTime start, DateTime end, Action<Ews.Appointment> setting = null);

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
		Ews.Appointment Update(string uniqueId, Action<Ews.Appointment> update);

		#endregion

		#region UpdateAsync

		/// <summary>
		/// 非同期で
		/// 予定を更新します。
		/// </summary>
		/// <param name="uniqueId">予定のID</param>
		/// <param name="update">更新するメソッド</param>
		/// <returns>予定を返します。</returns>
		Task<Ews.Appointment> UpdateAsync(string uniqueId, Action<Ews.Appointment> update);

		#endregion

		#endregion

		#region 空き時間確認

		/// <summary>
		/// 指定した時間枠内のユーザー、ルーム、リソースのセットの可用性に関する詳細情報を取得します。
		/// </summary>
		/// <param name="addresses">可用性情報を取得する出席者のアドレス。</param>
		/// <param name="options">返される情報を制御するオプション。</param>
		/// <param name="requestedData">要求されたデータ。(フリー/ビジーおよび/または提案)</param>
		/// <returns>各ユーザーの可用性情報が表示されます。
		/// 要求内のユーザーの順序によって、応答内の各ユーザーの可用性データの順序が決まります。</returns>
		Ews.GetUserAvailabilityResults GetUserAvailability(IEnumerable<string> addresses, Ews.AvailabilityOptions options, Ews.AvailabilityData requestedData = Ews.AvailabilityData.FreeBusyAndSuggestions);

		/// <summary>
		/// 非同期で
		/// 指定した時間枠内のユーザー、ルーム、リソースのセットの可用性に関する詳細情報を取得します。
		/// </summary>
		/// <param name="addresses">可用性情報を取得する出席者のアドレス。</param>
		/// <param name="options">返される情報を制御するオプション。</param>
		/// <param name="requestedData">要求されたデータ。(フリー/ビジーおよび/または提案)</param>
		/// <returns>各ユーザーの可用性情報が表示されます。
		/// 要求内のユーザーの順序によって、応答内の各ユーザーの可用性データの順序が決まります。</returns>
		Task<Ews.GetUserAvailabilityResults> GetUserAvailabilityAsync(IEnumerable<string> addresses, Ews.AvailabilityOptions options, Ews.AvailabilityData requestedData = Ews.AvailabilityData.FreeBusyAndSuggestions);

		#endregion

		#region Bind

		/// <summary>
		/// 既存の予定にバインドし、その最初のクラスのプロパティを読み込みます。
		/// このメソッドを呼び出すと、EWSが呼び出されます。
		/// </summary>
		/// <param name="uniqueId">予定ID</param>
		/// <param name="additionalProperties">読み込むプロパティの配列</param>
		/// <returns>予定を返します。</returns>
		Ews.Appointment Bind(string uniqueId, params Ews.PropertyDefinitionBase[] additionalProperties);

		/// <summary>
		/// 既存の予定にバインドし、その最初のクラスのプロパティを読み込みます。
		/// このメソッドを呼び出すと、EWSが呼び出されます。
		/// </summary>
		/// <param name="id">予定ID</param>
		/// <param name="additionalProperties">読み込むプロパティの配列</param>
		/// <returns>予定を返します。</returns>
		Ews.Appointment Bind(Ews.ItemId id, params Ews.PropertyDefinitionBase[] additionalProperties);

		/// <summary>
		/// 非同期で
		/// 既存の予定にバインドし、その最初のクラスのプロパティを読み込みます。
		/// このメソッドを呼び出すと、EWSが呼び出されます。
		/// </summary>
		/// <param name="uniqueId">予定ID</param>
		/// <param name="additionalProperties">読み込むプロパティの配列</param>
		/// <returns>予定を返します。</returns>
		Task<Ews.Appointment> BindAsync(string uniqueId, params Ews.PropertyDefinitionBase[] additionalProperties);

		/// <summary>
		/// 非同期で
		/// 既存の予定にバインドし、その最初のクラスのプロパティを読み込みます。
		/// このメソッドを呼び出すと、EWSが呼び出されます。
		/// </summary>
		/// <param name="id">予定ID</param>
		/// <param name="additionalProperties">読み込むプロパティの配列</param>
		/// <returns>予定を返します。</returns>
		Task<Ews.Appointment> BindAsync(Ews.ItemId id, params Ews.PropertyDefinitionBase[] additionalProperties);

		#endregion

		#region 会議室一覧コレクション取得

		/// <summary>
		/// 会議室一覧のコレクションを取得します。
		/// </summary>
		IEnumerable<Ews.EmailAddress> GetRoomLists();

		/// <summary>
		/// 非同期で
		/// 会議室一覧のコレクションを取得します。
		/// </summary>
		Task<IEnumerable<Ews.EmailAddress>> GetRoomListsAsync();

		#endregion

		#region 会議室コレクション取得

		/// <summary>
		/// 非同期で
		/// 会議室のコレクションを取得します。
		/// </summary>
		IEnumerable<Ews.EmailAddress> GetRooms();

		/// <summary>
		/// 会議室のコレクションを取得します。
		/// </summary>
		Task<IEnumerable<Ews.EmailAddress>> GetRoomsAsync();

		/// <summary>
		/// 非同期で
		/// 会議室のコレクションを取得します。
		/// </summary>
		IEnumerable<Ews.AttendeeInfo> GetRoomsAsAttendee();

		/// <summary>
		/// 会議室のコレクションを取得します。
		/// </summary>
		Task<IEnumerable<Ews.AttendeeInfo>> GetRoomsAsAttendeeAsync();

		#endregion

		#endregion
	}
}
