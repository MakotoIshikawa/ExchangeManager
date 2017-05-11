using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager.Model {
	/// <summary>
	/// 会議の予定を表すクラスです。
	/// </summary>
	[JsonObject(nameof(MeetingModel))]
	public class MeetingModel : PlanModel {
		#region コンストラクタ

		public MeetingModel() {
		}

		/// <summary>
		/// 予定情報の新しいインスタンスを生成します。
		/// </summary>
		/// <param name="subject">件名</param>
		/// <param name="conferenceRoom">会議室情報</param>
		/// <param name="start">開始時刻</param>
		/// <param name="end">終了時刻</param>
		public MeetingModel(string subject, Ews.EmailAddress conferenceRoom, DateTime start, DateTime end)
			: base(subject, start, end) {
			this.ConferenceRoom = conferenceRoom;
			this.Location = conferenceRoom.Name;
		}

		/// <summary>
		/// 予定情報の新しいインスタンスを生成します。
		/// </summary>
		/// <param name="subject">件名</param>
		/// <param name="conferenceRoom">会議室情報</param>
		/// <param name="start">開始時刻</param>
		/// <param name="interval">間隔</param>
		public MeetingModel(string subject, Ews.EmailAddress conferenceRoom, DateTime start, TimeSpan interval)
			: base(subject, start, interval) {
			this.ConferenceRoom = conferenceRoom;
		}

		#endregion

		#region プロパティ

		/// <summary>
		/// 会議室の情報を取得、設定します。
		/// </summary>
		public Ews.EmailAddress ConferenceRoom { get; set; }

		/// <summary>
		/// 出席者の一覧を取得します。
		/// </summary>
		public List<string> Attendees { get; set; }

		#endregion
	}
}
