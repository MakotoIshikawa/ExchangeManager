using System.Collections.Generic;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager.Extensions {
	/// <summary>
	/// List の追加を拡張するメソッドを提供します。
	/// </summary>
	public static partial class ListAddExtension {
		#region メソッド

		#region Add

		/// <summary>
		/// 末尾に AttendeeInfo のインスタンスを追加します。
		/// </summary>
		/// <param name="this">AttendeeInfo のリスト</param>
		/// <param name="smtpAddress">SMTPアドレス</param>
		/// <param name="attendeeType">会議出席者のタイプ</param>
		/// <param name="excludeConflicts">この参加者が利用できない時間が返されるべきかどうかを示す値</param>
		public static void Add(this List<AttendeeInfo> @this, string smtpAddress, MeetingAttendeeType attendeeType, bool excludeConflicts = false)
			=> @this.Add(new AttendeeInfo {
				SmtpAddress = smtpAddress,
				AttendeeType = attendeeType,
				ExcludeConflicts = excludeConflicts,
			});

		/// <summary>
		/// 末尾に EmailAddress のインスタンスを追加します。
		/// </summary>
		/// <param name="this">EmailAddress のリスト</param>
		/// <param name="address">アドレス</param>
		/// <param name="name">名前</param>
		public static void Add(this List<EmailAddress> @this, string address, string name = null)
			=> @this.Add(new EmailAddress(address) {
				Name = name,
			});

		#endregion

		#endregion
	}
}
