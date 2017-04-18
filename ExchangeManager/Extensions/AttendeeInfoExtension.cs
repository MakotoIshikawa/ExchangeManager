using System.Collections.Generic;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager.Extensions {
	/// <summary>
	/// AttendeeInfo を拡張するメソッドを提供します。
	/// </summary>
	public static partial class AttendeeInfoExtension {
		/// <summary>
		/// 末尾に AttendeeInfo のインスタンスを追加します。
		/// </summary>
		/// <param name="this"></param>
		/// <param name="smtpAddress">SMTPアドレス</param>
		/// <param name="attendeeType">会議出席者のタイプ</param>
		public static void Add(this List<AttendeeInfo> @this, string smtpAddress, MeetingAttendeeType attendeeType) {
			@this.Add(new AttendeeInfo() {
				SmtpAddress = smtpAddress,
				AttendeeType = attendeeType
			});
		}
	}
}
