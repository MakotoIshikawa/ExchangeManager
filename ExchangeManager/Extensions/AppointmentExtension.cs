using System;
using System.Linq;
using System.Threading.Tasks;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager.Extensions {
	/// <summary>
	/// Appointment を拡張するメソッドを提供します。
	/// </summary>
	public static class AppointmentExtension {
		#region メソッド

		#region 更新

		/// <summary>
		/// このアポイントに加えられたローカル変更を適用します。
		/// このメソッドを呼び出すと、少なくとも1回EWSが呼び出されます。
		/// 添付ファイルが追加または削除された場合、EWSへの複数の呼び出しが行われる可能性があります。
		/// </summary>
		/// <param name="this">Appointment</param>
		/// <param name="update">更新するメソッド</param>
		/// <returns>アポイントを返します。</returns>
		public static Ews.Appointment Update(this Ews.Appointment @this, Action<Ews.Appointment> update) {
			update?.Invoke(@this);

			// 明示的に指定しない限り、デフォルトではSendToAllAndSaveCopyを使用します。
			// これにより、予定を会議に変換できます。
			// これを避けるには、非会議でSendToNoneを明示的に設定します。
			var mode = @this.IsMeeting
				? Ews.SendInvitationsOrCancellationsMode.SendToAllAndSaveCopy
				: Ews.SendInvitationsOrCancellationsMode.SendToNone;

			// 更新要求をExchangeサーバーに送信します。
			@this.Update(Ews.ConflictResolutionMode.AlwaysOverwrite, mode);

			return @this;
		}

		/// <summary>
		/// 非同期でこのアポイントに加えられたローカル変更を適用します。
		/// </summary>
		/// <param name="this">Appointment</param>
		/// <param name="update">更新するメソッド</param>
		/// <returns>アポイントを返します。</returns>
		public static async Task<Ews.Appointment> UpdateAsync(this Ews.Appointment @this, Action<Ews.Appointment> update)
			=> await Task.Run(() => @this.Update(update));

		#endregion

		#region 保存

		/// <summary>
		/// このアポイントをカレンダーフォルダに保存します。
		/// このメソッドを呼び出すと、少なくとも1回EWSが呼び出されます。
		/// 添付ファイルが追加されていれば、EWSへの複数の呼び出しが行われる可能性があります。
		/// </summary>
		/// <param name="this">Appointment</param>
		/// <param name="setting">アポイントの詳細設定をするメソッド</param>
		/// <returns>アポイントを返します。</returns>
		public static Ews.Appointment Save(this Ews.Appointment @this, Action<Ews.Appointment> setting = null) {
			setting?.Invoke(@this);

			var mode = (@this.RequiredAttendees.Any())
				? Ews.SendInvitationsMode.SendToAllAndSaveCopy
				: Ews.SendInvitationsMode.SendToNone;

			// 予定をカレンダーに保存します。
			@this.Save(mode);

			return @this;
		}

		/// <summary>
		/// 非同期でこのアポイントをカレンダーフォルダに保存します。
		/// </summary>
		/// <param name="this">Appointment</param>
		/// <param name="setting">アポイントの詳細設定をするメソッド</param>
		/// <returns>アポイントを返します。</returns>
		public static async Task<Ews.Appointment> SaveAsync(this Ews.Appointment @this, Action<Ews.Appointment> setting = null)
			=> await Task.Run(() => @this.Save(setting));

		#endregion

		#endregion
	}
}