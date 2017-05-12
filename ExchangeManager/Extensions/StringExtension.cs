using System;
using System.Net.Mail;

namespace ExchangeManager.Extensions {
	/// <summary>
	/// String を拡張するメソッドを提供します。
	/// </summary>
	public static partial class StringExtension {
		#region メソッド

		/// <summary>
		/// 文字列がメールアドレス形式かどうか判定します。
		/// </summary>
		/// <param name="this">String</param>
		/// <returns>メールアドレス形式であれば true を返します。</returns>
		public static bool IsMailAddress(this string @this) {
			try {
				var a = new MailAddress(@this);

				return true;
			} catch (Exception) {
				return false;
			}
		}

		#endregion
	}
}