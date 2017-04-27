﻿using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExchangeManager.Extensions;
using ExchangeManager.Interface;
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
		/// <param name="username">ユーザー名 (メールアドレス)</param>
		/// <param name="password">パスワード</param>
		protected ExchangeManagerBase(string username, string password) {
			this.UserName = username;
			this.Password = password;
		}

		#endregion

		#region プロパティ

		/// <summary>
		/// ユーザー名 (メールアドレス)
		/// </summary>
		public string UserName { get; protected set; }

		/// <summary>
		/// パスワード
		/// </summary>
		public string Password { get; protected set; }

		/// <summary>
		/// ExchangeService
		/// </summary>
		public Ews.ExchangeService Service { get; protected set; }

		#endregion

		#region メソッド

		/// <summary>
		/// EWS の新しいインスタンスを生成します。
		/// </summary>
		/// <param name="username">ユーザー名 (メールアドレス)</param>
		/// <param name="password">パスワード</param>
		/// <returns>生成した EWS のインスタンスを返します。</returns>
		protected abstract Ews.ExchangeService CreateService(string username, string password);

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

		#endregion

		#endregion
	}
}
