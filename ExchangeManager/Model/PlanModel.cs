using System;
using System.ComponentModel;
using ExchangeManager.Extensions;
using Newtonsoft.Json;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager.Model {
	/// <summary>
	/// 予定情報を格納するクラスです。
	/// </summary>
	[JsonObject(nameof(PlanModel))]
	public class PlanModel {
		#region フィールド

		#endregion

		#region コンストラクタ

		public PlanModel() {
		}

		/// <summary>
		/// 予定情報の新しいインスタンスを生成します。
		/// </summary>
		/// <param name="subject">件名</param>
		/// <param name="start">開始時刻</param>
		/// <param name="end">終了時刻</param>
		public PlanModel(string subject, DateTime start, DateTime end) {
			this.Subject = subject;
			this.Start = start;
			this.End = end;

			if (this.Interval.Ticks < 0) {
				throw new ArgumentException($"終了時刻が開始時刻よりも前に設定されています。", $"{nameof(end)}");
			}
		}

		/// <summary>
		/// 予定情報の新しいインスタンスを生成します。
		/// </summary>
		/// <param name="subject">件名</param>
		/// <param name="start">開始時刻</param>
		/// <param name="interval">間隔</param>
		public PlanModel(string subject, DateTime start, TimeSpan interval) {
			this.Subject = subject;
			this.Start = start;
			this.End = start + interval;
		}

		/// <summary>
		/// 予定情報の新しいインスタンスを生成します。
		/// </summary>
		/// <param name="subject">件名</param>
		/// <param name="start">開始時刻</param>
		/// <param name="ticks">100 ナノ秒単位で表される期間</param>
		public PlanModel(string subject, DateTime start, long ticks) : this(subject, start, new TimeSpan(ticks)) {
		}

		/// <summary>
		/// 予定情報の新しいインスタンスを生成します。
		/// </summary>
		/// <param name="subject">件名</param>
		/// <param name="start">開始時刻</param>
		/// <param name="hours">時間数</param>
		/// <param name="minutes">分数</param>
		/// <param name="seconds">秒数</param>
		public PlanModel(string subject, DateTime start, int hours, int minutes, int seconds = 0) : this(subject, start, new TimeSpan(hours, minutes, seconds)) {
		}

		/// <summary>
		/// 予定情報(終日)の新しいインスタンスを生成します。
		/// </summary>
		/// <param name="subject">件名</param>
		/// <param name="start">開始時刻</param>
		public PlanModel(string subject, DateTime start) : this(subject, start.Date, 24, 0) {
		}

		#endregion

		#region プロパティ

		/// <summary>
		/// 件名
		/// </summary>
		public string Subject { get; set; }

		/// <summary>
		/// 本文
		/// </summary>
		public string Body { get; set; }

		/// <summary>
		/// 場所
		/// </summary>
		public string Location { get; set; }

		/// <summary>
		/// 期間
		/// </summary>
		[JsonIgnore]
		public Ews.TimeWindow Period => new Ews.TimeWindow(this.Start, this.End);

		/// <summary>
		/// 開始時刻
		/// </summary>
		public DateTime Start { get; set; }

		/// <summary>
		/// 終了時刻
		/// </summary>
		public DateTime End { get; set; }

		/// <summary>
		/// 時間間隔
		/// </summary>
		[JsonIgnore]
		public TimeSpan Interval => this.End - this.Start;

		#endregion

		#region メソッド

		/// <summary>
		/// JSON 形式の文字列に変換します。
		/// </summary>
		/// <returns>JSON 形式の文字列を返します。</returns>
		public string ToJson()
			=> this.ToJson(true);

		#endregion
	}
}
