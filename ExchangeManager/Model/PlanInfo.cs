using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExchangeManager.Model {
	/// <summary>
	/// 予定情報を格納するクラスです。
	/// </summary>
	public class PlanInfo {
		#region コンストラクタ

		/// <summary>
		/// 予定情報の新しいインスタンスを生成します。
		/// </summary>
		/// <param name="subject">件名</param>
		/// <param name="start">開始時刻</param>
		/// <param name="end">終了時刻</param>
		public PlanInfo(string subject, DateTime start, DateTime end) {
			this.Duration = end - start;

			if (this.Duration.Ticks < 0) {
				throw new ArgumentException($"終了時刻が開始時刻よりも前に設定されています。", $"{nameof(end)}");
			}

			this.Subject = subject;
			this.Start = start;
			this.End = end;
		}

		/// <summary>
		/// 予定情報の新しいインスタンスを生成します。
		/// </summary>
		/// <param name="subject">件名</param>
		/// <param name="start">開始時刻</param>
		/// <param name="duration">期間</param>
		public PlanInfo(string subject, DateTime start, TimeSpan duration) {
			this.Duration = duration;

			this.Subject = subject;
			this.Start = start;
			this.End = start + duration;
		}

		/// <summary>
		/// 予定情報の新しいインスタンスを生成します。
		/// </summary>
		/// <param name="subject">件名</param>
		/// <param name="start">開始時刻</param>
		/// <param name="ticks">100 ナノ秒単位で表される期間</param>
		public PlanInfo(string subject, DateTime start, long ticks) : this(subject, start, new TimeSpan(ticks)) {
		}

		/// <summary>
		/// 予定情報の新しいインスタンスを生成します。
		/// </summary>
		/// <param name="subject">件名</param>
		/// <param name="start">開始時刻</param>
		/// <param name="hours">時間数</param>
		/// <param name="minutes">分数</param>
		/// <param name="seconds">秒数</param>
		/// <param name="milliseconds">ミリ秒数</param>
		public PlanInfo(string subject, DateTime start, int hours, int minutes, int seconds = 0) : this(subject, start, new TimeSpan(hours, minutes, seconds)) {
		}

		/// <summary>
		/// 予定情報(終日)の新しいインスタンスを生成します。
		/// </summary>
		/// <param name="subject">件名</param>
		/// <param name="start">開始時刻</param>
		public PlanInfo(string subject, DateTime start) : this(subject, start.Date, 24, 0) {
		}

		#endregion

		#region プロパティ

		/// <summary>
		/// 件名
		/// </summary>
		public string Subject { get; protected set; }

		/// <summary>
		/// 開始時刻
		/// </summary>
		public DateTime Start { get; protected set; }

		/// <summary>
		/// 終了時刻
		/// </summary>
		public DateTime End { get; protected set; }

		/// <summary>
		/// 期間
		/// </summary>
		public TimeSpan Duration { get; protected set; }

		#endregion

		#region メソッド

		#endregion
	}
}
