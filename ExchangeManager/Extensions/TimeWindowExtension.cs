using System;
using System.Collections.Generic;
using System.Linq;
using ExtensionsLibrary.Extensions;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager.Extensions {
	/// <summary>
	/// TimeWindow を拡張するメソッドを提供します。
	/// </summary>
	public static partial class TimeWindowExtension {
		#region メソッド

		#region IsWithinRange

		/// <summary>
		/// 指定した期間の範囲内かどうかを判定します。
		/// </summary>
		/// <param name="this">TimeWindow</param>
		/// <param name="value">期間</param>
		/// <returns>範囲内であれば true。それ以外は false を返します。</returns>
		public static bool IsWithinRange(this Ews.TimeWindow @this, Ews.CalendarEvent value)
			=> value.StartTime <= @this.StartTime && @this.StartTime < value.EndTime
			|| value.StartTime < @this.EndTime && @this.EndTime <= value.EndTime;

		/// <summary>
		/// 指定した期間の範囲内かどうかを判定します。
		/// </summary>
		/// <param name="this">TimeWindow</param>
		/// <param name="value">期間</param>
		/// <returns>範囲内であれば true。それ以外は false を返します。</returns>
		public static bool IsWithinRange(this Ews.TimeWindow @this, Ews.TimeWindow value)
			=> value.StartTime <= @this.StartTime && @this.StartTime < value.EndTime
			|| value.StartTime < @this.EndTime && @this.EndTime <= value.EndTime;

		#endregion

		#region GetBlankPlans

		/// <summary>
		/// 指定した時間帯と時間間隔で
		/// 空の予定を列挙します。
		/// </summary>
		/// <param name="this">日付</param>
		/// <param name="openingTime">開業時刻</param>
		/// <param name="closingTime">終業時刻</param>
		/// <param name="intervalPerMinutes">分刻みの間隔</param>
		/// <returns>空の予定の列挙を返します。</returns>
		public static IEnumerable<Ews.TimeWindow> GetBlankPlans(this DateTime @this, double openingTime, double closingTime, int intervalPerMinutes) {
			var ret = @this.GetBlankPlansPerMinutes(openingTime, closingTime, intervalPerMinutes);
			var blankPlans = ret.ToCurrentAndNextPair()
				.Select(cn => new Ews.TimeWindow(cn.Item1, cn.Item2));
			return blankPlans;
		}

		private static LinkedList<DateTime> GetBlankPlansPerMinutes(this DateTime @this, double openingTime, double closingTime, int intervalPerMinutes) {
			var today = @this.Date;
			var b = new TimeSpan(Convert.ToInt64(TimeSpan.TicksPerHour * openingTime));
			var f = new TimeSpan(Convert.ToInt64(TimeSpan.TicksPerHour * closingTime));
			var ret = GetTimeSpans(b, f, intervalPerMinutes).Select(t => today + t).ToList();
			return new LinkedList<DateTime>(ret);
		}

		private static IEnumerable<TimeSpan> GetTimeSpans(TimeSpan open, TimeSpan close, int intervalPerMinutes) {
			if (open > close) {
				throw new ArgumentException($"終了時刻が、開始時刻よりも前に設定されています。: {close}", nameof(close));
			}

			var tt = new TimeSpan(TimeSpan.TicksPerMinute * intervalPerMinutes);
			for (var m = open; m <= close; m += tt) {
				yield return m;
			}
		}

		#endregion

		#region GetBlankTimes

		/// <summary>
		/// 予定イベントの列挙から、予定のない時間帯の列挙を取得します。
		/// </summary>
		/// <param name="this">CalendarEvent の列挙</param>
		/// <param name="openingTime">開業時刻</param>
		/// <param name="closingTime">終業時刻</param>
		/// <param name="intervalPerMinutes">分刻みの間隔</param>
		/// <returns>予定のない時間帯の列挙を返します。</returns>
		public static IEnumerable<Ews.TimeWindow> GetBlankTimes(this IEnumerable<Ews.CalendarEvent> @this, double openingTime, double closingTime, int intervalPerMinutes) {
			var tws = @this.Select(ev => new Ews.TimeWindow(ev.StartTime, ev.EndTime));
			return tws.GetBlankTimes(openingTime, closingTime, intervalPerMinutes);
		}

		/// <summary>
		/// 時間帯の列挙から、予定のない時間帯の列挙を取得します。
		/// </summary>
		/// <param name="this">TimeWindow の列挙</param>
		/// <param name="openingTime">開業時刻</param>
		/// <param name="closingTime">終業時刻</param>
		/// <param name="intervalPerMinutes">分刻みの間隔</param>
		/// <returns>予定のない時間帯の列挙を返します。</returns>
		public static IEnumerable<Ews.TimeWindow> GetBlankTimes(this IEnumerable<Ews.TimeWindow> @this, double openingTime, double closingTime, int intervalPerMinutes) {
			var days = @this.Select(ev => ev.StartTime.Date).Distinct();

			var times = (
				from day in days
				let blankPlans = day.GetBlankPlans(openingTime, closingTime, intervalPerMinutes)
				let dayPlans = @this.Where(d => d.StartTime.Date == day)
				select blankPlans.Where(p => !dayPlans.Any(d => p.IsWithinRange(d)))
			).SelectMany(p => p);

			return times;
		}

		#endregion

		#endregion
	}
}
