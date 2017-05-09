using System;
using System.Collections.Generic;
using System.Linq;
using ExchangeManager.Extensions;
using ExtensionsLibrary.Extensions;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager.Extensions {
	/// <summary>
	/// TimeWindow を拡張するメソッドを提供します。
	/// </summary>
	public static partial class TimeWindowExtension {
		/// <summary>
		/// 指定した時間帯と時間間隔で
		/// 空の予定を列挙します。
		/// </summary>
		/// <param name="this">日付</param>
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

		public static IEnumerable<Ews.TimeWindow> GetBlankTimes(this List<Ews.CalendarEvent> @this, double openingTime, double closingTime, int intervalPerMinutes) {
			var days = @this.Select(ev => ev.StartTime.Date).Distinct();

			var times = (
				from day in days
				let blankPlans = day.GetBlankPlans(openingTime, closingTime, intervalPerMinutes)
				let dayPlans = @this.Where(d => d.StartTime.Date == day)
				select blankPlans.Where(p => !dayPlans.Any(d =>
					 d.StartTime <= p.StartTime && p.StartTime < d.EndTime
					 || d.StartTime < p.EndTime && p.EndTime <= d.EndTime
				))
			).SelectMany(p => p);

			return times;
		}
	}
}
