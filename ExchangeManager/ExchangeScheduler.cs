using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using ExchangeManager.Interface;
using ExtensionsLibrary.Extensions;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager {
	/// <summary>
	/// Exchange の予定表を
	/// </summary>
	/// <typeparam name="TExchangeManager">IExchangeManager を実装する型</typeparam>
	public class ExchangeScheduler {
		#region フィールド

		private Ews.AvailabilityOptions _options = new Ews.AvailabilityOptions();

		#endregion

		#region コンストラクタ

		public ExchangeScheduler(IExchangeManager manager, DateTime startTime, DateTime endTime, IEnumerable<Ews.AttendeeInfo> attendees) {
			this.Manager = manager;
			this.DetailedSuggestionsWindow = new Ews.TimeWindow(startTime, endTime);
			this.Attendees = attendees;
		}

		#endregion

		#region プロパティ

		/// <summary>
		/// Exchange を管理するオブジェクトを取得します。
		/// </summary>
		protected IExchangeManager Manager { get; set; }

		/// <summary>
		/// 出席者のコレクションを取得します。
		/// </summary>
		public IEnumerable<Ews.AttendeeInfo> Attendees { get; protected set; }

		/// <summary>
		/// GetUserAvailability メソッドを介して要求できるデータの種類を定義します。
		/// <para>デフォルトは [FreeBusyAndSuggestions] です。</para>
		/// </summary>
		/// <remarks>
		/// <para>[FreeBusy] : 空き時間情報のみを返します。</para>
		/// <para>[Suggestions] : 提案のみを返します。</para>
		/// <para>[FreeBusyAndSuggestions] : 空き時間情報と提案情報の両方を返します。</para>
		/// </remarks>
		public Ews.AvailabilityData RequestedData { get; set; } = Ews.AvailabilityData.FreeBusyAndSuggestions;

		#region AvailabilityOptions

		/// <summary>
		/// 推奨される会議時間を使用して更新する会議の開始時刻を取得または設定します。
		/// </summary>
		public DateTime? CurrentMeetingTime {
			get { return this._options.CurrentMeetingTime; }
			set { this._options.CurrentMeetingTime = value; }
		}

		/// <summary>
		/// 推奨される会議時間に関する詳細情報が返される時間ウィンドウを取得または設定します。
		/// </summary>
		public Ews.TimeWindow DetailedSuggestionsWindow {
			get { return this._options.DetailedSuggestionsWindow; }
			set { this._options.DetailedSuggestionsWindow = value; }
		}

		/// <summary>
		/// 推奨される会議時間に関する開始日時を取得または設定します。
		/// </summary>
		public DateTime StartTime {
			get { return this.DetailedSuggestionsWindow.StartTime; }
			set { this.DetailedSuggestionsWindow.StartTime = value; }
		}

		/// <summary>
		/// 推奨される会議時間に関する終了日時を取得または設定します。
		/// </summary>
		public DateTime EndTime {
			get { return this.DetailedSuggestionsWindow.EndTime; }
			set { this.DetailedSuggestionsWindow.EndTime = value; }
		}

		/// <summary>
		/// GetUserAvailability メソッドによって返されたデータに基づいて変更される会議のグローバルオブジェクトIDを取得または設定します。
		/// </summary>
		public string GlobalObjectId {
			get { return this._options.GlobalObjectId; }
			set { this._options.GlobalObjectId = value; }
		}

		/// <summary>
		/// 推奨される会議時間としての資格を得るために、その期間に期間を開いておく必要がある出席者の割合を取得または設定します。
		/// <para>値は 1～49 の間でなければなりません。</para>
		/// <para>デフォルト値は 25 です。</para>
		/// </summary>
		public int GoodSuggestionThreshold {
			get { return this._options.GoodSuggestionThreshold; }
			set { this._options.GoodSuggestionThreshold = value; }
		}

		/// <summary>
		/// 1日あたりの通常の営業時間外に推奨される会議時間の数を取得または設定します。
		/// <para>値は 0～48 の間でなければなりません。</para>
		/// <para>デフォルト値は 10 です。</para>
		/// </summary>
		public int MaximumNonWorkHoursSuggestionsPerDay {
			get { return this._options.MaximumNonWorkHoursSuggestionsPerDay; }
			set { this._options.MaximumNonWorkHoursSuggestionsPerDay = value; }
		}

		/// <summary>
		/// 1日に返される推奨会議時間の数を取得または設定します。
		/// <para>値は 0～48 の間でなければなりません。</para>
		/// <para>デフォルト値は 10 です。</para>
		/// </summary>
		public int MaximumSuggestionsPerDay {
			get { return this._options.MaximumSuggestionsPerDay; }
			set { this._options.MaximumSuggestionsPerDay = value; }
		}

		/// <summary>
		/// 提案を取得する会議の期間を取得または設定します。
		/// <para>値は 30～1440 の間でなければなりません。</para>
		/// <para>デフォルト値は 60 です。</para>
		/// </summary>
		public int MeetingDuration {
			get { return this._options.MeetingDuration; }
			set { this._options.MeetingDuration = value; }
		}

		/// <summary>
		/// [FreeBusyMerged] ビュー内の連続する2つのスロット間の時間差を取得または設定します。
		/// <para>値は 5～1440 の間でなければなりません。</para>
		/// <para>デフォルト値は 30 です。</para>
		/// </summary>
		public int MergedFreeBusyInterval {
			get { return this._options.MergedFreeBusyInterval; }
			set { this._options.MergedFreeBusyInterval = value; }
		}

		/// <summary>
		/// 返される提案の最小品質を取得または設定します。
		/// <para>デフォルトは [Fair] です。</para>
		/// </summary>
		public Ews.SuggestionQuality MinimumSuggestionQuality {
			get { return this._options.MinimumSuggestionQuality; }
			set { this._options.MinimumSuggestionQuality = value; }
		}

		/// <summary>
		/// 要求された空き時間ビューの種類を取得または設定します。
		/// <para>デフォルト値は [Detailed] です。</para>
		/// </summary>
		/// <remarks>
		/// <para>[None] : ビューは返されませんでした。
		/// GetUserAvailability メソッドの呼び出しでこの値を指定することはできません。</para>
		/// <para>[MergedOnly] : 集約された空き時間情報ストリームを表します。
		/// 1つのフォレストの対象ユーザーに可用性サービスが構成されていないフォレスト間のシナリオでは、
		/// リクエスタの可用性サービスは空き時間情報パブリックフォルダから対象ユーザーの空き時間情報を取得します。
		/// パブリックフォルダは空き時間情報のみをマージ形式で保存するため、[MergedOnly] は利用可能な唯一の情報です。</para>
		/// <para>[FreeBusy] : 従来のステータス情報（空き、ビジー、暫定、OOF）を表します。
		/// これには、予定の開始時刻と終了時刻も含まれます。
		/// 集約された空き時間情報ストリームの代わりに個々の会議の開始時間と終了時間が提供されるため、
		/// このビューは従来の空き時間ビューよりも豊富です。</para>
		/// <para>[FreeBusyMerged] : [FreeBusy] のすべてのプロパティを結合し、空き時間情報を結合したストリームを表します。</para>
		/// <para>[Detailed] : レガシーステータス情報を表します。空き、ビジー、暫定、およびOOF。
		/// 予定の開始/終了時刻。主語、場所、重要性などの任命のさまざまな特性。
		/// この要求されたビューは、要求しているユーザーが特権を持つ情報の最大量を返します。
		/// マージされた空き時間情報のみが利用可能な場合、
		/// Microsoft Exchange Server 2003フォレスト内のユーザーの情報を要求する場合と同様に、
		/// MergedOnlyが返されます。
		/// それ以外の場合は、FreeBusyまたはDetailedが返されます。</para>
		/// <para>[DetailedMerged] : マージされた空き時間情報のストリームと共に、詳細のすべてのプロパティを表します。
		/// 例えば、Exchange 2003を実行しているコンピュータにメールボックスが存在する場合など、
		/// 空き時間情報が1つのみ結合されている場合、[MergedOnly] が返されます。
		/// それ以外の場合は、[FreeBusyMerged] または [DetailedMerged] が返されます。</para>
		/// </remarks>
		public Ews.FreeBusyViewType RequestedFreeBusyView {
			get { return this._options.RequestedFreeBusyView; }
			set { this._options.RequestedFreeBusyView = value; }
		}

		#endregion

		#endregion

		#region メソッド

		#region 空き時間確認

		/// <summary>
		/// 空き時間を取得します。
		/// </summary>
		/// <param name="startTime">開始時間</param>
		/// <param name="endTime">終了時間</param>
		/// <param name="goodSuggestionThreshold">推奨される会議時間としての資格を得るために、その期間に期間を開いておく必要がある出席者の割合を取得または設定します。1～49でなければなりません。デフォルト値は25です。</param>
		/// <param name="maximumNonWorkHoursSuggestionsPerDay">1日あたりの通常の営業時間外に推奨される会議時間の数を取得または設定します。0～48の間でなければなりません。デフォルト値は0です。</param>
		/// <param name="maximumSuggestionsPerDay">1日に返される推奨会議時間の数を取得または設定します。0～48の間でなければなりません。デフォルト値は10です。</param>
		/// <param name="meetingDuration">提案を取得する会議の所要時間を分単位で取得または設定します。30～1440でなければなりません。既定値は60です。</param>
		/// <returns>空き時間の情報を返します。</returns>
		protected virtual Ews.GetUserAvailabilityResults GetUserAvailability(DateTime startTime, DateTime endTime, int goodSuggestionThreshold = 25, int maximumNonWorkHoursSuggestionsPerDay = 0, int maximumSuggestionsPerDay = 10, int meetingDuration = 60)
			=> this.Manager.GetUserAvailability(this.Attendees, startTime, endTime, goodSuggestionThreshold, maximumNonWorkHoursSuggestionsPerDay, maximumSuggestionsPerDay, meetingDuration);

		protected virtual async Task<Ews.GetUserAvailabilityResults> GetUserAvailabilityAsync(DateTime startTime, DateTime endTime, int goodSuggestionThreshold = 25, int maximumNonWorkHoursSuggestionsPerDay = 0, int maximumSuggestionsPerDay = 10, int meetingDuration = 60)
			=> await this.Manager.GetUserAvailabilityAsync(this.Attendees, startTime, endTime, goodSuggestionThreshold, maximumNonWorkHoursSuggestionsPerDay, maximumSuggestionsPerDay, meetingDuration);

		protected virtual Ews.GetUserAvailabilityResults GetUserAvailability()
			=> this.Manager.GetUserAvailability(this.Attendees, this._options, this.RequestedData);

		protected virtual async Task<Ews.GetUserAvailabilityResults> GetUserAvailabilityAsync()
			=> await this.Manager.GetUserAvailabilityAsync(this.Attendees, this._options, this.RequestedData);

		#endregion

		/// <summary>
		/// 指定した期間の推奨会議時間のコレクションを取得します。
		/// </summary>
		/// <param name="startTime">開始時間</param>
		/// <param name="endTime">終了時間</param>
		/// <param name="goodSuggestionThreshold">推奨される会議時間としての資格を得るために、その期間に期間を開いておく必要がある出席者の割合を取得または設定します。1～49でなければなりません。デフォルト値は25です。</param>
		/// <param name="maximumNonWorkHoursSuggestionsPerDay">1日あたりの通常の営業時間外に推奨される会議時間の数を取得または設定します。0～48の間でなければなりません。デフォルト値は0です。</param>
		/// <param name="maximumSuggestionsPerDay">1日に返される推奨会議時間の数を取得または設定します。0～48の間でなければなりません。デフォルト値は10です。</param>
		/// <param name="meetingDuration">提案を取得する会議の所要時間を分単位で取得または設定します。30～1440でなければなりません。既定値は60です。</param>
		/// <returns>推奨会議時間のコレクションを返します。</returns>
		public Collection<Ews.Suggestion> GetSuggestions(DateTime startTime, DateTime endTime, int goodSuggestionThreshold = 25, int maximumNonWorkHoursSuggestionsPerDay = 0, int maximumSuggestionsPerDay = 10, int meetingDuration = 60)
			=> this.GetUserAvailability(startTime, endTime, goodSuggestionThreshold, maximumNonWorkHoursSuggestionsPerDay, maximumSuggestionsPerDay, meetingDuration)?.Suggestions;

		public Collection<Ews.Suggestion> GetSuggestions()
			=> this.GetUserAvailability()?.Suggestions;

		/// <summary>
		/// 出席者のカレンダーイベントのコレクションを取得します。
		/// </summary>
		/// <param name="startTime">開始時間</param>
		/// <param name="endTime">終了時間</param>
		/// <param name="goodSuggestionThreshold">推奨される会議時間としての資格を得るために、その期間に期間を開いておく必要がある出席者の割合を取得または設定します。1～49でなければなりません。デフォルト値は25です。</param>
		/// <param name="maximumNonWorkHoursSuggestionsPerDay">1日あたりの通常の営業時間外に推奨される会議時間の数を取得または設定します。0～48の間でなければなりません。デフォルト値は0です。</param>
		/// <param name="maximumSuggestionsPerDay">1日に返される推奨会議時間の数を取得または設定します。0～48の間でなければなりません。デフォルト値は10です。</param>
		/// <param name="meetingDuration">提案を取得する会議の所要時間を分単位で取得または設定します。30～1440でなければなりません。既定値は60です。</param>
		/// <returns></returns>
		public Dictionary<string, Collection<Ews.CalendarEvent>> GetUserAvailabilities(DateTime startTime, DateTime endTime, int goodSuggestionThreshold = 25, int maximumNonWorkHoursSuggestionsPerDay = 0, int maximumSuggestionsPerDay = 10, int meetingDuration = 60) {
			var results = GetUserAvailability(startTime, endTime, goodSuggestionThreshold, maximumNonWorkHoursSuggestionsPerDay, maximumSuggestionsPerDay, meetingDuration);
			return this.Attendees.Zip(results.AttendeesAvailability, (at, av) => new {
				Attendee = at,
				Availability = av,
			}).ToDictionary(a => a.Attendee.SmtpAddress, a => a.Availability.CalendarEvents);
		}

		public Dictionary<string, Collection<Ews.CalendarEvent>> GetUserAvailabilities() {
			var results = this.GetUserAvailability();
			return this.Attendees.Zip(results.AttendeesAvailability, (at, av) => new {
				Attendee = at,
				Availability = av,
			}).ToDictionary(a => a.Attendee.SmtpAddress, a => a.Availability.CalendarEvents);
		}

		public virtual async Task<Dictionary<string, Collection<Ews.CalendarEvent>>> GetUserAvailabilitiesAsync() {
			var results = await this.GetUserAvailabilityAsync();
			return this.Attendees.Zip(results.AttendeesAvailability, (at, av) => new {
				Attendee = at,
				Availability = av,
			}).ToDictionary(a => a.Attendee.SmtpAddress, a => a.Availability.CalendarEvents);
		}

		#endregion
	}
}
