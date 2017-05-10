using ExchangeManager.Interface;
using ExchangeManager.Primitives;
using Ews = Microsoft.Exchange.WebServices.Data;

namespace ExchangeManager {
	/// <summary>
	/// オンプレミスの Exchange Server の情報を管理するクラスです。
	/// </summary>
	public class ExchangeServerManager : ExchangeManagerBase, IExchangeManager {
		#region フィールド

		#endregion

		#region コンストラクタ

		/// <summary>
		/// コンストラクタ
		/// </summary>
		public ExchangeServerManager() : base() {
			this.Service = this.CreateService();
		}

		#endregion

		#region プロパティ

		#endregion

		#region メソッド

		/// <summary>
		/// EWS の新しいインスタンスを生成します。
		/// </summary>
		/// <returns>生成した EWS のインスタンスを返します。</returns>
		protected override Ews.ExchangeService CreateService()
			=> new Ews.ExchangeService(Ews.ExchangeVersion.Exchange2013_SP1) {
				UseDefaultCredentials = true,
				TraceEnabled = true,
				TraceFlags = Ews.TraceFlags.All,
			};

		#endregion
	}
}
