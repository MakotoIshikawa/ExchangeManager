using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace ExchangeManager.Extensions {
	/// <summary>
	/// JSON 変換を拡張するメソッドを提供します。
	/// </summary>
	public static partial class JsonConvertExtension {
		/// <summary>
		/// JSON 形式の文字列に変換します。
		/// </summary>
		/// <typeparam name="T">オブジェクトの型</typeparam>
		/// <param name="this">オブジェクト</param>
		/// <param name="ignoreNull">null 値を無視するかどうか</param>
		/// <returns>JSON 形式の文字列を返します。</returns>
		public static string ToJson<T>(this T @this, bool ignoreNull = true)
			=> JsonConvert.SerializeObject(@this, new JsonSerializerSettings {
				NullValueHandling = ignoreNull ? NullValueHandling.Ignore : NullValueHandling.Include,
			});

		/// <summary>
		/// JSON 形式の文字列をオブジェクトに変換します。
		/// </summary>
		/// <typeparam name="T">オブジェクトの型</typeparam>
		/// <param name="this">文字列</param>
		/// <param name="ignoreNull">null 値を無視するかどうか</param>
		/// <returns>変換したオブジェクトを返します。</returns>
		public static T Deserialize<T>(this string @this, bool ignoreNull = true)
			=> JsonConvert.DeserializeObject<T>(@this, new JsonSerializerSettings {
				NullValueHandling = ignoreNull ? NullValueHandling.Ignore : NullValueHandling.Include,
			});
	}
}
