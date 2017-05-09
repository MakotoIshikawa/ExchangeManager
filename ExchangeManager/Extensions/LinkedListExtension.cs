using System;
using System.Collections.Generic;
using System.Linq;

namespace ExchangeManager.Extensions {
	/// <summary>
	/// LinkedList を拡張するメソッドを提供します。
	/// </summary>
	public static partial class LinkedListExtension {
		#region メソッド

		/// <summary>
		/// LinkedListNode を列挙します。
		/// </summary>
		/// <typeparam name="T">要素の型</typeparam>
		/// <param name="this">LinkedList</param>
		/// <returns>LinkedListNode の列挙を返します。</returns>
		private static IEnumerable<LinkedListNode<T>> ToEnumerable<T>(this LinkedList<T> @this) {
			for (var node = @this.First; node != null; node = node.Next) {
				yield return node;
			}
		}

		/// <summary>
		/// 各要素に対して、指定された処理を実行します。
		/// </summary>
		/// <typeparam name="T">要素の型</typeparam>
		/// <param name="this">LinkedList</param>
		/// <param name="action">各要素に対して実行するメソッド</param>
		public static void ForEach<T>(this LinkedList<T> @this, Action<LinkedListNode<T>> action)
			=> @this.ToEnumerable().ToList().ForEach(action);

		/// <summary>
		/// LinkedList の各要素を新しいフォームに射影します。
		/// </summary>
		/// <typeparam name="TSource">source の要素の型</typeparam>
		/// <typeparam name="TResult">selector によって返される値の型</typeparam>
		/// <param name="this">LinkedList</param>
		/// <param name="selector">各要素に適用する変換関数。</param>
		/// <returns>source の各要素に対して変換関数を呼び出した結果として得られる要素を含む列挙を返します。</returns>
		public static IEnumerable<TResult> Select<TSource, TResult>(this LinkedList<TSource> @this, Func<LinkedListNode<TSource>, TResult> selector)
			=> @this.ToEnumerable().Select(selector);

		/// <summary>
		/// 述語に基づいて値のシーケンスをフィルター処理します。
		/// </summary>
		/// <typeparam name="TSource">source の要素の型</typeparam>
		/// <param name="this">LinkedList</param>
		/// <param name="predicate">各要素が条件を満たしているかどうかをテストする関数。</param>
		/// <returns>条件を満たす、入力シーケンスの要素を含む列挙を返します。</returns>
		public static IEnumerable<LinkedListNode<TSource>> Where<TSource>(this LinkedList<TSource> @this, Func<LinkedListNode<TSource>, bool> predicate)
			=> @this.ToEnumerable().Where(predicate);

		/// <summary>
		/// 現在の値と次の値のペアの列挙に変換します。
		/// </summary>
		/// <typeparam name="T">要素の型</typeparam>
		/// <param name="this">LinkedList</param>
		/// <returns>現在の値と次の値のペアの列挙を返します。</returns>
		public static IEnumerable<Tuple<T, T>> ToCurrentAndNextPair<T>(this LinkedList<T> @this)
			=> @this
			.Where(node => node.Next != null)
			.Select(node => Tuple.Create(node.Value, node.Next.Value));

		#endregion
	}
}
