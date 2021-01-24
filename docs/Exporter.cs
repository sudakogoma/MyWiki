using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Controls;
using System.Windows.Data;

namespace
{
    /// <summary>
    /// <see cref="DataGridView"/> が持つ全レコードを表示値でCSV出力するクラス
    /// </summary>
    internal class XXX
    {
        /// <summary>
        /// 出力文字コード
        /// </summary>
        private readonly Encoding _encoding;

        /// <summary>
        /// 区切り文字
        /// </summary>
        private readonly string _separator;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="encoding">
        /// 文字コード。 <para/>
        /// 未指定の場合、UTF8(BOMあり)となります。
        /// </param>
        /// <param name="separator">
        /// 区切り文字。<para/>
        /// 未指定の場合、現在のカルチャに関連付けられている区切り文字となります。
        /// </param>
        public XXX(Encoding encoding = null, string separator = null)
        {
            _encoding = encoding ?? new UTF8Encoding(true);
            _separator = string.IsNullOrEmpty(separator)
                ? CultureInfo.CurrentCulture.TextInfo.ListSeparator
                : separator;
        }

        /// <summary>
        /// CSV出力を行います。
        /// </summary>
        /// <param name="grid">対象 <see cref="DataGridView"/></param>
        /// <param name="filePath">出力先ファイル</param>
        public void Export(DataGridView grid, string filePath)
        {
            using (var writer = new StreamWriter(filePath, false, _encoding))
            {
                writer.WriteLine(CreateHeaderLine(grid.Columns));

                foreach (var line in CreateBodyLines(grid.CollectionView, grid.Columns))
                {
                    writer.WriteLine(line);
                }
            }
        }

        /// <summary>
        /// CSV出力用にエスケープ等のフォーマット処理を行った文字列に変換します。
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private string ConvertFormattedTextForCsv(string text)
        {
            // " のエスケープを行い、" で囲う
            return $"\"{text.Replace("\"", "\"\"")}\"";
        }

        /// <summary>
        /// 列ヘッダ部分の文字列を生成します。
        /// </summary>
        /// <param name="columns"></param>
        /// <returns></returns>
        private string CreateHeaderLine(IEnumerable<DataGridViewColumn> columns)
        {
            // 列ヘッダからテキストデータを取得するローカル関数
            string getColumnHeaderText(DataGridViewColumn c)
            {
                if (c.HeaderTemplate != null)
                {
                    switch (c.HeaderTemplate.LoadContent())
                    {
                        case TextBlock textBlock:
                            return textBlock.Text;
                        default:
                            // 新たに追加した HeaderTemplate の構造に合わせて処理を追加してください。
                            throw new NotImplementedException("This HeaderTemplate is not supported.");
                    }
                }

                return string.IsNullOrEmpty(c.Header)
                    ? c.Property
                    : c.Header;
            }

            var headers = columns
                .Select(x => getColumnHeaderText(x))
                .Select(x => ConvertFormattedTextForCsv(x));
            return string.Join(_separator, headers);
        }

        /// <summary>
        /// データ本体部分の文字列一覧を生成します。
        /// </summary>
        /// <param name="collectionView"></param>
        /// <param name="columns"></param>
        /// <returns></returns>
        private IEnumerable<string> CreateBodyLines(ICollectionView collectionView, IEnumerable<DataGridViewColumn> columns)
        {
            // 列フィルタを無視し全ての行を対象にするために CollectionView を新たに作成する
            var notFilteredCollectionView = new CollectionViewSource
            {
                Source = collectionView.SourceCollection
            };
            notFilteredCollectionView.SortDescriptions.AddRange(collectionView.SortDescriptions);

            foreach (dynamic item in notFilteredCollectionView.View)
            {
                var line = columns
                    .Select(x => x.PropertyInfo.GetValue(item)?.ToString() ?? "???")
                    .Select(x => ConvertFormattedTextForCsv(x));
                yield return string.Join(_separator, line);
            }
        }
    }
}
