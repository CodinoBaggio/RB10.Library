using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RB10.Library.Extensions
{
    public static class NPOIExtensions
    {
        /// <summary>
        /// セル値を文字列形式で取得します。
        /// 計算式のセルは計算結果を取得します。
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static string GetStringCellValue(this ICell cell)
        {
            string cellStr = string.Empty;
            switch (cell.CellType)
            {
                // 文字列型
                case CellType.String:
                    cellStr = cell.StringCellValue;
                    break;
                // 数値型（日付の場合もここに入る）
                case CellType.Numeric:
                    // セルが日付情報が単なる数値かを判定
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        // 日付型
                        // 本来はスタイルに合わせてフォーマットすべきだが、
                        // うまく表示できないケースが若干見られたので固定のフォーマットとして取得
                        cellStr = cell.DateCellValue.ToString("yyyy/MM/dd HH:mm:ss");
                    }
                    else
                    {
                        // 数値型
                        cellStr = cell.NumericCellValue.ToString();
                    }
                    break;

                // bool型(文字列でTrueとか入れておけばbool型として扱われた)
                case CellType.Boolean:
                    cellStr = cell.BooleanCellValue.ToString();
                    break;

                // 入力なし
                case CellType.Blank:
                    cellStr = cell.ToString();
                    break;

                // 数式
                case CellType.Formula:
                    // 下記で数式の文字列が取得される
                    //cellStr = cell.CellFormula.ToString();

                    // 数式の元となったセルの型を取得して同様の処理を行う
                    // コメントは省略
                    switch (cell.CachedFormulaResultType)
                    {
                        case CellType.String:
                            cellStr = cell.StringCellValue;
                            break;
                        case CellType.Numeric:

                            if (DateUtil.IsCellDateFormatted(cell))
                            {
                                cellStr = cell.DateCellValue.ToString("yyyy/MM/dd HH:mm:ss");
                            }
                            else
                            {
                                cellStr = cell.NumericCellValue.ToString();
                            }
                            break;
                        case CellType.Boolean:
                            cellStr = cell.BooleanCellValue.ToString();
                            break;
                        case CellType.Blank:
                            break;
                        case CellType.Error:
                            cellStr = cell.ErrorCellValue.ToString();
                            break;
                        case CellType.Unknown:
                            break;
                        default:
                            break;
                    }
                    break;

                // エラー
                case CellType.Error:
                    cellStr = cell.ErrorCellValue.ToString();
                    break;

                // 型不明なセル
                case CellType.Unknown:
                    break;
                // もっと不明なセル
                default:
                    break;
            }

            return cellStr;
        }

        /// <summary>
        /// シート内の使用範囲を取得します。
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static string[,] UsedRange(this ISheet sheet)
        {
            // 使用している最大列のインデックスを求める
            int maxCol = 0;
            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null)
                {
                    continue;
                }
                else
                {
                    if (maxCol < row.LastCellNum - 1) maxCol = row.LastCellNum - 1;
                }
            }

            // シート内容を取得
            List<List<string>> rowList = new List<List<string>>();
            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i) ?? sheet.CreateRow(i);

                List<string> cells = new List<string>();
                for (int j = 0; j <= maxCol; j++)
                {
                    var cell = row.GetCell(j) ?? row.CreateCell(j);
                    cells.Add(cell.GetStringCellValue());
                }
                rowList.Add(cells);
            }

            string[,] usedRange = new string[rowList.Count, rowList.Max(x => x.Count)];
            for (int i = 0; i < rowList.Count; i++)
            {
                for (int j = 0; j < rowList[i].Count; j++)
                {
                    usedRange[i, j] = rowList[i][j];
                }
            }

            return usedRange;
        }
    }
}
