using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;

using System.Text.RegularExpressions;

using System.Windows.Forms;

using System.Text.Json;
using System.Text.Json.Nodes;

using ExcelDna.Integration;

using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace ExcelDnaTest
{
    //public static class Class1
    //{
    //    [ExcelFunction(Description = "My first .NET function")]
    //    public static string SayHello(string name)
    //    {
    //        return "Hello " + name;
    //    }
    //}

    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        private IRibbonUI ribbon;

        public void OnLoad(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public override string GetCustomUI(string RibbonID)
        {
            return @"
      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='OnLoad'>
      <ribbon>
        <tabs>
          <tab id='tab1' label='Sheet Renderer'>
            <group id='group1' label='Render'>
              <button id='button1' label='JSONファイル選択' size='large' imageMso='FileSave' onAction='OnSelectJsonFileButtonPressed'/>
              <button id='button2' label='Render' size='large' imageMso='TableDrawTable' onAction='OnRenderButtonPressed'/>
              <editBox id='fileNameBox' label='JSONファイル' sizeString='hoge\\20XX-XX-XX\\index' getText='GetFileName' onChange='OnTextChanged' />
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }

        string JsonFilePath { get; set; }

        static string GetFilePathWithoutExtension(string path)
        {
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(path);
            string directory = Path.GetDirectoryName(path);
            string newFilePath = Path.Combine(directory, fileNameWithoutExtension);

            return newFilePath;
        }

        // 指定されたファイルパスから最後の指定された数のフォルダ名を取得する
        static string GetLastFolders(string path, int count)
        {
            // ディレクトリ部分を取得
            string directoryPath = Path.GetDirectoryName(path);
            if (directoryPath == null)
            {
                return string.Empty;
            }

            // フォルダ名を分割
            string[] folders = directoryPath.Split(Path.DirectorySeparatorChar);

            // 取得するフォルダの数を調整
            int folderCount = Math.Min(count, folders.Length);

            // 最後の指定された数のフォルダ名を取得
            int startIndex = folders.Length - folderCount;
            string result = string.Join(Path.DirectorySeparatorChar.ToString(), folders.Skip(startIndex).ToArray());
            return result;
        }

        public string GetFileName(IRibbonControl control)
        {
            if (JsonFilePath == null)
            {
                return "";
            }
            int folderCount = 2; // 取得するフォルダの数を指定
            string result = GetLastFolders(JsonFilePath, folderCount);

            // ファイル名を取得して連結
            string fileNathWithoutExtension = Path.GetFileNameWithoutExtension(JsonFilePath);

            return Path.Combine(result, fileNathWithoutExtension);
        }
        public void OnTextChanged(IRibbonControl control, string text)
        {
            // ユーザーがテキストを変更した場合に元のテキストに戻す
            ribbon.InvalidateControl("fileNameBox");
        }

        public void OnSelectJsonFileButtonPressed(IRibbonControl control)
        {
            string selectedFile = OpenFile();
            if (selectedFile != null)
            {
                JsonFilePath = selectedFile;
                ribbon.InvalidateControl("fileNameBox");
            }
        }

        string OpenFile()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "JSONファイル (*.json)|*.json";
                openFileDialog.Title = "ファイルを選択してください";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }
                else
                {
                    return null;
                }
            }
        }

        static void InsertRowsAndCopyFormulas(Excel.Worksheet worksheet, int startRow, int rowCount)
        {
            // 行を挿入
            Excel.Range insertRange = worksheet.Rows[startRow].Resize[rowCount];
            insertRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

            // 挿入した行の上の行から数式をコピー
            Excel.Range sourceRange = worksheet.Rows[startRow - 1];
            Excel.Range destinationRange = worksheet.Rows[startRow].Resize[rowCount];
            sourceRange.Copy(destinationRange);
        }

        static void DeleteRows(Excel.Worksheet worksheet, int startRow, int rowCount)
        {
            Excel.Range rows = worksheet.Rows[startRow].Resize[rowCount];
            rows.Delete();
        }

        // JsonNode から指定した名前のオブジェクトの直下のプロパティをすべて Dictionary<string, string> として返却する
        static Dictionary<string, string> GetPropertiesFromJsonNode(JsonNode jsonNode, string objectName)
        {
            JsonNode objectNode = jsonNode[objectName];

            if (objectNode != null && objectNode is JsonObject jsonObject)
            {
                Dictionary<string, string> result = new Dictionary<string, string>();

                foreach (var property in jsonObject)
                {
                    if (property.Value is JsonValue value && value.TryGetValue(out string stringValue))
                    {
                        result[property.Key] = stringValue;
                    }
                }

                return result;
            }

            return null;
        }

        public void OnRenderButtonPressed(IRibbonControl control)
        {
            if (JsonFilePath == null)
            {
                MessageBox.Show("JSONファイルを指定してください。");
                return;
            }

            string jsonString = File.ReadAllText(JsonFilePath);

            JsonNode jsonObject = JsonNode.Parse(jsonString);

            const string templateFileName = "template.xlsm";
            string newFilePath = GetFilePathWithoutExtension(JsonFilePath);
            Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;

            string templateFilePath = FilePathHelper.GetAbsolutePathFromExecutingDirectory(templateFileName);

            Excel.Workbook workbook = CreateCopiedWorkbook(excelApp, templateFilePath, newFilePath);

            MacroControl.DisableMacros();
            excelApp.DisplayAlerts = false;
            excelApp.ScreenUpdating = false;

            const string indexSheetName = "表紙";
            const string templateSheetName = "単列チェック結果format";

            var confData = GetPropertiesFromJsonNode(jsonObject, "variables");

            // 特定のプロパティ（items）を配列としてアクセス
            JsonArray items = jsonObject["children"].AsArray();

            List<string> sheetNames = new List<string>();
            Excel.Worksheet sheet = workbook.Sheets[templateSheetName];

            foreach (JsonNode sheetNode in items)
            {
                string newSheetName = sheetNode["text"].ToString();

                // シートをコピーしてリネーム
                sheet.Copy(After: workbook.Sheets[workbook.Sheets.Count]);
                Excel.Worksheet newSheet = workbook.Sheets[workbook.Sheets.Count];
                newSheet.Name = newSheetName;

                RenderSheet(sheetNode, confData, newSheet);

                sheetNames.Add(newSheetName);
            }

            // index sheet にシート名を入力
            Excel.Worksheet indexSheet = workbook.Sheets[indexSheetName];
            int indexStartRow = 16;
            int indexEndRow = 35;
            int indexStartColumn = 2;
            int indexRowCount = indexEndRow - indexStartRow + 1;

            // 行が足りなければ挿入
            if (sheetNames.Count > indexRowCount)
            {
                int numberOfRows = sheetNames.Count - indexRowCount;

                InsertRowsAndCopyFormulas(indexSheet, indexEndRow, numberOfRows);
            }
            // 多ければ削除
            else if (sheetNames.Count < indexRowCount)
            {
                int numberOfRowsToDelete = indexRowCount - sheetNames.Count;

                DeleteRows(indexSheet, indexStartRow, numberOfRowsToDelete);
            }

            SetValueInSheetAsColumn(indexSheet, indexStartRow, indexStartColumn, sheetNames);

            // 幅をautofit
            indexSheet.Cells[indexStartRow, indexStartColumn].Resize(indexRowCount).Columns.AutoFit();

            // 最後にindexシートを選択状態にしておく
            indexSheet.Select();

            excelApp.ScreenUpdating = true;
            excelApp.DisplayAlerts = true;
            MacroControl.EnableMacros();

            workbook.Save();
        }

        // マクロを一時的に黙らせたい
        // XXX: うまく動いてない様子
        public class MacroControl
        {
            public static void DisableMacros()
            {
                Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
                excelApp.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            }

            public static void EnableMacros()
            {
                Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
                excelApp.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityByUI;
            }
        }

        static void TraverseTree(JsonNode node, List<List<string>> result, List<string> currentPath, int currentDepth, ref int maxDepth, ref int leafCount)
        {
            if (node == null) return;

            string text = node["text"]?.ToString();
            currentPath.Add(text);
            var nullifiedPath = Enumerable.Repeat<string>(null, currentPath.Count).ToList();

            JsonArray children = node["children"] as JsonArray;
            if (children != null && children.Count > 0)
            {
                for (int i = 0; i < children.Count; i++)
                {
                    JsonNode child = children[i];
                    var currentOrNullifiedPath = new List<string>(i == 0 ? currentPath : nullifiedPath);
                    TraverseTree(child, result, currentOrNullifiedPath, currentDepth + 1, ref maxDepth, ref leafCount);
                }
            }
            else
            {
                // 現在のパスを結果に追加
                result.Add(new List<string>(currentPath));
                leafCount++;
            }

            // 最大深度を更新
            if (currentDepth > maxDepth)
            {
                maxDepth = currentDepth;
            }
        }

        static List<List<string>> TraverseTreeFromRoot(JsonNode rootNode, out int leafCount, out int maxDepth)
        {
            List<List<string>> result = new List<List<string>>();
            maxDepth = 0;
            leafCount = 0;

            TraverseTree(rootNode, result, new List<string>(), 1, ref maxDepth, ref leafCount);

            return result;
        }

        static T[,] ConvertTo2DArray<T>(List<List<T>> list)
        {
            int rows = list.Count;
            int cols = rows > 0 ? list.Max(subList => subList.Count) : 0;
            T[,] array = new T[rows, cols];

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < list[i].Count; j++)
                {
                    array[i, j] = list[i][j];
                }
            }

            return array;
        }

        static void ReplaceTrailingNullsInLastColumn(string[,] array, string replacement)
        {
            int numRows = array.GetLength(0);
            int numCols = array.GetLength(1);

            for (int i = 0; i < numRows; i++)
            {
                // 後ろから連続するnullを置換文字列に置き換える
                for (int j = numCols - 1; j >= 0; j--)
                {
                    if (array[i, j] != null)
                    {
                        break; // nullでない要素が見つかったらループを終了
                    }
                    array[i, j] = replacement;
                }
            }
        }

        static string[,] RemoveFirstColumn(string[,] array)
        {
            int rows = array.GetLength(0);
            int cols = array.GetLength(1);

            string[,] result = new string[rows, cols - 1];

            for (int i = 0; i < rows; i++)
            {
                for (int j = 1; j < cols; j++)
                {
                    result[i, j - 1] = array[i, j];
                }
            }

            return result;
        }

        static List<List<string>> RemoveFirstColumn(List<List<string>> list)
        {
            List<List<string>> result = new List<List<string>>();

            foreach (var row in list)
            {
                List<string> newRow = new List<string>(row.Skip(1));
                result.Add(newRow);
            }

            return result;
        }

        static void RenderSheet(JsonNode sheetNode, Dictionary<string, string> confData, Excel.Worksheet sheet)
        {
            int leafCount;
            int maxDepth;
            List<List<string>> result = TraverseTreeFromRoot(sheetNode, out leafCount, out maxDepth);

            // 左端はシート名なので削除
            result = RemoveFirstColumn(result);
            maxDepth--;

            // 2次元配列に変換
            string[,] arrayResult = ConvertTo2DArray(result);

            // 右端にダミー文字追加
            ReplaceTrailingNullsInLastColumn(arrayResult, "---");

            int startColumn = 3;
            int endColumn = 6;
            int columnWidth = endColumn - startColumn + 1;
            int startRow = 25;
            int endRow = 102;
            int rowHeight = endRow - startRow + 1;

            int initialDateColumn = 13;
            int initialResultColumn = 7;

            if (maxDepth > columnWidth)
            {
                int numberOfColumns = maxDepth - columnWidth;
                Excel.Range startColumnRange = sheet.Columns[endColumn];

                // 列を挿入
                startColumnRange.Resize[1, numberOfColumns].EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            }

            if (leafCount > rowHeight)
            {
                int numberOfRows = leafCount - rowHeight;

                InsertRowsAndCopyFormulas(sheet, endRow, numberOfRows);
            }
            else if (leafCount < rowHeight)
            {
                int numberOfRowsToDelete = rowHeight - leafCount;

                DeleteRows(sheet, startRow, numberOfRowsToDelete);
            }

            // XXX: 深度が列より少ない場合は今のテンプレに合わせていろいろやる…
            if (maxDepth < columnWidth)
            {
                // とりあえず E 列は削除
                Excel.Range column = sheet.Columns[5];
                column.Delete();
                columnWidth--;

                // C, D 列を削除するのはいろいろ面倒な作りのようなので、
                // 貼り付け列を調整してお茶を濁す（C, D 列は空欄にする）
                startColumn += columnWidth - maxDepth;
            }

            SetValueInSheet(sheet, startRow, startColumn, arrayResult);

            // 「チェック予定日」の列に無条件で START_DATE を入れる
            int dateColumn = startColumn + maxDepth + (initialDateColumn - initialResultColumn);
            string dateString = confData["START_DATE"];
            // 文字列をDateTimeに変換
            if (DateTime.TryParse(dateString, out DateTime dateValue))
            {
                SetRangeValue(sheet, startRow, dateColumn, leafCount, 1, dateValue);
            }

            // XXX: 先頭が【*】なセルの対応
            int resultColumn = startColumn + maxDepth;
            for (int i = 0; i < result.Count; i++)
            {
                for (int j = 0; j < result[i].Count; j++)
                {
                    string text = result[i][j];

                    if (text == null)
                    {
                        continue;
                    }

                    const string pattern = @"^【.*】";
                    bool isMatch = Regex.IsMatch(text, pattern);

                    if (!isMatch)
                    {
                        continue;
                    }

                    // ここより右のセルの色を変える
                    const int cellThemeColorId = 5;   // 水色っぽい色
                    const int fontColorIndex = 2;   // 白
                    var cells = sheet.Cells[startRow + i, startColumn + j];
                    cells = cells.Resize(1, maxDepth - j);
                    cells.Interior.ThemeColor = cellThemeColorId;
                    cells.Font.ColorIndex = fontColorIndex;

                    // 結果欄に - を入力
                    var resultCell = sheet.Cells[startRow + i, resultColumn];
                    resultCell.Value = "-";

                    // チェック予定日欄を空欄にする
                    var dateCell = sheet.Cells[startRow + i, dateColumn];
                    dateCell.Value = null;

                    break;
                }
            }
        }

        // 指定した範囲のセルに同じ値を代入します
        static void SetRangeValue<T>(Excel.Worksheet sheet, int startRow, int startColumn, int rowCount, int columnCount, T value)
        {
            Excel.Range startCell = (Excel.Range)sheet.Cells[startRow, startColumn];
            Excel.Range range = startCell.Resize[rowCount, columnCount];

            range.Value = value;
        }

        static void SetValueInSheet<T>(Excel.Worksheet sheet, int startRow, int startColumn, T[,] array)
        {
            Excel.Range range = sheet.Cells[startRow, startColumn] as Excel.Range;
            range = range.Resize[array.GetLength(0), array.GetLength(1)];
            range.Value = array;
        }

        static void SetValueInSheet<T>(Excel.Worksheet sheet, int startRow, int startColumn, List<List<T>> list)
        {
            T[,] array = ConvertTo2DArray(list);
            SetValueInSheet(sheet, startRow, startColumn, array);
        }

        static void SetValueInSheet<T>(Excel.Worksheet sheet, int startRow, int startColumn, List<T> list, bool isRow = true)
        {
            int length = list.Count;
            T[,] array = new T[isRow ? 1 : length, isRow ? length : 1];

            for (int i = 0; i < length; i++)
            {
                if (isRow)
                {
                    array[0, i] = list[i];
                }
                else
                {
                    array[i, 0] = list[i];
                }
            }

            SetValueInSheet(sheet, startRow, startColumn, array);
        }

        // リストの値を行として貼り付けます
        static void SetValueInSheetAsRow<T>(Excel.Worksheet sheet, int startRow, int startColumn, List<T> list)
        {
            SetValueInSheet(sheet, startRow, startColumn, list, true);
        }

        // リストの値を列として貼り付けます
        static void SetValueInSheetAsColumn<T>(Excel.Worksheet sheet, int startRow, int startColumn, List<T> list)
        {
            SetValueInSheet(sheet, startRow, startColumn, list, false);
        }

        class FilePathHelper
        {
            public static string GetAbsolutePathFromExecutingDirectory(string fileName)
            {
                // 実行ディレクトリを取得
                string executingDirectory = AppContext.BaseDirectory;

                // ファイル名を実行ディレクトリに結合して絶対パスを取得
                string absolutePath = Path.Combine(executingDirectory, fileName);

                return absolutePath;
            }
        }

        static Excel.Workbook CreateCopiedWorkbook(Excel.Application excelApp, string filePath, string newFilePath)
        {
            // Read-onlyでファイルを開く
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath, ReadOnly: true);

            // Save the workbook with the new name, overwriting if it already exists
            excelApp.DisplayAlerts = false;
            workbook.SaveAs(newFilePath, Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled,
                        AccessMode: Excel.XlSaveAsAccessMode.xlExclusive,
                        ConflictResolution: Excel.XlSaveConflictResolution.xlLocalSessionChanges,
                        AddToMru: false,
                        TextCodepage: false,
                        TextVisualLayout: false,
                        Local: true);
            excelApp.DisplayAlerts = true;

            // 読み取り専用のブックを閉じる
            workbook.Close(false);

            // 新しいファイルを開く
            workbook = excelApp.Workbooks.Open(newFilePath);

            // Excelを表示
            excelApp.Visible = true;

            return workbook;
        }

    }

}
