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

using System.Drawing;

using System.Reflection;

using ExcelDna.Integration;

using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;


public class ProgressBarForm : Form
{
    private ProgressBar progressBar;
    private Label progressLabel;
    private Label sheetNameLabel;
    private int totalSheets;
    private int completedSheets;

    public ProgressBarForm(int totalSheets)
    {
        this.totalSheets = totalSheets;
        this.completedSheets = 0;

        // フォームのサイズを設定
        this.Width = 500;
        this.Height = 180;

        // フォームのスタイルを設定
        this.FormBorderStyle = FormBorderStyle.FixedSingle;
        this.ControlBox = false;
        this.Text = string.Empty;

        // プログレスバーを作成
        progressBar = new ProgressBar();
        progressBar.Width = 450;
        progressBar.Height = 20;
        progressBar.Top = 30;
        progressBar.Left = 25;
        Controls.Add(progressBar);

        // 進行状況を表示するラベルを作成
        progressLabel = new Label();
        progressLabel.Width = 450;
        progressLabel.Top = 60;
        progressLabel.Left = 25;
        progressLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
        Controls.Add(progressLabel);

        // シート名を表示するラベルを作成
        sheetNameLabel = new Label();
        sheetNameLabel.Width = 450;
        sheetNameLabel.Top = 90;
        sheetNameLabel.Left = 25;
        sheetNameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
        Controls.Add(sheetNameLabel);

        // フォームの閉じる操作を無効にする
        this.FormClosing += new FormClosingEventHandler(ProgressBarForm_FormClosing);
    }

    private void ProgressBarForm_FormClosing(object sender, FormClosingEventArgs e)
    {
        // フォームの閉じる操作をキャンセル
        e.Cancel = true;
    }

    public void UpdateSheetName(string sheetName)
    {
        if (InvokeRequired)
        {
            Invoke(new Action<string>(UpdateSheetName), sheetName);
        }
        else
        {
            sheetNameLabel.Text = $"作成中のシート: {sheetName}";
            completedSheets++;
            progressBar.Value = (int)((double)completedSheets / totalSheets * 100);
            progressLabel.Text = $"進行状況: {completedSheets} / {totalSheets}";
        }
    }

    public void CloseForm()
    {
        if (InvokeRequired)
        {
            Invoke(new Action(CloseForm));
        }
        else
        {
            // FormClosingイベントハンドラーを一時的に無効にする
            this.FormClosing -= ProgressBarForm_FormClosing;
            this.Close();
        }
    }

    protected override void WndProc(ref Message m)
    {
        const int WM_NCHITTEST = 0x84;
        const int HTCLIENT = 0x1;
        const int HTCAPTION = 0x2;

        if (m.Msg == WM_NCHITTEST)
        {
            base.WndProc(ref m);
            if ((int)m.Result == HTCLIENT)
            {
                m.Result = (IntPtr)HTCAPTION;
                return;
            }
        }
        base.WndProc(ref m);
    }

}


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
        private ProgressBarForm progressBarForm;

        public void OnLoad(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public override string GetCustomUI(string RibbonID)
        {
            string projectName = Assembly.GetExecutingAssembly().GetName().Name;
            return $@"
      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='OnLoad'>
      <ribbon>
        <tabs>
          <tab id='tab1' label='{projectName}'>
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

        // セル内の {{*}} を置き換える
        public static void ReplaceValues(Excel.Worksheet sheet, Dictionary<string, string> replacements)
        {
            Excel.Range usedRange = sheet.UsedRange;

            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;

            object[,] values = usedRange.Value2;
            bool[,] modified = new bool[rowCount + 1, colCount + 1];

            Regex regex = new Regex(@"\{\{([_A-Za-z]\w*)\}\}");
            Regex urlRegex = new Regex(@"https?://[^\s/$.?#].[^\s]*");

            // まず全セルの値を配列に取り込む
            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= colCount; col++)
                {
                    if (!(values[row, col] is string))
                    {
                        continue;
                    }

                    string cellValue = (string)values[row, col];

                    Match match = regex.Match(cellValue);
                    if (match.Success)
                    {
                        string key = match.Groups[1].Value;
                        if (replacements.ContainsKey(key))
                        {
                            values[row, col] = regex.Replace(cellValue, replacements[key]);
                            modified[row, col] = true;
                        }
                        else
                        {
                            values[row, col] = regex.Replace(cellValue, "");
                            modified[row, col] = true;
                        }
                    }
                    else
                    {
                        values[row, col] = null;
                    }
                }
            }

            // 必要なセルのみ HasFormula をチェックして書き戻す
            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= colCount; col++)
                {
                    if (modified[row, col] && !usedRange.Cells[row, col].HasFormula)
                    {
                        string newValue = values[row, col] as string;
                        if (urlRegex.IsMatch(newValue))
                        {
                            // 新しいリンクを設定
                            sheet.Hyperlinks.Add(usedRange.Cells[row, col], newValue);
                        }
                        // セルの値を設定
                        usedRange.Cells[row, col].Value2 = newValue;
                    }
                }
            }
        }

        public async void OnRenderButtonPressed(IRibbonControl control)
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

            string templateFilePath = GetAbsolutePathFromExecutingDirectory(templateFileName);

            MacroControl.DisableMacros();

            Excel.Workbook workbook = CreateCopiedWorkbook(excelApp, templateFilePath, newFilePath);

            excelApp.DisplayAlerts = false;
            excelApp.ScreenUpdating = false;
            excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
            excelApp.EnableEvents = false;

            const string indexSheetName = "表紙";
            const string templateSheetName = "単列チェック結果format";

            var confData = GetPropertiesFromJsonNode(jsonObject, "variables");

            // 特定のプロパティ（items）を配列としてアクセス
            JsonArray items = jsonObject["children"].AsArray();

            List<string> sheetNames = new List<string>();
            Excel.Worksheet sheet = workbook.Sheets[templateSheetName];

            List<string> missingImagePaths = new List<string>();

            // +1 は index シート
            progressBarForm = new ProgressBarForm(items.Count + 1);
            progressBarForm.Show();

            await Task.Run(() =>
            {
                foreach (JsonNode sheetNode in items)
                {
                    string newSheetName = sheetNode["text"].ToString();

                    // プログレスバーを更新
                    progressBarForm.Invoke(new Action<string>(progressBarForm.UpdateSheetName), newSheetName);

                    // シートをコピーしてリネーム
                    sheet.Copy(After: workbook.Sheets[workbook.Sheets.Count]);
                    Excel.Worksheet newSheet = workbook.Sheets[workbook.Sheets.Count];
                    newSheet.Name = newSheetName;

                    var missingImagePathsInSheet = RenderSheet(sheetNode, confData, newSheet);

                    missingImagePaths.AddRange(missingImagePathsInSheet);

                    sheetNames.Add(newSheetName);

                }

                // プログレスバーを更新
                progressBarForm.Invoke(new Action<string>(progressBarForm.UpdateSheetName), indexSheetName);

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

                // テンプレ処理
                ReplaceValues(indexSheet, confData);

                // 幅をautofit
                indexSheet.Cells[indexStartRow, indexStartColumn].Resize(indexRowCount).Columns.AutoFit();

                // 最後にindexシートを選択状態にしておく
                indexSheet.Select();

                // 処理が完了したらフォームを閉じる
                progressBarForm.Invoke(new Action(progressBarForm.CloseForm));
            });

            progressBarForm.Close();

            excelApp.EnableEvents = true;
            excelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            excelApp.ScreenUpdating = true;
            excelApp.DisplayAlerts = true;

            MacroControl.EnableMacros();

            if (missingImagePaths.Any())
            {
                ShowMissingImageFilesDialog(missingImagePaths);
            }

            workbook.Save();
        }

        static void ShowMissingImageFilesDialog(IEnumerable<string> missingFiles)
        {
            Form form = new Form
            {
                Text = "Missing Files",
                Width = 400,
                Height = 300,
                TopMost = true // topmostに設定
            };

            Label label = new Label
            {
                Text = "以下の画像ファイルが見つかりませんでした。",
                Dock = DockStyle.Top,
                AutoSize = true,
            };
            label.Font = new Font(label.Font.FontFamily, label.Font.Size * 2); // フォントサイズを2倍に設定

            // ラベルのテキスト幅に基づいてフォームの幅を設定
            using (Graphics g = label.CreateGraphics())
            {
                SizeF size = g.MeasureString(label.Text, label.Font);
                form.Width = (int)size.Width + 40; // 余白を追加
            }

            TextBox textBox = new TextBox
            {
                Multiline = true,
                Dock = DockStyle.Fill,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
            };
            textBox.Font = new Font(textBox.Font.FontFamily, textBox.Font.Size * 2); // フォントサイズを2倍に設定

            missingFiles = missingFiles.Distinct();
            foreach (var file in missingFiles)
            {
                textBox.AppendText(file + Environment.NewLine);
            }

            form.Controls.Add(textBox);
            form.Controls.Add(label);

            form.Show(); // モードレスウィンドウとして表示
        }

        // マクロを一時的に黙らせたい
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

        class NodeData
        {
            public string kind { get; set; }
            public int group { get; set; }
            public int depthInGroup { get; set; }
            public string id { get; set; }
            public string text { get; set; }
            public string comment { get; set; }
            public string imageFilePath { get; set; }
            public Dictionary<string, object> initialValues { get; set; }
        };

        static void TraverseTree(JsonNode node, List<List<NodeData>> result, List<NodeData> currentPath, int currentDepth, ref int maxDepth, List<JsonNode> leafNodes)
        {
            if (node == null) return;

            NodeData nodeData = node.Deserialize<NodeData>();
            currentPath.Add(nodeData);
            var nullifiedPath = Enumerable.Repeat<NodeData>(null, currentPath.Count).ToList();

            JsonArray children = node["children"] as JsonArray;
            if (children != null && children.Count > 0)
            {
                for (int i = 0; i < children.Count; i++)
                {
                    JsonNode child = children[i];
                    var currentOrNullifiedPath = new List<NodeData>(i == 0 ? currentPath : nullifiedPath);
                    TraverseTree(child, result, currentOrNullifiedPath, currentDepth + 1, ref maxDepth, leafNodes);
                }
            }
            else
            {
                // 現在のパスを結果に追加
                result.Add(new List<NodeData>(currentPath));
                leafNodes.Add(node);
            }

            // 最大深度を更新
            if (currentDepth > maxDepth)
            {
                maxDepth = currentDepth;
            }
        }

        static List<List<NodeData>> TraverseTreeFromRoot(JsonNode rootNode, out List<JsonNode> leafNodes, out int maxDepth)
        {
            List<List<NodeData>> result = new List<List<NodeData>>();
            maxDepth = 0;
            leafNodes = new List<JsonNode>();

            TraverseTree(rootNode, result, new List<NodeData>(), 1, ref maxDepth, leafNodes);

            return result;
        }

        static TOutput[,] ConvertTo2DArray<TInput, TOutput>(List<List<TInput>> list, Func<TInput, TOutput> selector)
        {
            int rows = list.Count;
            int cols = rows > 0 ? list.Max(subList => subList.Count) : 0;
            TOutput[,] array = new TOutput[rows, cols];

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < list[i].Count; j++)
                {
                    array[i, j] = selector(list[i][j]);
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

        static T[,] RemoveFirstColumn<T>(T[,] array)
        {
            int rows = array.GetLength(0);
            int cols = array.GetLength(1);

            T[,] result = new T[rows, cols - 1];

            for (int i = 0; i < rows; i++)
            {
                for (int j = 1; j < cols; j++)
                {
                    result[i, j - 1] = array[i, j];
                }
            }

            return result;
        }

        static List<List<T>> RemoveFirstColumn<T>(List<List<T>> list)
        {
            List<List<T>> result = new List<List<T>>();

            foreach (var row in list)
            {
                List<T> newRow = new List<T>(row.Skip(1));
                result.Add(newRow);
            }

            return result;
        }

        static object GetJsonValue(JsonValue jsonValue)
        {
            if (jsonValue.TryGetValue(out int intValue))
            {
                return intValue;
            }
            else if (jsonValue.TryGetValue(out double doubleValue))
            {
                return doubleValue;
            }
            else if (jsonValue.TryGetValue(out bool boolValue))
            {
                return boolValue;
            }
            else if (jsonValue.TryGetValue(out string stringValue))
            {
                return stringValue;
            }
            else
            {
                return jsonValue.ToString();
            }
        }

        static object GetJsonValue(JsonNode jsonNode)
        {
            return (jsonNode != null) ? GetJsonValue(jsonNode.AsValue()) : null;
        }

        static IEnumerable<object> ExtractPropertyValuesFromInitialValues(IEnumerable<JsonNode> leafNodes, string propertyName)
        {
            return leafNodes.Select(node =>
            {
                var initialValuesNode = node["initialValues"];
                return (initialValuesNode != null) ? GetJsonValue(initialValuesNode[propertyName]) : null;
            });
        }

        static IEnumerable<object> ExtractPropertyValues(IEnumerable<JsonNode> leafNodes, string propertyName)
        {
            return leafNodes.Select(node => GetJsonValue(node[propertyName]));
        }

        static Excel.Range GetRange(Excel.Worksheet sheet, int startRow, int startColumn, int rowCount, int columnCount)
        {
            Excel.Range startCell = (Excel.Range)sheet.Cells[startRow, startColumn];
            Excel.Range endCell = (Excel.Range)sheet.Cells[startRow + rowCount - 1, startColumn + columnCount - 1];
            Excel.Range range = sheet.Range[startCell, endCell];
            return range;
        }

        static void AutoFitColumnsIfNarrower(Excel.Range range)
        {
            Excel.Worksheet sheet = range.Worksheet;

            // 各列の元の幅を保存
            double[] originalWidths = new double[range.Columns.Count];
            int index = 0;
            foreach (Excel.Range column in range.Columns)
            {
                originalWidths[index++] = column.ColumnWidth;
            }

            // 一度にAutoFitを適用
            range.Columns.AutoFit();

            // 各列の幅をチェックして、元の幅よりも広くなった場合は元に戻す
            index = 0;
            foreach (Excel.Range column in range.Columns)
            {
                if (column.ColumnWidth > originalWidths[index])
                {
                    column.ColumnWidth = originalWidths[index];
                }
                index++;
            }
        }

        class RangeInfo
        {
            public int? IdColumnOffset { get; set; }
            public HashSet<int> IgnoreColumnOffsets { get; set; }
        }

        IEnumerable<string> RenderSheet(JsonNode sheetNode, Dictionary<string, string> confData, Excel.Worksheet sheet)
        {
            List<JsonNode> leafNodes;
            int maxDepth;
            List<List<NodeData>> result = TraverseTreeFromRoot(sheetNode, out leafNodes, out maxDepth);
            int leafCount = leafNodes.Count;
            List<string> missingImagePaths = new List<string>();

            // 左端はシート名なので削除
            result = RemoveFirstColumn(result);
            maxDepth--;

            // 2次元配列に変換
            string[,] arrayResult = ConvertTo2DArray(result, x => x?.text);

            // 右端にダミー文字追加
            ReplaceTrailingNullsInLastColumn(arrayResult, "---");

            int startColumn = 3;
            const int endColumn = 6;
            int columnWidth = endColumn - startColumn + 1;
            const int startRow = 25;
            const int endRow = 102;
            int rowHeight = endRow - startRow + 1;

            const int initialDateColumn = 13;
            const int initialPlanColumn = 11;
            const int initialActualTimeColumn = 12;
            const int initialResultColumn = 7;

            // 左端だと階層浅い時に非表示にできないC,D列になるので右端で
            int initialIdColumn = 14;

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

            AutoFitColumnsIfNarrower(GetRange(sheet, startRow, startColumn, leafCount, maxDepth));

            // 「チェック予定日」列に無条件で START_DATE を入れる
            int dateColumnOffset = initialDateColumn - initialResultColumn;
            int dateColumn = startColumn + maxDepth + dateColumnOffset;
            string dateString = confData["START_DATE"];
            // 文字列をDateTimeに変換
            if (DateTime.TryParse(dateString, out DateTime dateValue))
            {
                SetRangeValue(sheet, startRow, dateColumn, leafCount, 1, dateValue);
            }

            // 「チェック予定日」の右隣の列に ID を入れて非表示にする
            int idColumnOffset = initialIdColumn - initialResultColumn;
            int idColumn = startColumn + maxDepth + idColumnOffset;
            var ids = ExtractPropertyValues(leafNodes, "id");
            SetValueInSheet(sheet, startRow, idColumn, ids, false);
            Excel.Range idColumnRange = sheet.Columns[idColumn];
            idColumnRange.EntireColumn.Hidden = true;

            // 「チェック結果」列に node の initialValues.result を入れる
            int resultColumnOffset = initialResultColumn - initialResultColumn;
            int resultColumn = startColumn + maxDepth + resultColumnOffset;
            var results = ExtractPropertyValuesFromInitialValues(leafNodes, "result");
            SetValueInSheet(sheet, startRow, resultColumn, results, false);

            // 「計画時間(分)」列に node の initialValues.estimated_time を入れる
            int planColumnOffset = initialPlanColumn - initialResultColumn;
            int planColumn = startColumn + maxDepth + planColumnOffset;
            var estimatedTimes = ExtractPropertyValuesFromInitialValues(leafNodes, "estimated_time");
            SetValueInSheet(sheet, startRow, planColumn, estimatedTimes, false);

            int actualTimeColumnOffset = initialActualTimeColumn - initialResultColumn;

            // XXX: コメント系セルの対応
            for (int i = 0; i < result.Count; i++)
            {
                for (int j = 0; j < result[i].Count; j++)
                {
                    NodeData node = result[i][j];

                    if (node == null)
                    {
                        continue;
                    }

                    bool InitializeCommentCell(string text_, string pattern, int cellColorIndex, int fontColorIndex)
                    {
                        if (!Regex.IsMatch(text_, pattern))
                        {
                            return false;
                        }

                        // ここより右のセルの色を変える
                        var cells = sheet.Cells[startRow + i, startColumn + j];
                        cells = cells.Resize(1, maxDepth - j);
                        cells.Interior.ColorIndex = cellColorIndex;
                        cells.Font.ColorIndex = fontColorIndex;

                        // チェック予定日欄を空欄にする
                        var dateCell = sheet.Cells[startRow + i, dateColumn];
                        dateCell.Value = null;

                        return true;
                    }

                    string text = node.text;

                    const string descPattern = @"^【.*】";
                    const int descCellColorIndex = 37;   // 水色っぽい色
                    const int descFontColorIndex = 1;   // 黒
                    if (InitializeCommentCell(text, descPattern, descCellColorIndex, descFontColorIndex))
                    {
                        break;
                    }

                    const string warningPattern = @"^※※";
                    const int warningCellThemeColorId = 22;
                    const int warningFontColorIndex = 9;
                    if (InitializeCommentCell(text, warningPattern, warningCellThemeColorId, warningFontColorIndex))
                    {
                        break;
                    }
                }
            }

            // 名前付き範囲として追加
            const string sheetRangeName = "SS_SHEET";
            var rangeforNamedRange = GetRange(sheet, startRow, resultColumn, leafCount, 1 + actualTimeColumnOffset);
            var namedRange = sheet.Names.Add(Name: sheetRangeName, RefersTo: rangeforNamedRange);
            RangeInfo rangeInfo = new RangeInfo
            {
                IdColumnOffset = idColumnOffset,
                IgnoreColumnOffsets = new HashSet<int> { planColumnOffset },
            };
            var serializer = new SerializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();

            namedRange.Comment = serializer.Serialize(rangeInfo);
            
            // 画像を貼る
            for (int i = 0; i < result.Count; i++)
            {
                for (int j = 0; j < result[i].Count; j++)
                {
                    NodeData node = result[i][j];

                    if (node == null)
                    {
                        continue;
                    }

                    if (node.imageFilePath != null)
                    {
                        const string noImageFilePath = "images/no_image.jpg";
                        string path = GetAbsolutePathFromBasePath(JsonFilePath, node.imageFilePath);
                        var cell = sheet.Cells[startRow + i, startColumn + j];

                        if (!File.Exists(path))
                        {
                            // XXX: 毎回パス構築はムダ
                            path = GetAbsolutePathFromExecutingDirectory(noImageFilePath);
                            missingImagePaths.Add(node.imageFilePath);
                        }

                        AddPictureAsComment(cell, path);
                    }
                }
            }

            return missingImagePaths;
        }

        static void AddPictureAsComment(Excel.Range cell, string path)
        {
            // 画像のサイズを取得
            System.Drawing.Image image = System.Drawing.Image.FromFile(path);
            float imageWidth = image.Width;
            float imageHeight = image.Height;
            image.Dispose();

            // コメントを追加し、画像を背景に設定
            var comment = cell.AddComment(" ");
            comment.Visible = false;
            comment.Shape.Fill.UserPicture(path);
            comment.Shape.Height = imageHeight;
            comment.Shape.Width = imageWidth;
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

        static void SetValueInSheet<T, TOutput>(Excel.Worksheet sheet, int startRow, int startColumn, List<List<T>> list, Func<T, TOutput> selector)
        {
            TOutput[,] array = ConvertTo2DArray(list, selector);
            SetValueInSheet(sheet, startRow, startColumn, array);
        }

        static void SetValueInSheet<T>(Excel.Worksheet sheet, int startRow, int startColumn, IEnumerable<T> source, bool isRow = true)
        {
            int length = source.Count();
            T[,] array = new T[isRow ? 1 : length, isRow ? length : 1];
            int index = 0;

            foreach (var item in source)
            {
                if (isRow)
                {
                    array[0, index] = item;
                }
                else
                {
                    array[index, 0] = item;
                }
                index++;
            }

            SetValueInSheet(sheet, startRow, startColumn, array);
        }

        // リストの値を行として貼り付けます
        static void SetValueInSheetAsRow<T>(Excel.Worksheet sheet, int startRow, int startColumn, IEnumerable<T> source)
        {
            SetValueInSheet(sheet, startRow, startColumn, source, true);
        }

        // リストの値を列として貼り付けます
        static void SetValueInSheetAsColumn<T>(Excel.Worksheet sheet, int startRow, int startColumn, IEnumerable<T> source)
        {
            SetValueInSheet(sheet, startRow, startColumn, source, false);
        }

        static string GetAbsolutePathFromExecutingDirectory(string relativePath)
        {
            // 実行ディレクトリを取得
            string executingDirectory = AppContext.BaseDirectory;

            // GetAbsolutePathFromBasePath メソッドを呼び出して絶対パスを取得
            return GetAbsolutePathFromBasePath(executingDirectory, relativePath);
        }

        static string GetAbsolutePathFromBasePath(string basePath, string relativePath)
        {
            // basePath がファイルパスの場合、そのディレクトリを取得
            if (File.Exists(basePath))
            {
                basePath = Path.GetDirectoryName(basePath);
            }

            // 絶対パス基準の相対パスを絶対パスに変換
            string absolutePath = Path.Combine(basePath, relativePath);

            return absolutePath;
        }

        static string GetRelativePathFromExecutingDirectory(string absolutePath)
        {
            // 実行ディレクトリを取得
            string executingDirectory = AppContext.BaseDirectory;

            return GetRelativePathFromBasePath(executingDirectory, absolutePath);
        }

        static string GetRelativePathFromBasePath(string basePath, string absolutePath)
        {
            Uri baseUri = new Uri(basePath);
            Uri absoluteUri = new Uri(absolutePath);
            Uri relativeUri = baseUri.MakeRelativeUri(absoluteUri);
            return Uri.UnescapeDataString(relativeUri.ToString().Replace('/', Path.DirectorySeparatorChar));
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
