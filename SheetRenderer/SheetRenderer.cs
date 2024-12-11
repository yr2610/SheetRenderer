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

using System.Diagnostics;

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

    class RangeInfo
    {
        public int? IdColumnOffset { get; set; }
        public HashSet<int> IgnoreColumnOffsets { get; set; }
    }

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
                            <splitButton id='splitButton1' size='large'>
                              <button id='button2' label='Render' imageMso='TableDrawTable' onAction='OnRenderButtonPressed'/>
                              <menu id='menu1'>
                                <button id='button2a' label='Rerender' onAction='OnRerenderButtonPressed'/>
                              </menu>
                            </splitButton>
                            <button id='button3' label='Update Sheet' size='large' imageMso='TableSharePointListsRefreshList' onAction='OnUpdateCurrentSheetButtonPressed' getEnabled='GetUpdateCurrentSheetButtonEnabled'/>
                            <editBox id='fileNameBox' label='JSONファイル' sizeString='hoge\\20XX-XX-XX\\index' getText='GetFileName' onChange='OnTextChanged' />
                          </group>
                        </tab>
                      </tabs>
                    </ribbon>
                  </customUI>";
        }

        public void OnRerenderButtonPressed(IRibbonControl control)
        {
            MessageBox.Show("RerenderButtonPressed");
        }

        string JsonFilePath { get; set; }

        public class RenderLog
        {
            public string SourceFilePath { get; set; }
            public string User { get; set; }
        }

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

        const string indexSheetNameCustomPropertyName = "IndexSheetName";
        const string templateSheetNameCustomPropertyName = "TemplateSheetName";
        const string ssProjectIdCustomPropertyName = "SSProjectId";

        const string ssSheetRangeName = "SS_SHEET";

        class RangeInfo
        {
            public int? IdColumnOffset { get; set; }
            public HashSet<int> IgnoreColumnOffsets { get; set; }
        }

        class SheetAddressInfo
        {
            public string Address { get; set; }
            public RangeInfo RangeInfo { get; set; }
        }

        static SheetAddressInfo GetSheetAddressInfo(Excel.Worksheet sheet)
        {
            Excel.Name namedRange = sheet.GetNamedRange(ssSheetRangeName);

            if (namedRange == null)
            {
                return null;
            }

            // 名前付き範囲が存在する場合、その範囲を使用
            string address = namedRange.RefersToRange.Address;
            RangeInfo rangeInfo = null;

            // コメントが存在する場合、それを YAML として解析
            if (namedRange.Comment != null)
            {
                var deserializer = new DeserializerBuilder()
                    .WithNamingConvention(CamelCaseNamingConvention.Instance)
                    .Build();
                rangeInfo = deserializer.Deserialize<RangeInfo>(namedRange.Comment);
            }

            return new SheetAddressInfo
            {
                Address = address,
                RangeInfo = rangeInfo
            };
        }

        bool UpdateCurrentSheetButtonEnabled { get; set; } = true;

        public bool GetUpdateCurrentSheetButtonEnabled(IRibbonControl control)
        {
            return UpdateCurrentSheetButtonEnabled;
        }

        static Excel.Range GetSheetNamesRangeFromIndexSheet(Excel.Worksheet indexSheet)
        {
            return indexSheet.Range["SS_SHEETNAMELIST"];
        }

        static IEnumerable<object> GetSheetIdsFromIndexSheet(Excel.Worksheet indexSheet)
        {
            Excel.Name namedRange = indexSheet.Names.Item(ssSheetRangeName);
            var ssRange = namedRange.RefersToRange;

            var deserializer = new DeserializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();
            Debug.Assert(namedRange.Comment != null, "namedRange.Comment != null");
            RangeInfo rangeInfo = deserializer.Deserialize<RangeInfo>(namedRange.Comment);

            var sheetIds = ssRange.GetColumnValues(rangeInfo.IdColumnOffset.Value);

            return sheetIds;
        }

        // idValues を key にした行（List<object>）の dictionary を作る
        static Dictionary<string, List<object>> CreateRowDictionaryWithIDKeys(object[,] values, IEnumerable<object> idValues)
        {
#if true
            var dictionary = new Dictionary<string, List<object>>();
            int rowIndex = 1;

            foreach (var idValue in idValues)
            {
                if (idValue == null)
                {
                    rowIndex++;
                    continue;
                }

                string id = idValue.ToString();
                var rowValues = new List<object>();

                for (int j = 1; j <= values.GetLength(1); j++)
                {
                    rowValues.Add(values[rowIndex, j]);
                }

                dictionary[id] = rowValues;
                rowIndex++;
            }

            return dictionary;
#else
        // LINQ駆使した版
        return idValues
            .Zip(Enumerable.Range(1, values.GetLength(0)), (idValue, rowIndex) => (idValue, rowIndex))
            .Where(pair => pair.idValue != null)
            .ToDictionary(
                pair => pair.idValue.ToString(),
                pair => Enumerable.Range(1, values.GetLength(1))
                    .Select(colIndex => values[pair.rowIndex, colIndex])
                    .ToList()
            );
#endif
        }

        static object[,] CopyValuesById(object[,] baseValues, IEnumerable<object> baseIdValues, Dictionary<string, List<object>> valuesDictionary, HashSet<int> ignoreColumnOffsets)
        {
            object[,] result = (object[,])baseValues.Clone();

            int rowIndex = 1; // 1-originのため、1から開始

            foreach (var idValue in baseIdValues)
            {
                if (idValue == null)
                {
                    rowIndex++;
                    continue;
                }

                string id = idValue.ToString();

                if (valuesDictionary.TryGetValue(id, out var values))
                {
                    int colIndex = 1; // 1-originに変換
                    foreach (var value in values)
                    {
                        if (!ignoreColumnOffsets.Contains(colIndex - 1))
                        {
                            result[rowIndex, colIndex] = value;
                        }
                        colIndex++;
                    }
                }
                rowIndex++;
            }

            return result;
        }

        static Excel.Range GetRange(Excel.Worksheet sheet, SheetAddressInfo sheetAddressInfo)
        {
            var rangeAddress = sheetAddressInfo.Address;
            var range = sheet.Range[rangeAddress];

            return range;
        }

        static object[,] GetValues(Excel.Worksheet sheet, SheetAddressInfo sheetAddressInfo)
        {
            var range = GetRange(sheet, sheetAddressInfo);
            var values = range.Value2 as object[,];

            return values;
        }

        static IEnumerable<object> GetIds(Excel.Worksheet sheet, SheetAddressInfo sheetAddressInfo)
        {
            var rangeInfo = sheetAddressInfo?.RangeInfo;
            Debug.Assert(rangeInfo != null, "rangeInfo != null");
            Debug.Assert(rangeInfo.IdColumnOffset.HasValue, "rangeInfo.IdColumnOffset.HasValue");
            var idColumnOffset = rangeInfo.IdColumnOffset.Value;
            var rangeAddress = sheetAddressInfo.Address;
            var idValues = sheet.GetColumnWithOffset(rangeAddress, idColumnOffset);

            return idValues;
        }

        static bool AreHashSetsEqual(HashSet<int> set1, HashSet<int> set2)
        {
            // 個数を比較し、かつ各値を比較
            return set1.Count == set2.Count && !set1.Except(set2).Any();
        }

        public void OnUpdateCurrentSheetButtonPressed(IRibbonControl control)
        {
            Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook workbook = excelApp.ActiveWorkbook;
            if (workbook == null)
            {
                return;
            }
            var sheet = workbook.ActiveSheet as Excel.Worksheet;
            if (sheet == null)
            {
                return;
            }

            // 作ったシートも元のシートと同じ状態にする
            var activeCellPosition = excelApp.GetActiveCellPosition();
            var scrollPosition = excelApp.GetScrollPosition();
            var activeSheetZoom = excelApp.GetActiveSheetZoom();

            string projectId = workbook.GetCustomProperty(ssProjectIdCustomPropertyName);
            if (projectId == null)
            {
                string projectName = Assembly.GetExecutingAssembly().GetName().Name;
                MessageBox.Show($"{projectName} で生成されたブックではありません。");
                return;
            }

            // projectId があるなら他のもあるという前提
            string indexSheetName = workbook.GetCustomProperty(indexSheetNameCustomPropertyName);
            string templateSheetName = workbook.GetCustomProperty(templateSheetNameCustomPropertyName);
            var lastRenderLog = workbook.GetCustomProperty<RenderLog>("RenderLog");
            Debug.Assert(indexSheetName != null, "indexSheetName != null");
            Debug.Assert(templateSheetName != null, "templateSheetName != null");
            Debug.Assert(lastRenderLog != null, "lastRenderLog != null");

            string jsonFilePath;

            // lastRenderLog.User が今のユーザーと異なる、もしくは lastRenderLog.SourceFilePath が見つからない場合、前回生成時の環境と異なるとみなしてファイル選択させる
            if (lastRenderLog.User != Environment.UserName || !File.Exists(lastRenderLog.SourceFilePath))
            {
                MessageBox.Show($"Project ID が「{projectId}」のソースファイルを選択してください。");
                jsonFilePath = OpenFile();
                // キャンセルされたら何もしない
                if (jsonFilePath == null)
                {
                    return;
                }
            }
            else
            {
                jsonFilePath = lastRenderLog.SourceFilePath;
            }

            string jsonString = File.ReadAllText(jsonFilePath);
            JsonNode jsonObject = JsonNode.Parse(jsonString);
            var confData = GetPropertiesFromJsonNode(jsonObject, "variables");

            if (confData["project"] != projectId)
            {
                MessageBox.Show($"Project ID({projectId})が異なります。");
                return;
            }

            // 選択中の json として設定
            JsonFilePath = jsonFilePath;
            ribbon.InvalidateControl("fileNameBox");

            // 今開いているシートの id を index sheet から取得
            var indexSheet = workbook.Sheets[indexSheetName] as Excel.Worksheet;
            var sheetIds = GetSheetIdsFromIndexSheet(indexSheet);
            var sheetNameRange = GetSheetNamesRangeFromIndexSheet(indexSheet);
            var sheetNames = sheetNameRange.GetColumnValues(0);
            int sheetIndex = sheetNames.ToList().IndexOf(sheet.Name);

            if (sheetIndex == -1)
            {
                MessageBox.Show($"有効なシートが選択されていません。");
                return;
            }

            // シート名とIDをペアにして辞書に変換
            //var sheetIdMap = sheetNames.Zip(sheetIds, (sheetName, id) => new { sheetName, id })
            //                           .ToDictionary(x => x.sheetName, x => x.id);
            //string activeSheetId = sheetIdMap[sheet.Name].ToString();
            string activeSheetId = sheetIds.ElementAt(sheetIndex).ToString();
            string sheetName2 = sheetNameRange[sheetIndex + 1].value;

            JsonArray items = jsonObject["children"].AsArray();

            // jsonObject から同じ id の node を取得
            JsonNode targetSheetNode = null;
            foreach (JsonNode sheetNode in items)
            {
                string sheetId = sheetNode["id"].ToString();
                if (sheetId == activeSheetId)
                {
                    targetSheetNode = sheetNode;
                    break;
                }
            }
            // 現在のシートと同じIDのノードがなければ終了
            if (targetSheetNode == null)
            {
                MessageBox.Show($"現在のシートと同じIDのノードが存在しません。");
                return;
            }

            string sheetName = sheet.Name;
            string newSheetName = targetSheetNode["text"].ToString();

            if (newSheetName != sheetName)
            {
                // 新しい名前のシートがすでにあったら中止
                if (workbook.GetSheetIfExists(newSheetName) != null)
                {
                    MessageBox.Show($"変更後のシート名と同名のシートがすでに存在します。");
                    return;
                }
            }

            excelApp.ScreenUpdating = false;
            excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
            excelApp.EnableEvents = false;

            MacroControl.DisableMacros();

            if (newSheetName != sheetName)
            {
                // シート名が変わっていたら index sheet にも反映
                sheetNameRange.Cells[1 + sheetIndex].Value2 = newSheetName;
            }

            Excel.Worksheet templateSheet = workbook.Sheets[templateSheetName];
            templateSheet.Copy(After: sheet);

            // コピーされたシートはアクティブシートになるので、それを取得
            Excel.Worksheet newSheet = (Excel.Worksheet)templateSheet.Application.ActiveSheet;

            // 元のシートから今の入力内容を取り込む
            var sheetAddressInfo = GetSheetAddressInfo(sheet);
            var sheetValues = GetValues(sheet, sheetAddressInfo);
            var sheetValueIds = GetIds(sheet, sheetAddressInfo);

            // 元のシートを削除
            excelApp.DisplayAlerts = false;
            sheet.Delete();
            excelApp.DisplayAlerts = true;

            // 元のシートと同じ名前でも良いように元シート削除後に名前変更
            newSheet.Name = newSheetName;

            // シート作成
            // node, 画像ファイルの比較はしない
            var missingImagePathsInSheet = RenderSheet(targetSheetNode, confData, newSheet);

            var newSheetAddressInfo = GetSheetAddressInfo(newSheet);
            var newSheetRange = GetRange(newSheet, newSheetAddressInfo);
            var newSheetValueIds = GetIds(newSheet, newSheetAddressInfo);

            var ignoreColumnOffsets = newSheetAddressInfo.RangeInfo.IgnoreColumnOffsets;
            Debug.Assert(AreHashSetsEqual(ignoreColumnOffsets, sheetAddressInfo.RangeInfo.IgnoreColumnOffsets), "AreHashSetsEqual(ignoreColumnOffsets, sheetAddressInfo.RangeInfo.IgnoreColumnOffsets)");

            // idValues を key にした行（List<object>）の dictionary を作る
            var valuesDictionary = CreateRowDictionaryWithIDKeys(sheetValues, sheetValueIds);

            // newSheet の Values のコピーを作って、元のシートの Values から id を基に上書きコピーする
            // idが見つからない行、ignoreColumn は何もしないので、newSheet のものが採用される
            var result = CopyValuesById(newSheetRange.Value2, newSheetValueIds, valuesDictionary, ignoreColumnOffsets);

            newSheetRange.Value2 = result;

            // シートを元の状態と同じにする
            newSheet.Activate();
            excelApp.SetActiveCellPosition(activeCellPosition);
            excelApp.SetActiveSheetZoom(activeSheetZoom);   // scroll より後に zoom をセットすると微妙にずれるっぽい
            excelApp.SetScrollPosition(scrollPosition);

            excelApp.StatusBar = false;
            excelApp.ScreenUpdating = true;
            excelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            excelApp.EnableEvents = true;

            MacroControl.EnableMacros();

            if (missingImagePathsInSheet.Any())
            {
                ShowMissingImageFilesDialog(missingImagePathsInSheet);
            }

            // TODO: RenderLog 書き出す処理を共通化
            RenderLog renderLog = new RenderLog
            {
                SourceFilePath = JsonFilePath,
                User = Environment.UserName
            };
            workbook.SetCustomProperty("RenderLog", renderLog);
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

            string indexSheetName = workbook.GetCustomProperty(indexSheetNameCustomPropertyName);
            string templateSheetName = workbook.GetCustomProperty(templateSheetNameCustomPropertyName);
            //Debug.Assert(false, "assert test");
            if (indexSheetName == null)
            {
                // ブックを保存せずに閉じる
                workbook.Close(false);
                MessageBox.Show($"{templateFileName} のカスタムプロパティに {indexSheetNameCustomPropertyName} が設定されていません。");
                return;
            }
            if (templateSheetName == null)
            {
                // ブックを保存せずに閉じる
                workbook.Close(false);
                MessageBox.Show($"{templateFileName} のカスタムプロパティに {templateSheetNameCustomPropertyName} が設定されていません。");
                return;
            }

            excelApp.DisplayAlerts = false;
            excelApp.ScreenUpdating = false;
            excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
            excelApp.EnableEvents = false;

            var confData = GetPropertiesFromJsonNode(jsonObject, "variables");

            // 特定のプロパティ（items）を配列としてアクセス
            JsonArray items = jsonObject["children"].AsArray();

            List<string> sheetNames = new List<string>();
            Excel.Worksheet sheet = workbook.Sheets[templateSheetName];

            List<(string filePath, string sheetName, string address)> missingImagePaths = new List<(string filePath, string sheetName, string address)>();

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

                RenderIndexSheet(items, confData, indexSheet);

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

            RenderLog renderLog = new RenderLog
            {
                SourceFilePath = JsonFilePath,
                User = Environment.UserName
            };
            workbook.SetCustomProperty("RenderLog", renderLog);

            //var lastRenderLog = workbook.GetCustomProperty<RenderLog>("RenderLog");

            string projectId = confData["project"];
            workbook.SetCustomProperty(ssProjectIdCustomPropertyName, projectId);

            workbook.Save();
        }

        static void RenderIndexSheet(IEnumerable<JsonNode> sheetNodes, Dictionary<string, string> confData, Excel.Worksheet indexSheet)
        {
            var sheetNameListRange = indexSheet.GetNamedRange("SS_SHEETNAMELIST").RefersToRange;

            int indexStartRow = sheetNameListRange.Row;
            int indexRowCount = sheetNameListRange.Rows.Count;
            int indexEndRow = sheetNameListRange.Rows[indexRowCount].Row;
            int indexStartColumn = sheetNameListRange.Column;
            string idColumnAddress = "T";
            int idColumn = indexSheet.ColumnAddressToIndex(idColumnAddress);

            string syncStartColumnAddress = "Q";
            int syncStartColumn = indexSheet.ColumnAddressToIndex(syncStartColumnAddress);
            int syncStartColumnCount = 1;
            int[] syncIgnoreColumnOffsets = { };

            var sheetNames = ExtractPropertyValues(sheetNodes, "text");
            var sheetNamesCount = sheetNames.Count();

            // 行が足りなければ挿入
            if (sheetNamesCount > indexRowCount)
            {
                int numberOfRows = sheetNamesCount - indexRowCount;

                indexSheet.InsertRowsAndCopyFormulas(indexEndRow, numberOfRows);
            }
            // 多ければ削除
            else if (sheetNamesCount < indexRowCount)
            {
                int numberOfRowsToDelete = indexRowCount - sheetNamesCount;

                indexSheet.DeleteRows(indexStartRow, numberOfRowsToDelete);
            }

            indexSheet.SetValueInSheetAsColumn(indexStartRow, indexStartColumn, sheetNames);

            // テンプレ処理
            ReplaceValues(indexSheet, confData);

            // 幅をautofit
            indexSheet.Cells[indexStartRow, indexStartColumn].Resize(indexRowCount).Columns.AutoFit();

            // 適当な位置に列挿入して ID を入れて非表示にする
            var ids = ExtractPropertyValues(sheetNodes, "id");
            Excel.Range column = (Excel.Range)indexSheet.Columns[idColumn];
            column.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            indexSheet.SetValueInSheet(indexStartRow, idColumn, ids, false);
            Excel.Range idColumnRange = indexSheet.Columns[idColumn];
            idColumnRange.EntireColumn.Hidden = true;

            // 名前付き範囲として追加
            var rangeforNamedRange = indexSheet.GetRange(indexStartRow, syncStartColumn, sheetNamesCount, syncStartColumnCount);
            var namedRange = indexSheet.Names.Add(Name: ssSheetRangeName, RefersTo: rangeforNamedRange);
            RangeInfo rangeInfo = new RangeInfo
            {
                IdColumnOffset = idColumn - syncStartColumn,
                IgnoreColumnOffsets = new HashSet<int>(syncIgnoreColumnOffsets),
            };
            var serializer = new SerializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();

            namedRange.Comment = serializer.Serialize(rangeInfo);

        }

        static void ShowMissingImageFilesDialog(IEnumerable<(string filePath, string sheetName, string address)> missingFiles)
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

            ListBox listBox = new ListBox
            {
                Dock = DockStyle.Fill,
            };
            listBox.Font = new Font(listBox.Font.FontFamily, listBox.Font.Size * 2); // フォントサイズを2倍に設定

            //missingFiles = missingFiles.Distinct();
            foreach (var file in missingFiles)
            {
                listBox.Items.Add(file.filePath);
            }

            listBox.Click += (sender, e) =>
            {
                if (listBox.SelectedItem != null)
                {
                    var selectedCell = missingFiles.ElementAtOrDefault(listBox.SelectedIndex);
                    var sheetName = selectedCell.sheetName;
                    var cellAddress = selectedCell.address;
                    var excelApp = (Excel.Application)ExcelDnaUtil.Application;
                    var sheet = (Excel.Worksheet)excelApp.Sheets[sheetName];
                    var range = sheet.Range[cellAddress];
                    sheet.Activate();
                    range.Select();
                }
            };

            form.Controls.Add(listBox);
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

        static void SetValueInSheet<T, TOutput>(Excel.Worksheet sheet, int startRow, int startColumn, List<List<T>> list, Func<T, TOutput> selector)
        {
            TOutput[,] array = ConvertTo2DArray(list, selector);
            sheet.SetValueInSheet(startRow, startColumn, array);
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

        IEnumerable<(string filePath, string sheetName, string address)> RenderSheet(JsonNode sheetNode, Dictionary<string, string> confData, Excel.Worksheet sheet)
        {
            List<JsonNode> leafNodes;
            int maxDepth;
            List<List<NodeData>> result = TraverseTreeFromRoot(sheetNode, out leafNodes, out maxDepth);
            int leafCount = leafNodes.Count;
            List<(string filePath, string sheetName, string address)> missingImagePaths = new List<(string filePath, string sheetName, string address)>();
            
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

                sheet.InsertRowsAndCopyFormulas(endRow, numberOfRows);
            }
            else if (leafCount < rowHeight)
            {
                int numberOfRowsToDelete = rowHeight - leafCount;

                sheet.DeleteRows(startRow, numberOfRowsToDelete);
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

            sheet.SetValueInSheet(startRow, startColumn, arrayResult);

            sheet.GetRange(startRow, startColumn, leafCount, maxDepth).AutoFitColumnsIfNarrower();

            // 「チェック予定日」列に無条件で START_DATE を入れる
            int dateColumnOffset = initialDateColumn - initialResultColumn;
            int dateColumn = startColumn + maxDepth + dateColumnOffset;
            string dateString = confData["START_DATE"];
            // 文字列をDateTimeに変換
            if (DateTime.TryParse(dateString, out DateTime dateValue))
            {
                sheet.SetRangeValue(startRow, dateColumn, leafCount, 1, dateValue);
            }

            // 「チェック予定日」の右隣の列に ID を入れて非表示にする
            int idColumnOffset = initialIdColumn - initialResultColumn;
            int idColumn = startColumn + maxDepth + idColumnOffset;
            var ids = ExtractPropertyValues(leafNodes, "id");
            sheet.SetValueInSheet(startRow, idColumn, ids, false);
            Excel.Range idColumnRange = sheet.Columns[idColumn];
            idColumnRange.EntireColumn.Hidden = true;

            // 「チェック結果」列に node の initialValues.result を入れる
            int resultColumnOffset = initialResultColumn - initialResultColumn;
            int resultColumn = startColumn + maxDepth + resultColumnOffset;
            var results = ExtractPropertyValuesFromInitialValues(leafNodes, "result");
            sheet.SetValueInSheet(startRow, resultColumn, results, false);

            // 「計画時間(分)」列に node の initialValues.estimated_time を入れる
            int planColumnOffset = initialPlanColumn - initialResultColumn;
            int planColumn = startColumn + maxDepth + planColumnOffset;
            var estimatedTimes = ExtractPropertyValuesFromInitialValues(leafNodes, "estimated_time");
            sheet.SetValueInSheet(startRow, planColumn, estimatedTimes, false);

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
            var rangeforNamedRange = sheet.GetRange(startRow, resultColumn, leafCount, 1 + actualTimeColumnOffset);
            var namedRange = sheet.Names.Add(Name: ssSheetRangeName, RefersTo: rangeforNamedRange);
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
                            missingImagePaths.Add((filePath: node.imageFilePath, sheetName: sheet.Name, address: cell.Address));
                        }

                        AddPictureAsComment(cell, path);
                    }
                }
            }

            return missingImagePaths;
        }

        static void AddPictureAsComment(Excel.Range cell, string imageFilePath)
        {
            // 画像のサイズを取得
            System.Drawing.Image image = System.Drawing.Image.FromFile(imageFilePath);
            float imageWidth = image.Width;
            float imageHeight = image.Height;
            image.Dispose();

            // コメントを追加し、画像を背景に設定
            var comment = cell.AddComment(" ");
            comment.Visible = false;
            comment.Shape.Fill.UserPicture(imageFilePath);
            comment.Shape.Height = imageHeight;
            comment.Shape.Width = imageWidth;
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
