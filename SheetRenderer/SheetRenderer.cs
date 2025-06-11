﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using System.IO;

using System.Text.RegularExpressions;

using System.Windows.Forms;

using System.Security.Cryptography;

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

    public static class JsonNodeHasher
    {
        public static string ComputeSha256(this JsonNode sheetNode)
        {
            // JSONノードをバイト配列に直接変換
            byte[] jsonBytes = JsonSerializer.SerializeToUtf8Bytes(sheetNode);

            // SHA256ハッシュを計算
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hashBytes = sha256.ComputeHash(jsonBytes);
                // ハッシュを16進数文字列に変換
                return BitConverter.ToString(hashBytes).Replace("-", "").ToLower();
            }
        }

#if false
           public static string ComputeHash(JsonNode node)
           {
                // TODO: 時間計測して比較
#if false
                string jsonString = node.ToJsonString();
                using (var sha256 = SHA256.Create())
                {
                    byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(jsonString));
                    return BitConverter.ToString(hashBytes).Replace("-", "").ToLower();
                }
#else
                using (var sha256 = SHA256.Create())
                {
                    var result = ComputeHashRecursive(node, sha256);

                    // ハッシュ計算を完了するために TransformFinalBlock を呼び出す
                    sha256.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
                    byte[] hashBytes = sha256.Hash;

                    return BitConverter.ToString(hashBytes).Replace("-", "").ToLower();
                }
#endif
            }

            private static SHA256 ComputeHashRecursive(JsonNode node, SHA256 sha256)
            {
                if (node is JsonObject jsonObject)
                {
                    var sortedKeys = jsonObject.Select(kvp => kvp.Key).OrderBy(k => k);
                    foreach (var key in sortedKeys)
                    {
                        var keyBytes = Encoding.UTF8.GetBytes(key);
                        sha256.TransformBlock(keyBytes, 0, keyBytes.Length, keyBytes, 0);
                        ComputeHashRecursive(jsonObject[key], sha256);
                    }
                }
                else if (node is JsonArray jsonArray)
                {
                    foreach (var item in jsonArray)
                    {
                        ComputeHashRecursive(item, sha256);
                    }
                }
                else if (node is JsonValue jsonValue)
                {
                    var valueBytes = Encoding.UTF8.GetBytes(jsonValue.ToString());
                    sha256.TransformBlock(valueBytes, 0, valueBytes.Length, valueBytes, 0);
                }

                return sha256;
            }
#endif
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
                            <splitButton id='splitButton1' size='large'>
                              <button id='button2' label='ファイル更新' imageMso='TableDrawTable' onAction='OnRenderButtonPressed'/>
                              <menu id='menu1'>
                                <button id='button2a' label='ファイル新規作成' onAction='OnCreateNewButtonPressed'/>
                              </menu>
                            </splitButton>
                            <button id='button3' label='シート更新' size='large' imageMso='TableSharePointListsRefreshList' onAction='OnUpdateCurrentSheetButtonPressed' getEnabled='GetUpdateCurrentSheetButtonEnabled'/>
                          </group>
                        </tab>
                      </tabs>
                    </ribbon>
                  </customUI>";
        }

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

        public static string OpenSourceFile()
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

        // テンプレートセル情報
        class CellInfo
        {
            public string Address { get; set; }
            public string Value { get; set; }
        }

        static List<CellInfo> GetTemplateCells(Excel.Worksheet sheet)
        {
            Excel.Range usedRange = sheet.UsedRange;

            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;

            if (rowCount == 1 && colCount == 1)
            {
                // 配列として処理するため適当に2セルにする
                // colCount は 1 のままで良い
                usedRange = usedRange.Resize[1, 2];
            }

            int startRow = usedRange.Row;
            int startCol = usedRange.Column;

            // セルのアドレスを計算するヘルパーメソッド
            string GetCellAddress(int row, int col)
            {
                int actualRow = startRow + row - 1;
                int actualCol = startCol + col - 1;
                return $"{GetColumnLetter(actualCol)}{actualRow}";
            }

            // 列番号を列文字に変換するヘルパーメソッド
            string GetColumnLetter(int col)
            {
                string columnLetter = "";
                while (col > 0)
                {
                    int mod = (col - 1) % 26;
                    columnLetter = (char)(mod + 65) + columnLetter;
                    col = (col - mod) / 26;
                }
                return columnLetter;
            }

            object[,] values = usedRange.Value2;
            Regex regex = new Regex(@"\{\{[_A-Za-z]\w*\}\}");
            List<CellInfo> matchingCells = new List<CellInfo>();

            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= colCount; col++)
                {
                    if (!(values[row, col] is string))
                    {
                        continue;
                    }

                    string cellValue = (string)values[row, col];

                    // パターンにマッチするかどうかだけを判定
                    if (regex.IsMatch(cellValue))
                    {
                        string cellAddress = GetCellAddress(row, col);
                        var info = new CellInfo
                        {
                            Address = cellAddress,
                            Value = cellValue
                        };

                        matchingCells.Add(info);
                    }
                }
            }

            return matchingCells;
        }

        static string ReplacePlaceholders(string s, Dictionary<string, string> parameters)
        {
            // 正規表現パターンを定義します。識別子はC言語のルールに従います。
            var pattern = @"\{\{(\w[\w\d]*)\}\}";

            // 正規表現でマッチングされた部分を置換します。
            return Regex.Replace(s, pattern, match =>
            {
                var key = match.Groups[1].Value;
                if (parameters.TryGetValue(key, out var value))
                {
                    return value;
                }
                // 見つからないキーの場合はそのまま返す
                return match.Value;
            });
        }

        static void SetCellValues(Excel.Worksheet worksheet, List<CellInfo> cellInfos)
        {
            Regex urlRegex = new Regex(@"https?://[^\s/$.?#].[^\s]*");

            foreach (var cellInfo in cellInfos)
            {
                var range = worksheet.Range[cellInfo.Address];

                // セルに数式が含まれる場合はスキップ
                if (range.HasFormula)
                {
                    continue;
                }

                string value = cellInfo.Value as string;

                // URLの場合、ハイパーリンクを追加
                if (urlRegex.IsMatch(value ?? ""))
                {
                    worksheet.Hyperlinks.Add(range, value);
                }

                range.Value2 = value;
            }
        }

        // セル内の {{*}} を置き換える
        static void ReplaceValues(Excel.Worksheet sheet, Dictionary<string, string> replacements)
        {
            var yamlContent = sheet.GetCustomProperty(indexSheetTemplateCellsCustomPropertyName);

            if (string.IsNullOrEmpty(yamlContent))
            {
                return;
            }

            var replacedYamlContent = ReplacePlaceholders(yamlContent, replacements);
            var deserializer = new DeserializerBuilder()
                       .WithNamingConvention(NullNamingConvention.Instance)
                       .Build();
            var cellInfos = deserializer.Deserialize<List<CellInfo>>(replacedYamlContent);

            SetCellValues(sheet, cellInfos);
        }

        const string templateFileName = "template.xlsm";

        const string indexSheetNameCustomPropertyName = "IndexSheetName";
        const string templateSheetNameCustomPropertyName = "TemplateSheetName";
        const string ssProjectIdCustomPropertyName = "SSProjectId";
        const string sheetIdCustomPropertyName = "SheetId";
        const string sheetHashCustomPropertyName = "SheetHash";
        const string sheetImageHashCustomPropertyName = "SheetImageHash";
        const string indexSheetTemplateCellsCustomPropertyName = "SheetImageHash";

        const string ssSheetRangeName = "SS_SHEET";

        const string noImageFilePath = "images/no_image.jpg";

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
            var values = ExcelExtensions.GetValuesAs2DArray(range.Value2);

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

        class WorkbookInfo
        {
            public string ProjectId { get; set; }
            public string IndexSheetName { get; set; }
            public string TemplateSheetName { get; set; }
            public RenderLog LastRenderLog { get; set; }

            static public WorkbookInfo CreateFromWorkbook(Excel.Workbook workbook)
            {
                string projectId = workbook.GetCustomProperty(ssProjectIdCustomPropertyName);

                if (projectId == null)
                {
                    return null;
                }

                // projectId があるなら他のもあるという前提
                string indexSheetName = workbook.GetCustomProperty(indexSheetNameCustomPropertyName);
                string templateSheetName = workbook.GetCustomProperty(templateSheetNameCustomPropertyName);
                var lastRenderLog = workbook.GetCustomProperty<RenderLog>("RenderLog");
                Debug.Assert(indexSheetName != null, "indexSheetName != null");
                Debug.Assert(templateSheetName != null, "templateSheetName != null");
                Debug.Assert(lastRenderLog != null, "lastRenderLog != null");

                return new WorkbookInfo
                {
                    ProjectId = projectId,
                    IndexSheetName = indexSheetName,
                    TemplateSheetName = templateSheetName,
                    LastRenderLog = lastRenderLog,
                };
            }
        }

        class SheetValuesInfo
        {
            SheetAddressInfo sheetAddressInfo;

            public Excel.Range Range { get; set; }
            public object[,] Values { get; set; }
            public IEnumerable<object> Ids { get; set; }

            public HashSet<int> IgnoreColumnOffsets
            {
                get
                {
                    return sheetAddressInfo.RangeInfo.IgnoreColumnOffsets;
                }
            }

            public int RangeWidth
            {
                get
                {
                    return Range.Columns.Count;
                }
            }

            // idValues を key にした行（List<object>）の dictionary を作る
            public Dictionary<string, List<object>> RowDictionaryWithIDKeys
            {
                get
                {
                    return CreateRowDictionaryWithIDKeys(Values, Ids);
                }
            }

            public static SheetValuesInfo CreateFromSheet(Excel.Worksheet sheet)
            {
                var sheetAddressInfo = GetSheetAddressInfo(sheet);
                var sheetRange = GetRange(sheet, sheetAddressInfo);
                var sheetValues = GetValues(sheet, sheetAddressInfo);
                var sheetValueIds = GetIds(sheet, sheetAddressInfo);

                return new SheetValuesInfo()
                {
                    sheetAddressInfo = sheetAddressInfo,
                    Range = sheetRange,
                    Values = sheetValues,
                    Ids = sheetValueIds,
                };
            }

        }

        async Task ForceUpdateSheet(Excel.Worksheet sheet)
        {
            Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook workbook = sheet.Parent as Excel.Workbook;

            // 作ったシートも元のシートと同じ状態にする
            var activeCellPosition = excelApp.GetActiveCellPosition();
            var scrollPosition = excelApp.GetScrollPosition();
            var activeSheetZoom = excelApp.GetActiveSheetZoom();

            WorkbookInfo workbookInfo = WorkbookInfo.CreateFromWorkbook(workbook);

            if (workbookInfo == null)
            {
                string projectName = Assembly.GetExecutingAssembly().GetName().Name;
                MessageBox.Show($"{projectName} で生成されたブックではありません。");
                return;
            }

            string projectId = workbookInfo.ProjectId;

            // lastRenderLog.User が今のユーザーと異なる、もしくは lastRenderLog.SourceFilePath が見つからない場合、前回生成時の環境と異なるとみなしてファイル選択させる
            var lastRenderLog = workbookInfo.LastRenderLog;
            bool isSameUser = lastRenderLog.User == Environment.UserName;
            bool sourceFileExists = File.Exists(lastRenderLog.SourceFilePath);
            string jsonFilePath;

            if (!isSameUser || !sourceFileExists)
            {
                string message = null;
                if (!isSameUser)
                {
                    message = "最後に更新された環境と異なります。";
                }
                else if (!sourceFileExists)
                {
                    message = "ソースファイルが見つかりません。";
                }

                DialogResult fileSelectionResult = MessageBox.Show($"{message}\nProject ID が「{projectId}」のソースファイルを選択し直してください。", "確認", MessageBoxButtons.OKCancel);
                if (fileSelectionResult != DialogResult.OK)
                {
                    return;
                }

                jsonFilePath = OpenSourceFile();
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

            string activeSheetId = sheet.GetCustomProperty(sheetIdCustomPropertyName);

            // 今開いているシートの id を index sheet から取得
            var indexSheet = workbook.Sheets[workbookInfo.IndexSheetName] as Excel.Worksheet;
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
            if (activeSheetId == null)  // XXX: 古いバージョンの対応。ある程度稼働したら削除
            {
                activeSheetId = sheetIds.ElementAt(sheetIndex).ToString();
            }

            JsonArray sheetNodes = jsonObject["children"].AsArray();

            // jsonObject から同じ id の node を取得
            JsonNode targetSheetNode = null;
            foreach (JsonNode sheetNode in sheetNodes)
            {
                string id = sheetNode["id"].ToString();
                if (id == activeSheetId)
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

            // ファイルが変更されて保存されてない場合
            if (workbook.Saved == false)
            {
                // 保存確認ダイアログを表示
                DialogResult yesNoCancel = MessageBox.Show($"シートを更新する前にファイルの変更内容を保存しますか？", "確認", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);

                switch (yesNoCancel)
                {
                    case DialogResult.Yes:
                        workbook.Save();
                        break;
                    case DialogResult.No:
                        // 保存せずに続行
                        break;
                    case DialogResult.Cancel:
                        // キャンセルされた場合、処理を中止
                        return;
                }
            }

            if (newSheetName != sheetName)
            {
                // シート名が変わっていたら index sheet にも反映
                sheetNameRange.Cells[1 + sheetIndex].Value2 = newSheetName;
            }

            // シートが非表示の場合、コピーしてもシートがアクティブにならないので、一時的に表示状態にする
            Excel.Worksheet templateSheet = workbook.Sheets[workbookInfo.TemplateSheetName];
            var visible = templateSheet.Visible;
            templateSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            templateSheet.Copy(After: sheet);
            templateSheet.Visible = visible;

            // コピーされたシートはアクティブシートになるので、それを取得
            Excel.Worksheet newSheet = (Excel.Worksheet)templateSheet.Application.ActiveSheet;

            // 元のシートから今の入力内容を取り込む
            SheetValuesInfo sheetValuesInfo = SheetValuesInfo.CreateFromSheet(sheet);

            // 元のシートを削除
            excelApp.DisplayAlerts = false;
            sheet.Delete();
            excelApp.DisplayAlerts = true;

            // 元のシートと同じ名前でも良いように元シート削除後に名前変更
            newSheet.Name = newSheetName;

            // シート作成
            // node, 画像ファイルの比較はしない
            var missingImagePathsInSheet = RenderSheet(targetSheetNode, confData, jsonFilePath, newSheet, sheetValuesInfo);

            // シートの JsonNode の hash をカスタムプロパティに保存
            // XXX: hash には text を含めたくないので、hashを求める前に一時的に削除
            targetSheetNode.AsObject().Remove("text");
            string newSheetHash = targetSheetNode.ComputeSha256();
            //targetSheetNode["text"] = newSheetName;
            newSheet.SetCustomProperty(sheetHashCustomPropertyName, newSheetHash);

            // await の前にWindowsFormsSynchronizationContextを設定
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }

            var newSheetImageHash = await ComputeImagesHash(jsonFilePath, targetSheetNode);
            newSheet.SetCustomProperty(sheetImageHashCustomPropertyName, newSheetImageHash);

            // シートを元の状態と同じにする
            newSheet.Activate();
            excelApp.SetActiveCellPosition(activeCellPosition);

            var originalZoom = excelApp.ActiveWindow.Zoom;
            if (originalZoom == activeSheetZoom)
            {
                excelApp.ActiveWindow.Zoom = originalZoom + 1; // ズームレベルを一時的に変更
            }
            excelApp.ScreenUpdating = true;
            excelApp.SetActiveSheetZoom(activeSheetZoom);   // scroll より後に zoom をセットすると微妙にずれるっぽい
            excelApp.SetScrollPosition(scrollPosition);

            if (missingImagePathsInSheet.Any())
            {
                SynchronizationContext.Current.Post(_ => {
                    ShowMissingImageFilesDialog(missingImagePathsInSheet);
                }, null);
            }

            // TODO: RenderLog 書き出す処理を共通化
            RenderLog renderLog = new RenderLog
            {
                SourceFilePath = jsonFilePath,
                User = Environment.UserName
            };
            workbook.SetCustomProperty("RenderLog", renderLog);
        }

        public async void OnUpdateCurrentSheetButtonPressed(IRibbonControl control)
        {
            Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
            var sheet = excelApp.ActiveSheet as Excel.Worksheet;

            if (sheet == null)
            {
                MessageBox.Show($"アクティブなシートがありません。");
                return;
            }

            excelApp.ScreenUpdating = false;
            excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
            excelApp.EnableEvents = false;
            MacroControl.DisableMacros(excelApp);

            await ForceUpdateSheet(sheet);

            excelApp.StatusBar = false;
            excelApp.ScreenUpdating = true;
            excelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            excelApp.EnableEvents = true;
            MacroControl.EnableMacros(excelApp);

        }

        static SortedSet<string> CollectImageFilePaths(JsonNode node)
        {
            var imageFilePaths = new SortedSet<string>();

            void TraverseForImagePaths(JsonNode currentNode)
            {
                if (currentNode == null)
                {
                    return;
                }

                string imageFilePath = currentNode["imageFilePath"]?.GetValue<string>();

                if (imageFilePath != null)
                {
                    imageFilePaths.Add(imageFilePath);
                }

                if (currentNode["children"] is JsonArray children)
                {
                    foreach (JsonNode child in children)
                    {
                        TraverseForImagePaths(child);
                    }
                }
            }

            TraverseForImagePaths(node);

            return imageFilePaths;
        }

        // 画像のhashを計算
        // filepath順にソートしてhashを改行で連結
        // 差分検出用
        async Task<string> ComputeImagesHash(string jsonFilePath, JsonNode sheetNode)
        {
            var imageFilePaths = CollectImageFilePaths(sheetNode);

            if (!imageFilePaths.Any())
            {
                return null;
            }

            var tasks = imageFilePaths.Select(imagePath => Task.Run(() =>
            {
                string path = GetAbsolutePathFromBasePath(jsonFilePath, imagePath);

                if (!File.Exists(path))
                {
                    return new
                    {
                        Path = imagePath,
                        //AbsolutePath = path,
                        Hash = "no_image"
                    };
                }

                using (var sha256 = SHA256.Create())
                {
                    using (var stream = File.OpenRead(path))
                    {
                        return new
                        {
                            Path = imagePath,
                            //AbsolutePath = path,
                            Hash = BitConverter.ToString(sha256.ComputeHash(stream)).Replace("-", "").ToLower()
                        };
                    }
                }
            })).ToArray();

            await Task.WhenAll(tasks);

            //var results = tasks.Select(t => t.Result).ToList();
            return string.Join("\n", tasks.Select(t => t.Result.Hash));
        }

        string ComputeSheetHash(JsonNode sheetNode)
        {
            // XXX: hash には text を含めたくないので、hashを求める前に削除
            JsonNode clonedNode = sheetNode.DeepClone();

            clonedNode.AsObject().Remove("text");
            return clonedNode.ComputeSha256();
        }

        async Task UpdateAllSheets(Excel.Workbook workbook)
        {
            Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;

            WorkbookInfo workbookInfo = WorkbookInfo.CreateFromWorkbook(workbook);

            if (workbookInfo == null)
            {
                string projectName = Assembly.GetExecutingAssembly().GetName().Name;
                MessageBox.Show($"{projectName} で生成されたブックではありません。");
                return;
            }

            string projectId = workbookInfo.ProjectId;

            // lastRenderLog.User が今のユーザーと異なる、もしくは lastRenderLog.SourceFilePath が見つからない場合、前回生成時の環境と異なるとみなしてファイル選択させる
            var lastRenderLog = workbookInfo.LastRenderLog;
            bool isSameUser = lastRenderLog.User == Environment.UserName;
            bool sourceFileExists = File.Exists(lastRenderLog.SourceFilePath);
            string jsonFilePath;

            if (!isSameUser || !sourceFileExists)
            {
                string message = null;
                if (!isSameUser)
                {
                    message = "最後に更新された環境と異なります。";
                }
                else if (!sourceFileExists)
                {
                    message = "ソースファイルが見つかりません。";
                }

                DialogResult fileSelectionResult = MessageBox.Show($"{message}\nProject ID が「{projectId}」のソースファイルを選択し直してください。", "確認", MessageBoxButtons.OKCancel);
                if (fileSelectionResult != DialogResult.OK)
                {
                    return;
                }

                jsonFilePath = OpenSourceFile();
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

            string indexSheetName = workbook.GetCustomProperty(indexSheetNameCustomPropertyName);
            string templateSheetName = workbook.GetCustomProperty(templateSheetNameCustomPropertyName);

            if (indexSheetName == null)
            {
                MessageBox.Show($"カスタムプロパティに {indexSheetNameCustomPropertyName} が設定されていません。");
                return;
            }
            if (templateSheetName == null)
            {
                MessageBox.Show($"カスタムプロパティに {templateSheetNameCustomPropertyName} が設定されていません。");
                return;
            }

            // ファイルが変更されて保存されてない場合
            if (workbook.Saved == false)
            {
                // 保存確認ダイアログを表示
                DialogResult yesNoCancel = MessageBox.Show($"シートを更新する前にファイルの変更内容を保存しますか？", "確認", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);

                switch (yesNoCancel)
                {
                    case DialogResult.Yes:
                        workbook.Save();
                        break;
                    case DialogResult.No:
                        // 保存せずに続行
                        break;
                    case DialogResult.Cancel:
                        // キャンセルされた場合、処理を中止
                        return;
                }
            }

            // await の前にWindowsFormsSynchronizationContextを設定
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }
            var context = SynchronizationContext.Current;

            excelApp.DisplayAlerts = false;
            excelApp.ScreenUpdating = false;
            excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
            excelApp.EnableEvents = false;
            excelApp.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            // 作ったシートも元のシートと同じ状態にする
            var activeCellPosition = excelApp.GetActiveCellPosition();
            var scrollPosition = excelApp.GetScrollPosition();
            var activeSheetZoom = excelApp.GetActiveSheetZoom();
            var activeSheet = workbook.ActiveSheet as Excel.Worksheet;
            string activeSheetId = activeSheet.GetCustomProperty(sheetIdCustomPropertyName);
            string originalActiveSheetName = activeSheet.Name;

            // 今開いている book の id を index sheet から取得
            var indexSheet = workbook.Sheets[workbookInfo.IndexSheetName] as Excel.Worksheet;
            var sheetIds = GetSheetIdsFromIndexSheet(indexSheet).Select(item => item.ToString()).ToList();
            var sheetNameRange = GetSheetNamesRangeFromIndexSheet(indexSheet);
            var sheetNames = sheetNameRange.GetColumnValues(0).Select(item => item.ToString()).ToList();
            Dictionary<string, string> originalSheetNamesById = sheetIds.Zip(sheetNames, (id, name) => new { id, name })
                                                                         .ToDictionary(x => x.id, x => x.name);

            // 特定のプロパティ（items）を配列としてアクセス
            JsonArray sheetNodes = jsonObject["children"].AsArray();

            Dictionary<string, string> newSheetNamesById = sheetNodes.ToDictionary(
                item => item["id"].ToString(),
                item => item["text"].ToString()
            );

            // シートを削除をするので警告を出さないように
            excelApp.DisplayAlerts = false;

            // originalSheetNamesByIdから、idがnewSheetNamesByIdに含まれないものをフィルタリング
            IEnumerable<string> sheetsToRemoveIds = originalSheetNamesById.Keys
                .Where(key => !newSheetNamesById.ContainsKey(key));

            // シート削除
            foreach (string id in sheetsToRemoveIds)
            {
                string sheetName = originalSheetNamesById[id];
                workbook.Sheets[sheetName].Delete();
            }

            // originalSheetNamesByIdからも削除
            // originalSheetNamesByIdから、idがnewSheetNamesByIdに含まれるものをフィルタリング
            originalSheetNamesById = originalSheetNamesById
                .Where(kvp => newSheetNamesById.ContainsKey(kvp.Key))
                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            Dictionary<string, Excel.Worksheet> originalSheetsById = new Dictionary<string, Excel.Worksheet>();
            var sheetsToRename = new List<(Excel.Worksheet Sheet, string NewName, string Id)>();

            // originalSheetNamesById や originalSheetNamesById.Keys を foreach で回すと、値の変更で例外投げるので Keys.List() の foreach で回避
            foreach (var id in originalSheetNamesById.Keys.ToList())
            {
                string sheetName = originalSheetNamesById[id];
                Excel.Worksheet sheet = workbook.Sheets[sheetName];
                string newSheetName = newSheetNamesById[id];

                if (newSheetName != sheetName)
                {
                    if (!originalSheetNamesById.Values.Contains(newSheetName))
                    {
                        // newSheetName が一意なら直接リネーム
                        sheet.Name = newSheetName;
                        originalSheetNamesById[id] = newSheetName; // originalSheetNamesByIdも更新
                    }
                    else
                    {
                        // 一時的な名前を生成
                        // 変更後も重複しない名前にする
                        string tempName = $"{sheetName}_temp";
                        int counter = 1;
                        while (originalSheetNamesById.Values.Contains(tempName)
                            || newSheetNamesById.Values.Contains(tempName))
                        {
                            tempName = $"{sheetName}_temp{counter++}";
                        }
                        sheetsToRename.Add((sheet, newSheetName, id));
                        sheet.Name = tempName; // 一時的な名前に変更
                        originalSheetNamesById[id] = tempName; // originalSheetNamesByIdも更新
                    }
                }

                originalSheetsById.Add(id, sheet);
            }

            // すべてのシートが一時的な名前に変更された後、最終的な名前にリネーム
            foreach (var renameInfo in sheetsToRename)
            {
                string newSheetName = renameInfo.NewName;

                renameInfo.Sheet.Name = newSheetName;
                originalSheetNamesById[renameInfo.Id] = newSheetName; // originalSheetNamesByIdも更新
            }

            var missingImagePaths = new List<(string filePath, string sheetName, string address)>();

            // シートが非表示の場合、コピーしてもシートがアクティブにならないので、一時的に表示状態にする
            Excel.Worksheet templateSheet = workbook.Sheets[workbookInfo.TemplateSheetName];
            var templateSheetVisible = templateSheet.Visible;
            templateSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            // +1 は index シート
            progressBarForm = new ProgressBarForm(sheetNodes.Count + 1);
            progressBarForm.Show();

            // シートの JsonNode の hash をカスタムプロパティに保存
            Dictionary<string, Task<string>> newSheetHashTasks = new Dictionary<string, Task<string>>();
            foreach (JsonNode sheetNode in sheetNodes)
            {
                string id = sheetNode["id"].ToString();
                var task = Task.Run(() => ComputeSheetHash(sheetNode));

                newSheetHashTasks.Add(id, task);
            }

            await Task.Run(async () =>
            {
                foreach (JsonNode sheetNode in sheetNodes)
                {
                    string newSheetName = sheetNode["text"].ToString();
                    string id = sheetNode["id"].ToString();

                    // プログレスバーを更新
                    progressBarForm.Invoke(new Action<string>(progressBarForm.UpdateSheetName), newSheetName);

                    // 既存のシート
                    if (originalSheetsById.ContainsKey(id))
                    {
                        // 画像の hash 計算を開始しておく
                        var newSheetImageHashTask = ComputeImagesHash(jsonFilePath, sheetNode);
                        var sheet = originalSheetsById[id];
                        var sheetHash = sheet.GetCustomProperty(sheetHashCustomPropertyName);
                        string newSheetImageHash;

                        // 元データが同じで、画像も変更されていないなら生成しない
                        string newSheetHash = await newSheetHashTasks[id];
                        if (newSheetHash == sheetHash)
                        {
                            var sheetImageHash = sheet.GetCustomProperty(sheetImageHashCustomPropertyName);
                            newSheetImageHash = await newSheetImageHashTask;

                            if (newSheetImageHash == sheetImageHash)
                            {
                                continue;
                            }
                        }

                        // シートをコピー
                        templateSheet.Copy(After: sheet);
                        Excel.Worksheet newSheet = workbook.Sheets[sheet.Index + 1];

                        // 元のシートから今の入力内容を取り込む
                        SheetValuesInfo sheetValuesInfo = SheetValuesInfo.CreateFromSheet(sheet);

                        // 元のシートを削除
                        sheet.Delete();
                        newSheet.Name = newSheetName;

                        var missingImagePathsInSheet = RenderSheet(sheetNode, confData, jsonFilePath, newSheet, sheetValuesInfo);

                        missingImagePaths.AddRange(missingImagePathsInSheet);
                        newSheet.SetCustomProperty(sheetHashCustomPropertyName, newSheetHash);
                        newSheetImageHash = await newSheetImageHashTask;
                        newSheet.SetCustomProperty(sheetImageHashCustomPropertyName, newSheetImageHash);
                    }
                    else
                    {
                        // 新規シート作成

                        // 画像の hash 計算を開始しておく
                        var newSheetImageHashTask = ComputeImagesHash(jsonFilePath, sheetNode);

                        // シートをコピー
                        // 一旦は最後に追加。最後にまとめて並び替える
                        var beforeSheet = workbook.Sheets[workbook.Sheets.Count];
                        templateSheet.Copy(After: beforeSheet);
                        Excel.Worksheet newSheet = workbook.Sheets[beforeSheet.Index + 1];
                        newSheet.Name = newSheetName;

                        var missingImagePathsInSheet = RenderSheet(sheetNode, confData, jsonFilePath, newSheet, null);

                        missingImagePaths.AddRange(missingImagePathsInSheet);

                        string newSheetHash = await newSheetHashTasks[id];
                        newSheet.SetCustomProperty(sheetHashCustomPropertyName, newSheetHash);

                        var newSheetImageHash = await newSheetImageHashTask;
                        newSheet.SetCustomProperty(sheetImageHashCustomPropertyName, newSheetImageHash);
                    }
                }

                templateSheet.Visible = templateSheetVisible;

                // シートの並び順修正
                // リストに従ってシートを後ろに詰める
                List<string> sheetNamesInOrder = sheetNodes.Select(item => item["text"].ToString()).ToList();
                for (int i = 0; i < sheetNamesInOrder.Count; i++)
                {
                    Excel.Worksheet sheetToMove = workbook.Sheets[sheetNamesInOrder[i]];
                    int targetIndex = workbook.Sheets.Count - (sheetNamesInOrder.Count - 1 - i);

                    // シートが既に正しい位置にない場合のみ移動
                    if (sheetToMove.Index != targetIndex)
                    {
                        sheetToMove.Move(Type.Missing, workbook.Sheets[targetIndex]);
                    }
                }

                // プログレスバーを更新
                progressBarForm.Invoke(new Action<string>(progressBarForm.UpdateSheetName), indexSheetName);

                // 元のシートから今の入力内容を取り込む
                SheetValuesInfo indexSheetValuesInfo = SheetValuesInfo.CreateFromSheet(indexSheet);

                RenderIndexSheet(sheetNodes, confData, indexSheet, indexSheetValuesInfo);

                if (activeSheetId != null)
                {
                    if (newSheetNamesById.ContainsKey(activeSheetId))
                    {
                        // シートを元の状態と同じにする
                        var originalActiveSheet = workbook.Sheets[newSheetNamesById[activeSheetId]];

                        originalActiveSheet.Activate();
                        excelApp.SetActiveCellPosition(activeCellPosition);
                        excelApp.SetActiveSheetZoom(activeSheetZoom);   // scroll より後に zoom をセットすると微妙にずれるっぽい
                        excelApp.SetScrollPosition(scrollPosition);
                    }
                    else
                    {
                        // とりあえず index sheet を選択しておく
                        indexSheet.Activate();
                    }
                }
                else
                {
                    var originalActiveSheet = workbook.Sheets[originalActiveSheetName];

                    originalActiveSheet.Activate();

                    if (originalActiveSheetName == indexSheetName)
                    {
                        // index シートならシートを元の状態と同じにする
                        excelApp.SetActiveCellPosition(activeCellPosition);
                        excelApp.SetActiveSheetZoom(activeSheetZoom);   // scroll より後に zoom をセットすると微妙にずれるっぽい
                        excelApp.SetScrollPosition(scrollPosition);
                    }
                }

                // 処理が完了したらフォームを閉じる
                progressBarForm.Invoke(new Action(progressBarForm.CloseForm));
            });

            progressBarForm.Close();

            excelApp.DisplayAlerts = true;

            var originalZoom = excelApp.ActiveWindow.Zoom;
            excelApp.ActiveWindow.Zoom = originalZoom + 1; // ズームレベルを一時的に変更
            excelApp.ScreenUpdating = true;
            excelApp.ActiveWindow.Zoom = originalZoom; // 元に戻す
            excelApp.SetScrollPosition(scrollPosition);

            if (missingImagePaths.Any())
            {
                context.Post(_ => {
                    ShowMissingImageFilesDialog(missingImagePaths);
                }, null);
            }
        }

        static bool IsSameNameWorkbookOpen(string fileName)
        {
            Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;

            // ファイル名の拡張子を取り除く
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);

            // すべての開いているブックをチェック
            foreach (Excel.Workbook wb in excelApp.Workbooks)
            {
                // ファイル名の比較（拡張子を除いた名前で比較）
                if (Path.GetFileNameWithoutExtension(wb.FullName).Equals(fileNameWithoutExtension, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        async Task CreateNewWorkbook()
        {
            DialogResult fileSelectionResult = MessageBox.Show($"ファイルを新規作成します。\nソースファイルを選択してください。", "確認", MessageBoxButtons.OKCancel);
            if (fileSelectionResult != DialogResult.OK)
            {
                return;
            }

            var jsonFilePath = OpenSourceFile();
            // キャンセルされたら何もしない
            if (jsonFilePath == null)
            {
                return;
            }

            string jsonString = File.ReadAllText(jsonFilePath);

            JsonNode jsonObject = JsonNode.Parse(jsonString);
            var confData = GetPropertiesFromJsonNode(jsonObject, "variables");

            const string outputFilenameConfName = "outputFilename";

            //string newFilePath = GetFilePathWithoutExtension(jsonFilePath);
            string newFileName = confData.ContainsKey(outputFilenameConfName) ? confData[outputFilenameConfName] : Path.GetFileNameWithoutExtension(jsonFilePath);

            if (IsSameNameWorkbookOpen(newFileName))
            {
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(newFileName);
                MessageBox.Show($"'{fileNameWithoutExtension}'と同じ名前のファイルが既に開かれています。\nファイルを閉じてから再度実行してください。");
                return;
            }

            string jsonFileDirectory = Path.GetDirectoryName(jsonFilePath);
            string newFilePath = Path.Combine(jsonFileDirectory, newFileName);

            Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;

            string templateFilePath = GetAbsolutePathFromExecutingDirectory(templateFileName);

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
            excelApp.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            // 特定のプロパティ（items）を配列としてアクセス
            JsonArray sheetNodes = jsonObject["children"].AsArray();

            List<string> sheetNames = new List<string>();
            Excel.Worksheet sheet = workbook.Sheets[templateSheetName];

            List<(string filePath, string sheetName, string address)> missingImagePaths = new List<(string filePath, string sheetName, string address)>();

            // +1 は index シート
            progressBarForm = new ProgressBarForm(sheetNodes.Count + 1);
            progressBarForm.Show();

            await Task.Run(async () =>
            {
                foreach (JsonNode sheetNode in sheetNodes)
                {
                    // 画像の hash 計算を開始しておく
                    var newSheetImageHashTask = ComputeImagesHash(jsonFilePath, sheetNode);

                    // シートの JsonNode の hash 計算を開始しておく
                    var sheetHashTask = Task.Run(() => ComputeSheetHash(sheetNode));

                    string newSheetName = sheetNode["text"].ToString();

                    // プログレスバーを更新
                    progressBarForm.Invoke(new Action<string>(progressBarForm.UpdateSheetName), newSheetName);

                    // シートをコピーしてリネーム
                    sheet.Copy(After: workbook.Sheets[workbook.Sheets.Count]);
                    Excel.Worksheet newSheet = workbook.Sheets[workbook.Sheets.Count];
                    newSheet.Name = newSheetName;

                    var missingImagePathsInSheet = RenderSheet(sheetNode, confData, jsonFilePath, newSheet, null);

                    // シートの JsonNode の hash をカスタムプロパティに保存
                    string sheetHash = await sheetHashTask;
                    newSheet.SetCustomProperty(sheetHashCustomPropertyName, sheetHash);

                    var newSheetImageHash = await newSheetImageHashTask;
                    newSheet.SetCustomProperty(sheetImageHashCustomPropertyName, newSheetImageHash);

                    missingImagePaths.AddRange(missingImagePathsInSheet);

                    sheetNames.Add(newSheetName);
                }

                // プログレスバーを更新
                progressBarForm.Invoke(new Action<string>(progressBarForm.UpdateSheetName), indexSheetName);

                Excel.Worksheet indexSheet = workbook.Sheets[indexSheetName];

                // 新規作成時は index sheet のテンプレセルの情報を保存しておく
                var indexSheetTemplateCells = GetTemplateCells(indexSheet);
                var serializer = new SerializerBuilder()
                    .WithNamingConvention(NullNamingConvention.Instance)
                    .Build();
                var indexSheetTemplateCellsYaml = serializer.Serialize(indexSheetTemplateCells);

                indexSheet.SetCustomProperty(indexSheetTemplateCellsCustomPropertyName, indexSheetTemplateCellsYaml);

                RenderIndexSheet(sheetNodes, confData, indexSheet, null);

                // 最後にindexシートを選択状態にしておく
                indexSheet.Activate();

                // 処理が完了したらフォームを閉じる
                progressBarForm.Invoke(new Action(progressBarForm.CloseForm));
            });

            progressBarForm.Close();

            var originalZoom = excelApp.ActiveWindow.Zoom;

            excelApp.EnableEvents = true;
            excelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            excelApp.ActiveWindow.Zoom = originalZoom + 1;
            excelApp.ScreenUpdating = true;
            excelApp.ActiveWindow.Zoom = originalZoom;
            excelApp.DisplayAlerts = true;
            excelApp.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityByUI;

            if (missingImagePaths.Any())
            {
                ShowMissingImageFilesDialog(missingImagePaths);
            }

            RenderLog renderLog = new RenderLog
            {
                SourceFilePath = jsonFilePath,
                User = Environment.UserName
            };
            workbook.SetCustomProperty("RenderLog", renderLog);

            //var lastRenderLog = workbook.GetCustomProperty<RenderLog>("RenderLog");

            string projectId = confData["project"];
            workbook.SetCustomProperty(ssProjectIdCustomPropertyName, projectId);

            workbook.Save();
        }

        public async void OnCreateNewButtonPressed(IRibbonControl control)
        {
            Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;

            await CreateNewWorkbook();

            excelApp.EnableEvents = true;
            if (excelApp.ActiveWorkbook != null)
            {
                excelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            }
            excelApp.ScreenUpdating = true;
            excelApp.DisplayAlerts = true;
            excelApp.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityByUI;
        }

        public async void OnRenderButtonPressed(IRibbonControl control)
        {
            Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
            var sheet = excelApp.ActiveSheet as Excel.Worksheet;

            // アクティブなシートがない時は新規作成
            if (sheet == null)
            {
                await CreateNewWorkbook();
            }
            else
            {
                Excel.Workbook workbook = sheet.Parent as Excel.Workbook;

                await UpdateAllSheets(workbook);
            }

            excelApp.EnableEvents = true;
            if (excelApp.ActiveWorkbook != null)
            {
                excelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            }
            excelApp.ScreenUpdating = true;
            excelApp.DisplayAlerts = true;
            excelApp.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityByUI;
        }

        static void RenderIndexSheet(IEnumerable<JsonNode> sheetNodes, Dictionary<string, string> confData, Excel.Worksheet dstSheet, SheetValuesInfo sheetValuesInfo = null)
        {
            var sheetNameListRange = dstSheet.GetNamedRange("SS_SHEETNAMELIST").RefersToRange;

            int indexStartRow = sheetNameListRange.Row;
            int indexRowCount = sheetNameListRange.Rows.Count;
            int indexEndRow = sheetNameListRange.Rows[indexRowCount].Row;
            int indexStartColumn = sheetNameListRange.Column;
            string idColumnAddress = "T";
            int idColumn = dstSheet.ColumnAddressToIndex(idColumnAddress);

            string syncStartColumnAddress = "Q";
            int syncStartColumn = dstSheet.ColumnAddressToIndex(syncStartColumnAddress);
            int syncStartColumnCount = 1;
            int[] syncIgnoreColumnOffsets = { };

            var sheetNames = ExtractPropertyValues(sheetNodes, "text");
            var sheetNamesCount = sheetNames.Count();

            // 行が足りなければ挿入
            if (sheetNamesCount > indexRowCount)
            {
                int numberOfRows = sheetNamesCount - indexRowCount;

                dstSheet.InsertRowsAndCopyFormulas(indexStartRow + 1, numberOfRows);
            }
            // 多ければ削除
            else if (sheetNamesCount < indexRowCount)
            {
                int numberOfRowsToDelete = indexRowCount - sheetNamesCount;

                dstSheet.DeleteRows(indexStartRow, numberOfRowsToDelete);
            }

            dstSheet.SetValueInSheetAsColumn(indexStartRow, indexStartColumn, sheetNames);

            // テンプレ処理
            ReplaceValues(dstSheet, confData);

            // 幅をautofit
            dstSheet.Cells[indexStartRow, indexStartColumn].Resize(sheetNamesCount).Columns.AutoFit();

            // 適当な位置に列挿入して ID を入れて非表示にする
            var ids = ExtractPropertyValues(sheetNodes, "id");
            Excel.Range column = dstSheet.Columns[idColumn];
            // XXX: id列が非表示ならすでにid列が存在するとみなして上書きする。どうせこの仕様は廃止したいので一旦これで
            if (!column.EntireColumn.Hidden)
            {
                column.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            }
            dstSheet.SetValueInSheet(indexStartRow, idColumn, ids, false);
            Excel.Range idColumnRange = dstSheet.Columns[idColumn];
            idColumnRange.EntireColumn.Hidden = true;

            // 名前付き範囲として追加
            var rangeforNamedRange = dstSheet.GetRange(indexStartRow, syncStartColumn, sheetNamesCount, syncStartColumnCount);
            var namedRange = dstSheet.Names.Add(Name: ssSheetRangeName, RefersTo: rangeforNamedRange);
            RangeInfo rangeInfo = new RangeInfo
            {
                IdColumnOffset = idColumn - syncStartColumn,
                IgnoreColumnOffsets = new HashSet<int>(syncIgnoreColumnOffsets),
            };
            var serializer = new SerializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();

            namedRange.Comment = serializer.Serialize(rangeInfo);

            if (sheetValuesInfo != null)
            {
                SheetValuesInfo dstSheetValuesInfo = SheetValuesInfo.CreateFromSheet(dstSheet);
                var ignoreColumnOffsets = dstSheetValuesInfo.IgnoreColumnOffsets;
                Debug.Assert(AreHashSetsEqual(ignoreColumnOffsets, sheetValuesInfo.IgnoreColumnOffsets), "AreHashSetsEqual(ignoreColumnOffsets, sheetAddressInfo.RangeInfo.IgnoreColumnOffsets)");

                // idValues を key にした行（List<object>）の dictionary を作る
                var valuesDictionary = sheetValuesInfo.RowDictionaryWithIDKeys;

                // dstSheet の Values のコピーを作って、元のシートの Values から id を基に上書きコピーする
                // idが見つからない行、ignoreColumn は何もしないので、dstSheet のものが採用される
                var values = CopyValuesById(dstSheetValuesInfo.Values, dstSheetValuesInfo.Ids, valuesDictionary, ignoreColumnOffsets);

                dstSheetValuesInfo.Range.Value2 = values;
            }
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
        public static class MacroControl
        {
            public static void DisableMacros(Excel.Application excelApp)
            {
                excelApp.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            }

            public static void EnableMacros(Excel.Application excelApp)
            {
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

        IEnumerable<(string filePath, string sheetName, string address)> RenderSheet(JsonNode sheetNode, Dictionary<string, string> confData, string jsonFilePath, Excel.Worksheet dstSheet, SheetValuesInfo sheetValuesInfo = null)
        {
            List<(string filePath, string sheetName, string address)> missingImagePaths = new List<(string filePath, string sheetName, string address)>();

            // シートの JsonNode の hash をカスタムプロパティに保存
            // XXX: hash には text を含めたくないので、hashを求める前に一時的に削除
            var newSheetName = sheetNode["text"].ToString();

            //templateSheet.SetCustomProperty(sheetHashCustomPropertyName, newSheetHash);

            List<JsonNode> leafNodes;
            int maxDepth;
            List<List<NodeData>> result = TraverseTreeFromRoot(sheetNode, out leafNodes, out maxDepth);
            int leafCount = leafNodes.Count;

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
                Excel.Range startColumnRange = dstSheet.Columns[endColumn];

                // 列を挿入
                startColumnRange.Resize[1, numberOfColumns].EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            }

            if (leafCount > rowHeight)
            {
                int numberOfRows = leafCount - rowHeight;

                dstSheet.InsertRowsAndCopyFormulas(endRow, numberOfRows);
            }
            else if (leafCount < rowHeight)
            {
                int numberOfRowsToDelete = rowHeight - leafCount;

                dstSheet.DeleteRows(startRow, numberOfRowsToDelete);
            }

            // XXX: 深度が列より少ない場合は今のテンプレに合わせていろいろやる…
            if (maxDepth < columnWidth)
            {
                // とりあえず E 列は削除
                Excel.Range column = dstSheet.Columns[5];
                column.Delete();
                columnWidth--;

                // C, D 列を削除するのはいろいろ面倒な作りのようなので、
                // 貼り付け列を調整してお茶を濁す（C, D 列は空欄にする）
                startColumn += columnWidth - maxDepth;
            }

            dstSheet.SetValueInSheet(startRow, startColumn, arrayResult);

            dstSheet.GetRange(startRow, startColumn, leafCount, maxDepth).AutoFitColumnsIfNarrower();

            // 「チェック予定日」列に無条件で START_DATE を入れる
            int dateColumnOffset = initialDateColumn - initialResultColumn;
            int dateColumn = startColumn + maxDepth + dateColumnOffset;
            string dateString = confData["START_DATE"];
            // 文字列をDateTimeに変換
            if (DateTime.TryParse(dateString, out DateTime dateValue))
            {
                dstSheet.SetRangeValue(startRow, dateColumn, leafCount, 1, dateValue);
            }

            // 「チェック予定日」の右隣の列に ID を入れて非表示にする
            int idColumnOffset = initialIdColumn - initialResultColumn;
            int idColumn = startColumn + maxDepth + idColumnOffset;
            var ids = ExtractPropertyValues(leafNodes, "id");
            dstSheet.SetValueInSheet(startRow, idColumn, ids, false);
            Excel.Range idColumnRange = dstSheet.Columns[idColumn];
            idColumnRange.EntireColumn.Hidden = true;

            // 「チェック結果」列に node の initialValues.result を入れる
            int resultColumnOffset = initialResultColumn - initialResultColumn;
            int resultColumn = startColumn + maxDepth + resultColumnOffset;
            var results = ExtractPropertyValuesFromInitialValues(leafNodes, "result");
            dstSheet.SetValueInSheet(startRow, resultColumn, results, false);

            // 「計画時間(分)」列に node の initialValues.estimated_time を入れる
            int planColumnOffset = initialPlanColumn - initialResultColumn;
            int planColumn = startColumn + maxDepth + planColumnOffset;
            var estimatedTimes = ExtractPropertyValuesFromInitialValues(leafNodes, "estimated_time");
            dstSheet.SetValueInSheet(startRow, planColumn, estimatedTimes, false);

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

                    void ApplyCommentCell(int cellColorIndex, int fontColorIndex)
                    {
                        // ここより右のセルの色を変える
                        var cells = dstSheet.Cells[startRow + i, startColumn + j];
                        cells = cells.Resize(1, maxDepth - j);
                        cells.Interior.ColorIndex = cellColorIndex;
                        cells.Font.ColorIndex = fontColorIndex;

                        // チェック予定日欄を空欄にする
                        var dateCell = dstSheet.Cells[startRow + i, dateColumn];
                        dateCell.Value = null;
                    }

                    void ApplyCommentCellColor(Color cellColor, Color fontColor)
                    {
                        // ここより右のセルの色を変える
                        var cells = dstSheet.Cells[startRow + i, startColumn + j];
                        cells = cells.Resize(1, maxDepth - j);
                        cells.Interior.Color = ColorTranslator.ToOle(cellColor);
                        cells.Font.Color = ColorTranslator.ToOle(fontColor);

                        // チェック予定日欄を空欄にする
                        var dateCell = dstSheet.Cells[startRow + i, dateColumn];
                        dateCell.Value = null;
                    }

                    bool InitializeCommentCell(string text_, string pattern, int cellColorIndex, int fontColorIndex)
                    {
                        if (!Regex.IsMatch(text_, pattern))
                        {
                            return false;
                        }

                        ApplyCommentCell(cellColorIndex, fontColorIndex);

                        return true;
                    }

                    string text = node.text;

                    // NOTE: ユーザーがざっと内容を読むだけでも知っておくべき有用な情報です。
                    // TIP: 物事をより良く、または簡単に行うための役立つアドバイスです。
                    // IMPORTANT: ユーザーが目的を達成するために知っておくべき重要な情報です。
                    // WARNING: 問題を回避するために、ユーザーがすぐに注意を払う必要がある緊急の情報です。
                    // CAUTION: 特定の行動に伴うリスクや悪影響についての注意喚起です。
                    var tagColors = new Dictionary<string, (Color cellColor, Color fontColor)>
                    {
                        { "[!NOTE]", (ColorTranslator.FromHtml("#cce5ff"), ColorTranslator.FromHtml("#004085")) },
                        { "[!TIP]", (ColorTranslator.FromHtml("#d4edda"), ColorTranslator.FromHtml("#155724")) },
                        //{ "[!IMPORTANT]", (ColorTranslator.FromHtml("#d1ecf1"), ColorTranslator.FromHtml("#0c5460")) },
                        { "[!IMPORTANT]", (ColorTranslator.FromHtml("#e2dbff"), ColorTranslator.FromHtml("#5936bb")) },
                        { "[!WARNING]", (ColorTranslator.FromHtml("#fff3cd"), ColorTranslator.FromHtml("#856404")) },
                        { "[!CAUTION]", (ColorTranslator.FromHtml("#f8d7da"), ColorTranslator.FromHtml("#721c24")) }
                    };

                    bool applied = false;
                    foreach (var tag in tagColors.Keys)
                    {
                        if (text.StartsWith(tag, StringComparison.OrdinalIgnoreCase))
                        {
                            text = text.Substring(tag.Length).Trim();

                            ApplyCommentCellColor(tagColors[tag].cellColor, tagColors[tag].fontColor);
                            dstSheet.Cells[startRow + i, startColumn + j].Value = text;
                            dstSheet.Cells[startRow + i, resultColumn].Value = "-";
                            applied = true;
                            break;
                        }
                    }

                    if (applied)
                    {
                        break;
                    }

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
            var rangeforNamedRange = dstSheet.GetRange(startRow, resultColumn, leafCount, 1 + actualTimeColumnOffset);
            var namedRange = dstSheet.Names.Add(Name: ssSheetRangeName, RefersTo: rangeforNamedRange);
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
                        string path = GetAbsolutePathFromBasePath(jsonFilePath, node.imageFilePath);
                        var cell = dstSheet.Cells[startRow + i, startColumn + j];

                        if (!File.Exists(path))
                        {
                            // XXX: 毎回パス構築はムダ
                            path = GetAbsolutePathFromExecutingDirectory(noImageFilePath);
                            missingImagePaths.Add((filePath: node.imageFilePath, sheetName: dstSheet.Name, address: cell.Address));
                        }

                        AddPictureAsComment(cell, path);
                    }
                }
            }

            if (sheetValuesInfo != null)
            {
                SheetValuesInfo dstSheetValuesInfo = SheetValuesInfo.CreateFromSheet(dstSheet);
                var ignoreColumnOffsets = dstSheetValuesInfo.IgnoreColumnOffsets;
                Debug.Assert(AreHashSetsEqual(ignoreColumnOffsets, sheetValuesInfo.IgnoreColumnOffsets), "AreHashSetsEqual(ignoreColumnOffsets, sheetAddressInfo.RangeInfo.IgnoreColumnOffsets)");

                // idValues を key にした行（List<object>）の dictionary を作る
                var valuesDictionary = sheetValuesInfo.RowDictionaryWithIDKeys;

                // dstSheet の Values のコピーを作って、元のシートの Values から id を基に上書きコピーする
                // idが見つからない行、ignoreColumn は何もしないので、dstSheet のものが採用される
                var values = CopyValuesById(dstSheetValuesInfo.Values, dstSheetValuesInfo.Ids, valuesDictionary, ignoreColumnOffsets);

                dstSheetValuesInfo.Range.Value2 = values;
            }

            // シートID をカスタムプロパティに保存
            string id = sheetNode["id"].ToString();
            dstSheet.SetCustomProperty(sheetIdCustomPropertyName, id);

            return missingImagePaths;
        }

        static void AddPictureAsComment(Excel.Range cell, string imageFilePath)
        {
            using (System.Drawing.Image image = System.Drawing.Image.FromFile(imageFilePath))
            {
                float dpiX = image.HorizontalResolution;
                float dpiY = image.VerticalResolution;

                // ピクセル → ポイント（1インチ = 72ポイント）
                float widthInPoints = image.Width * 72f / dpiX;
                float heightInPoints = image.Height * 72f / dpiY;

                // コメントを追加し、画像を背景に設定
                var comment = cell.AddComment(" ");
                comment.Visible = false;
                comment.Shape.Fill.UserPicture(imageFilePath);
                comment.Shape.Width = widthInPoints;
                comment.Shape.Height = heightInPoints;
            }
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
