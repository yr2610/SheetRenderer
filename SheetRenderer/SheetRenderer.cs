using System;
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
using System.IO.Compression;

using System.Drawing;

using System.Reflection;

using System.Diagnostics;

using ExcelDna.Integration;

using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

using System.Net;
using System.Net.Http;

using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

using Microsoft.ClearScript.V8;


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

public class PullProgressForm : Form
{
    private readonly TextBox logTextBox;
    private readonly Label statusLabel;
    private readonly Button continueButton;
    private readonly TaskCompletionSource<bool> continueTcs;

    public PullProgressForm(string title = "最新版を取得", string initialStatusText = "ダウンロード中...")
    {
        this.Width = 760;
        this.Height = 420;
        this.FormBorderStyle = FormBorderStyle.FixedSingle;
        this.ControlBox = false;
        this.Text = title;
        this.StartPosition = FormStartPosition.CenterScreen;

        statusLabel = new Label();
        statusLabel.Left = 20;
        statusLabel.Top = 15;
        statusLabel.Width = 700;
        statusLabel.Height = 24;
        statusLabel.Text = initialStatusText;
        Controls.Add(statusLabel);

        logTextBox = new TextBox();
        logTextBox.Left = 20;
        logTextBox.Top = 45;
        logTextBox.Width = 700;
        logTextBox.Height = 280;
        logTextBox.Multiline = true;
        logTextBox.ReadOnly = true;
        logTextBox.ScrollBars = ScrollBars.Vertical;
        logTextBox.WordWrap = false;
        Controls.Add(logTextBox);

        continueButton = new Button();
        continueButton.Left = 560;
        continueButton.Top = 340;
        continueButton.Width = 160;
        continueButton.Height = 32;
        continueButton.Text = "処理中...";
        continueButton.Enabled = false;
        continueButton.Visible = false;
        continueButton.Click += ContinueButton_Click;
        Controls.Add(continueButton);

        continueTcs = new TaskCompletionSource<bool>();
    }

    private void ContinueButton_Click(object sender, EventArgs e)
    {
        continueButton.Enabled = false;
        continueTcs.TrySetResult(true);
    }

    public void AppendLine(string message)
    {
        if (InvokeRequired)
        {
            BeginInvoke(new Action<string>(AppendLine), message);
            return;
        }

        if (string.IsNullOrWhiteSpace(message))
        {
            return;
        }

        logTextBox.AppendText(message + Environment.NewLine);
        logTextBox.SelectionStart = logTextBox.TextLength;
        logTextBox.ScrollToCaret();
    }

    public void SetStatusText(string statusText)
    {
        if (InvokeRequired)
        {
            BeginInvoke(new Action<string>(SetStatusText), statusText);
            return;
        }

        statusLabel.Text = string.IsNullOrWhiteSpace(statusText) ? string.Empty : statusText;
    }

    public void ShowContinueButton(string buttonText, string statusText)
    {
        if (InvokeRequired)
        {
            BeginInvoke(new Action<string, string>(ShowContinueButton), buttonText, statusText);
            return;
        }

        statusLabel.Text = statusText;
        continueButton.Text = buttonText;
        continueButton.Visible = true;
        continueButton.Enabled = true;
    }

    public Task WaitForContinueAsync()
    {
        return continueTcs.Task;
    }

    public void CloseForm()
    {
        if (InvokeRequired)
        {
            Invoke(new Action(CloseForm));
            return;
        }

        Close();
    }
}

public sealed class ExcelUiSuspendScope : IDisposable
{
    private readonly Excel.Application excelApp;
    private readonly bool originalDisplayAlerts;
    private readonly bool originalScreenUpdating;
    private readonly Excel.XlCalculation originalCalculation;
    private readonly bool originalEnableEvents;

    public ExcelUiSuspendScope(Excel.Application excelApp)
    {
        this.excelApp = excelApp;
        if (this.excelApp == null)
        {
            return;
        }

        originalDisplayAlerts = this.excelApp.DisplayAlerts;
        originalScreenUpdating = this.excelApp.ScreenUpdating;
        originalCalculation = this.excelApp.Calculation;
        originalEnableEvents = this.excelApp.EnableEvents;

        this.excelApp.DisplayAlerts = false;
        this.excelApp.ScreenUpdating = false;
        this.excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
        this.excelApp.EnableEvents = false;
    }

    public void Dispose()
    {
        if (excelApp == null)
        {
            return;
        }

        excelApp.EnableEvents = originalEnableEvents;
        excelApp.Calculation = originalCalculation;
        excelApp.ScreenUpdating = originalScreenUpdating;
        excelApp.DisplayAlerts = originalDisplayAlerts;
    }
}


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
    private static PullSessionContext currentPullSession;
    private static PullSessionContext lastSuccessfulPullSession;
    [ThreadStatic]
    private static string currentGitLabBaseFileRelativePath;
    [ThreadStatic]
    private static string currentRootDirectory;

    private sealed class PullSessionContext
    {
        public string BaseUrl;
        public string ProjectId;
        public string RefName;
        public string Token;
        public string WorkRoot;
        public string EntryGitLabRelativePath;
        public string ManifestPath;
        public bool ManifestHasBeenWritten;
        public PullSessionLog SessionLog;
        public HashSet<string> ExpandedArchiveFolders;
        public Action<string> ProgressReporter;
    }

    private enum PullFileActionType
    {
        InitialFolderDownload,
        LazyFileRead,
        AlreadyExists,
        FileReadTrace
    }

    private sealed class PullFileActivity
    {
        public DateTime Timestamp;
        public PullFileActionType ActionType;
        public string RelativePath;
    }

    private sealed class PullManifest
    {
        public string BaseUrl { get; set; }

        public string ProjectId { get; set; }

        public string RefName { get; set; }

        public string EntryFilePath { get; set; }

        public string WorkRoot { get; set; }

        public string CreatedAt { get; set; }

        public List<PullManifestFileRecord> Files { get; set; }
    }

    private sealed class PullManifestFileRecord
    {
        public string GitLabRelativePath { get; set; }

        public string LocalPath { get; set; }

        public string SourceKind { get; set; }
    }

    private sealed class PullManifestReuseContext
    {
        public string SourceWorkRoot;

        public string ManifestPath;

        public PullManifest Manifest;

        public Dictionary<string, PullManifestFileRecord> FilesByGitLabRelativePath;
    }

    private sealed class PullExecutionResult
    {
        public string EntryLocalPath;

        public string JsonFilePath;

        public string NormalizedEntryPath;

        public string RefCommitId;
    }

    private sealed class SharedReceiveResult
    {
        public int AppliedCount;

        public int ConflictAppliedSheetCount;

        public int ConflictAppliedCellCount;

        public List<string> ConflictSheetNames = new List<string>();
    }

    private sealed class PullSessionLog
    {
        private readonly List<PullFileActivity> activities;

        public PullSessionLog(string workRoot, string entryGitLabRelativePath)
        {
            WorkRoot = workRoot;
            EntryGitLabRelativePath = entryGitLabRelativePath;
            activities = new List<PullFileActivity>();
        }

        public string WorkRoot { get; private set; }

        public string EntryGitLabRelativePath { get; private set; }

        public int TotalCount
        {
            get { return activities.Count; }
        }

        public void Add(PullFileActionType actionType, string relativePath)
        {
            activities.Add(new PullFileActivity
            {
                Timestamp = DateTime.Now,
                ActionType = actionType,
                RelativePath = relativePath
            });
        }

        public int CountByType(PullFileActionType actionType)
        {
            return activities.Count(a => a.ActionType == actionType);
        }

        public List<PullFileActivity> GetActivities()
        {
            return new List<PullFileActivity>(activities);
        }

        public string BuildSummaryText(int maxItems)
        {
            var lines = new List<string>();
            lines.Add("Pull completed.");
            lines.Add(string.Empty);
            lines.Add("WorkRoot:");
            lines.Add(WorkRoot ?? string.Empty);
            lines.Add(string.Empty);
            lines.Add("Entry:");
            lines.Add(EntryGitLabRelativePath ?? string.Empty);
            lines.Add(string.Empty);
            lines.Add("Event Count: " + TotalCount);
            lines.Add("- initial-folder-download: " + CountByType(PullFileActionType.InitialFolderDownload));
            lines.Add("- lazy-file-read: " + CountByType(PullFileActionType.LazyFileRead));
            lines.Add("- already-exists: " + CountByType(PullFileActionType.AlreadyExists));
            lines.Add("- file-read-trace: " + CountByType(PullFileActionType.FileReadTrace));
            lines.Add(string.Empty);
            lines.Add("File activity:");

            int shownCount = 0;
            foreach (var activity in activities)
            {
                if (shownCount >= maxItems)
                {
                    break;
                }

                lines.Add("[" + ToLabel(activity.ActionType) + "] " + activity.RelativePath);
                shownCount++;
            }

            if (TotalCount > shownCount)
            {
                lines.Add("... (" + (TotalCount - shownCount) + " more)");
            }

            return string.Join("\n", lines);
        }

        private static string ToLabel(PullFileActionType actionType)
        {
            switch (actionType)
            {
                case PullFileActionType.InitialFolderDownload:
                    return "initial-folder-download";
                case PullFileActionType.LazyFileRead:
                    return "lazy-file-read";
                case PullFileActionType.AlreadyExists:
                    return "already-exists";
                case PullFileActionType.FileReadTrace:
                    return "file-read-trace";
                default:
                    return "unknown";
            }
        }
    }

    public void OnLoad(IRibbonUI ribbonUI)
    {
        this.ribbon = ribbonUI;
    }

    public override string GetCustomUI(string RibbonID)
    {
        if (!AuthorizationHelper.IsAuthorizedUser())
        {
            return string.Empty;
        }

        string projectName = Assembly.GetExecutingAssembly().GetName().Name;
        return $@"
                <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='OnLoad'>
                <ribbon>
                    <tabs>
                    <tab id='tab1' label='{projectName}'>
                        <group id='group1' label='生成'>
                        <splitButton id='splitButton1' size='large'>
                            <button id='button2' label='更新' screentip='ファイル内の全シートを更新します' imageMso='TableDrawTable' onAction='OnRenderButtonPressed'/>
                            <menu id='menu1'>
                            <button id='button2a' label='新規作成' onAction='OnCreateNewButtonPressed'/>
                            <button id='button2b' label='再生成' onAction='OnRegenerateWorkbookPressed'/>
                            </menu>
                        </splitButton>
                        <button id='button3' label='シート更新' screentip='表示中のシートのみ更新します' size='large' imageMso='TableSharePointListsRefreshList' onAction='OnUpdateCurrentSheetButtonPressed' getEnabled='GetUpdateCurrentSheetButtonEnabled'/>
                        </group>
                        <group id='groupSync' label='同期'>
                        <splitButton id='splitButtonPull' size='large'>
                            <button id='buttonPull'
                                    label='最新版取得'
                                    screentip='保存済みの情報があれば最新版を取得して反映します'
                                    imageMso='RefreshAll'
                                    onAction='OnPullButtonPressed'/>
                            <menu id='menuPull'>
                                <button id='buttonPullCreate'
                                        label='新規作成'
                                        screentip='Pull 情報を入力して新しくブックを作成します'
                                        onAction='OnPullCreateButtonPressed'/>
                            </menu>
                        </splitButton>
                        <splitButton id='splitButtonShare' size='large'>
                            <button id='buttonShare'
                                    label='シート共有'
                                    screentip='表示中のシートの入力値を共有先へ送信します'
                                    imageMso='FileSendAsAttachment'
                                    onAction='OnShareCurrentSheetButtonPressed'/>
                            <menu id='menuShare'>
                                <button id='buttonShareDiff'
                                        label='シート差分表示'
                                        screentip='表示中のシートの共有差分を確認します'
                                        onAction='OnShowCurrentSheetDiffButtonPressed'/>
                                <button id='buttonShareRevert'
                                        label='シート変更を戻す'
                                        screentip='表示中のシートの未共有の変更を共有前の状態に戻します'
                                        onAction='OnRevertCurrentSheetChangesButtonPressed'/>
                                <button id='buttonShareAll'
                                        label='全シート共有'
                                        screentip='変更した全シートの入力値を共有先へ送信します'
                                        onAction='OnShareButtonPressed'/>
                            </menu>
                        </splitButton>
                        </group>
                        <group id='groupSettings' label='設定'>
                        <button id='buttonPullSourceSettings'
                                label='取得元設定'
                                screentip='Pull 新規作成で使う取得元 GitLab 情報を設定します'
                                size='large'
                                imageMso='CurrentViewSettings'
                                onAction='OnPullSourceSettingsButtonPressed'/>
                        <button id='buttonPullShareSettings'
                                label='共有先設定'
                                screentip='共有値同期で使う共有先 GitLab 情報を設定します'
                                size='large'
                                imageMso='CurrentViewSettings'
                                onAction='OnPullShareSettingsButtonPressed'/>
                        <button id='buttonTokenManager'
                                label='トークン管理'
                                screentip='保存済みトークンを一覧表示し、不要なものを削除します'
                                supertip='このPCに保存されたアクセストークンを管理します。&#10;&#10;期限切れ・不要になったトークンを選択して削除できます。&#10;削除後、次回同期時に再入力が求められます。'
                                size='large'
                                imageMso='AdpDiagramKeys'
                                onAction='OnTokenManagerButtonPressed'/>
                        </group>
                        <group id='groupDebug' label='DevTools'>
                        <button id='buttonDebugParse' label='Parse' screentip='Parseのみ（開発用）' imageMso='ControlToolboxOutlook' onAction='OnDebugParseButtonPressed'/>
                        <button id='buttonDebugNestedLazyRead' label='NestedRead' screentip='lazy-read の入れ子相対解決確認（開発用）' imageMso='ControlToolboxOutlook' onAction='OnDebugValidateNestedLazyReadButtonPressed'/>
                        <button id='buttonDebugRender' label='Render' screentip='Renderのみ（開発用）' imageMso='ControlToolboxOutlook' onAction='OnRenderOnlyDebugButtonPressed'/>
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
            openFileDialog.Filter = "Text ファイル (*.txt)|*.txt";
            openFileDialog.Title = "ソースファイルを選択してください";

            return (openFileDialog.ShowDialog() == DialogResult.OK)
                ? openFileDialog.FileName
                : null;
        }
    }

    static string TxtToJsonPath(string txtPath)
    {
        return Path.ChangeExtension(txtPath, ".json");
    }

    static string SelectSourceFileForRender(WorkbookInfo workbookInfo)
    {
        string projectId = workbookInfo.ProjectId;

        var lastRenderLog = workbookInfo.LastRenderLog;
        bool isSameUser = lastRenderLog.User == Environment.UserName;
        string storedSourceFilePath = NormalizeSourceFilePath(lastRenderLog.SourceFilePath);
        bool isPullWorkSourceFile = IsPullWorkSourceFilePath(storedSourceFilePath);
        bool sourceFileExists = CanReuseStoredSourceFilePath(storedSourceFilePath);
        string txtFilePath;

        if (!isSameUser || !sourceFileExists)
        {
            string message = !isSameUser
                ? "最後に更新された環境と異なります。"
                : isPullWorkSourceFile
                    ? "Pull で取得した一時ファイルはローカル更新に使用できません。"
                    : "ソースファイルが見つかりません。";

            DialogResult fileSelectionResult = MessageBox.Show(
                $"{message}\nProject ID が「{projectId}」の TXT を選択し直してください。", "確認",
                MessageBoxButtons.OKCancel);

            if (fileSelectionResult != DialogResult.OK)
            {
                return null;
            }

            txtFilePath = OpenSourceFile();
            if (txtFilePath == null)
            {
                return null;
            }
        }
        else
        {
            txtFilePath = storedSourceFilePath;
        }

        return txtFilePath;
    }

    static string NormalizeSourceFilePath(string storedPath)
    {
        if (string.IsNullOrEmpty(storedPath))
        {
            return storedPath;
        }

        if (string.Equals(Path.GetExtension(storedPath), ".json", StringComparison.OrdinalIgnoreCase))
        {
            string candidateTxtPath = Path.ChangeExtension(storedPath, ".txt");
            if (File.Exists(candidateTxtPath))
            {
                return candidateTxtPath;
            }
        }

        return storedPath;
    }

    static bool IsPathUnderDirectory(string path, string directoryPath)
    {
        if (string.IsNullOrWhiteSpace(path) || string.IsNullOrWhiteSpace(directoryPath))
        {
            return false;
        }

        try
        {
            string fullPath = Path.GetFullPath(path)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string fullDirectoryPath = Path.GetFullPath(directoryPath)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            if (string.Equals(fullPath, fullDirectoryPath, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            string directoryWithSeparator = fullDirectoryPath + Path.DirectorySeparatorChar;
            return fullPath.StartsWith(directoryWithSeparator, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    static bool IsPullWorkSourceFilePath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return false;
        }

        string normalizedPath = NormalizeSourceFilePath(path);
        return IsPathUnderDirectory(normalizedPath, GetPullWorkParentDirectory());
    }

    static bool CanReuseStoredSourceFilePath(string storedPath)
    {
        if (string.IsNullOrWhiteSpace(storedPath))
        {
            return false;
        }

        string normalizedPath = NormalizeSourceFilePath(storedPath);
        if (string.IsNullOrWhiteSpace(normalizedPath))
        {
            return false;
        }

        if (IsPullWorkSourceFilePath(normalizedPath))
        {
            return false;
        }

        return File.Exists(normalizedPath);
    }

    static string GetStoredSourceFilePathFromWorkbook(Excel.Workbook workbook)
    {
        if (workbook == null)
        {
            return null;
        }

        var renderLog = workbook.GetCustomProperty<RenderLog>("RenderLog");
        if (renderLog == null)
        {
            return null;
        }

        string storedSourceFilePath = NormalizeSourceFilePath(renderLog.SourceFilePath);
        if (!CanReuseStoredSourceFilePath(storedSourceFilePath))
        {
            return null;
        }

        return storedSourceFilePath;
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
    const string confHashCustomPropertyName = "ConfHash";
    const string indexSheetTemplateCellsCustomPropertyName = "SheetImageHash";
    const string gitLabPullInfoCustomPropertyName = "GitLabPullInfo";
    const string gitLabPullCommitIdCustomPropertyName = "GitLabPullCommitId";
    const string gitLabShareInfoCustomPropertyName = "GitLabShareInfo";
    const string sharedSheetSyncStateCustomPropertyName = "SharedSheetSyncState";
    const string sharedSheetBaseStoreSheetName = "SS_SYNC_STATE";

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

    // JsHost.Call の戻り値が「Quit を表す匿名型」かどうかを判定する
    static bool IsQuitResult(object result)
    {
        if (result == null)
        {
            return false;
        }

        var type = result.GetType();
        var quitProp = type.GetProperty("Quit", BindingFlags.Instance | BindingFlags.Public);
        if (quitProp == null || quitProp.PropertyType != typeof(bool))
        {
            return false;
        }

        var value = quitProp.GetValue(result, null);
        return value is bool b && b;
    }

    class WorkbookInfo
    {
        public string ProjectId { get; set; }
        public string IndexSheetName { get; set; }
        public string TemplateSheetName { get; set; }
        public RenderLog LastRenderLog { get; set; }
        public GitLabLastInput PullInfo { get; set; }
        public string PullCommitId { get; set; }
        public GitLabShareInfo ShareInfo { get; set; }
        public SharedSheetSyncState SharedSheetSyncState { get; set; }

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
            var pullInfo = workbook.GetCustomProperty<GitLabLastInput>(gitLabPullInfoCustomPropertyName);
            string pullCommitId = workbook.GetCustomProperty(gitLabPullCommitIdCustomPropertyName);
            var shareInfo = workbook.GetCustomProperty<GitLabShareInfo>(gitLabShareInfoCustomPropertyName);
            var sharedSheetSyncState = workbook.GetCustomProperty<SharedSheetSyncState>(sharedSheetSyncStateCustomPropertyName);
            Debug.Assert(indexSheetName != null, "indexSheetName != null");
            Debug.Assert(templateSheetName != null, "templateSheetName != null");
            Debug.Assert(lastRenderLog != null, "lastRenderLog != null");

            return new WorkbookInfo
            {
                ProjectId = projectId,
                IndexSheetName = indexSheetName,
                TemplateSheetName = templateSheetName,
                LastRenderLog = lastRenderLog,
                PullInfo = pullInfo,
                PullCommitId = pullCommitId,
                ShareInfo = shareInfo,
                SharedSheetSyncState = sharedSheetSyncState,
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
                if (sheetAddressInfo == null || sheetAddressInfo.RangeInfo == null || sheetAddressInfo.RangeInfo.IgnoreColumnOffsets == null)
                {
                    return new HashSet<int>();
                }

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

    private static object NormalizeSharedCellValue(object value)
    {
        if (value == null || value == DBNull.Value)
        {
            return null;
        }

        if (value is string || value is bool || value is double || value is float ||
            value is decimal || value is int || value is long || value is short ||
            value is byte || value is DateTime)
        {
            return value;
        }

        return value.ToString();
    }

    private static object[][] ConvertSheetValuesToJaggedArray(object[,] values)
    {
        if (values == null)
        {
            return new object[0][];
        }

        int rowCount = values.GetLength(0);
        int columnCount = values.GetLength(1);
        var result = new object[rowCount][];

        for (int row = 0; row < rowCount; row++)
        {
            result[row] = new object[columnCount];
            for (int col = 0; col < columnCount; col++)
            {
                result[row][col] = NormalizeSharedCellValue(values[row + 1, col + 1]);
            }
        }

        return result;
    }

    private static SharedRangeInfo CloneRangeInfo(RangeInfo rangeInfo)
    {
        if (rangeInfo == null)
        {
            return null;
        }

        return new SharedRangeInfo
        {
            IdColumnOffset = rangeInfo.IdColumnOffset,
            IgnoreColumnOffsets = rangeInfo.IgnoreColumnOffsets == null
                ? new HashSet<int>()
                : new HashSet<int>(rangeInfo.IgnoreColumnOffsets)
        };
    }

    private static JsonArray CreateSharedSheetValuesJsonArray(object[][] values)
    {
        var rows = new JsonArray();
        if (values == null)
        {
            return rows;
        }

        foreach (object[] row in values)
        {
            var rowArray = new JsonArray();
            if (row != null)
            {
                foreach (object value in row)
                {
                    if (value == null)
                    {
                        rowArray.Add((JsonNode)null);
                    }
                    else
                    {
                        rowArray.Add(JsonValue.Create(value));
                    }
                }
            }
            rows.Add(rowArray);
        }

        return rows;
    }

    private static JsonArray CreateSharedSheetRowIdsJsonArray(IEnumerable<object> rowIds)
    {
        var result = new JsonArray();
        if (rowIds == null)
        {
            return result;
        }

        foreach (object rowId in rowIds)
        {
            if (rowId == null)
            {
                result.Add((JsonNode)null);
            }
            else
            {
                result.Add(JsonValue.Create(NormalizeSharedCellValue(rowId)));
            }
        }

        return result;
    }

    private static JsonNode CreateSharedSheetJsonNode(SharedSheetDocument sheetDocument, bool includeHash)
    {
        if (sheetDocument == null)
        {
            return null;
        }

        var root = new JsonObject
        {
            ["project"] = sheetDocument.Project,
            ["sheetId"] = sheetDocument.SheetId,
            ["sheetName"] = sheetDocument.SheetName,
            ["rangeAddress"] = sheetDocument.RangeAddress,
            ["rowIds"] = CreateSharedSheetRowIdsJsonArray(sheetDocument.RowIds),
            ["values"] = CreateSharedSheetValuesJsonArray(sheetDocument.Values)
        };

        if (sheetDocument.RangeInfo != null)
        {
            var rangeInfoNode = new JsonObject
            {
                ["idColumnOffset"] = sheetDocument.RangeInfo.IdColumnOffset
            };

            var ignoreColumnOffsets = new JsonArray();
            if (sheetDocument.RangeInfo.IgnoreColumnOffsets != null)
            {
                foreach (int ignoreColumnOffset in sheetDocument.RangeInfo.IgnoreColumnOffsets.OrderBy(x => x))
                {
                    ignoreColumnOffsets.Add(ignoreColumnOffset);
                }
            }

            rangeInfoNode["ignoreColumnOffsets"] = ignoreColumnOffsets;
            root["rangeInfo"] = rangeInfoNode;
        }

        if (includeHash && !string.IsNullOrWhiteSpace(sheetDocument.Hash))
        {
            root["hash"] = sheetDocument.Hash;
        }

        return root;
    }

    private static string ComputeSharedSheetHash(SharedSheetDocument sheetDocument)
    {
        JsonNode jsonNode = CreateSharedSheetJsonNode(sheetDocument, includeHash: false);
        return jsonNode == null ? null : jsonNode.ComputeSha256();
    }

    private static SharedSheetDocument CreateSharedSheetDocument(Excel.Worksheet sheet)
    {
        if (sheet == null)
        {
            return null;
        }

        string sheetId = sheet.GetCustomProperty(sheetIdCustomPropertyName);
        if (string.IsNullOrWhiteSpace(sheetId))
        {
            return null;
        }

        SheetAddressInfo sheetAddressInfo = GetSheetAddressInfo(sheet);
        if (sheetAddressInfo == null)
        {
            return null;
        }

        SheetValuesInfo sheetValuesInfo = SheetValuesInfo.CreateFromSheet(sheet);
        Excel.Workbook workbook = sheet.Parent as Excel.Workbook;
        string projectId = workbook == null ? null : workbook.GetCustomProperty(ssProjectIdCustomPropertyName);

        var document = new SharedSheetDocument
        {
            Project = projectId,
            SheetId = sheetId,
            SheetName = sheet.Name,
            RangeAddress = sheetAddressInfo.Address,
            RangeInfo = CloneRangeInfo(sheetAddressInfo.RangeInfo),
            RowIds = sheetValuesInfo.Ids == null
                ? new object[0]
                : sheetValuesInfo.Ids.Select(NormalizeSharedCellValue).ToArray(),
            Values = ConvertSheetValuesToJaggedArray(sheetValuesInfo.Values)
        };
        document.Hash = ComputeSharedSheetHash(document);

        return document;
    }

    private static List<SharedSheetDocument> CollectSharedSheetDocuments(Excel.Workbook workbook)
    {
        var result = new List<SharedSheetDocument>();
        if (workbook == null)
        {
            return result;
        }

        foreach (Excel.Worksheet sheet in workbook.Sheets)
        {
            SharedSheetDocument sheetDocument = CreateSharedSheetDocument(sheet);
            if (sheetDocument != null)
            {
                result.Add(sheetDocument);
            }
        }

        return result;
    }

    private static SharedProjectManifest CreateSharedProjectManifest(Excel.Workbook workbook)
    {
        List<SharedSheetDocument> sheetDocuments = CollectSharedSheetDocuments(workbook);
        string projectId = workbook == null ? null : workbook.GetCustomProperty(ssProjectIdCustomPropertyName);

        return new SharedProjectManifest
        {
            Project = projectId,
            UpdatedAt = DateTime.UtcNow.ToString("o"),
            Sheets = sheetDocuments
                .OrderBy(x => x.SheetId, StringComparer.Ordinal)
                .Select(x => new SharedProjectManifestEntry
                {
                    SheetId = x.SheetId,
                    SheetName = x.SheetName,
                    Hash = x.Hash
                })
                .ToList()
        };
    }

    private static SharedSheetSyncState GetSharedSheetSyncState(Excel.Workbook workbook)
    {
        if (workbook == null)
        {
            return new SharedSheetSyncState
            {
                Sheets = new List<SharedSheetSyncStateEntry>()
            };
        }

        SharedSheetSyncState state = workbook.GetCustomProperty<SharedSheetSyncState>(sharedSheetSyncStateCustomPropertyName);
        if (state == null)
        {
            state = new SharedSheetSyncState();
        }

        if (state.Sheets == null)
        {
            state.Sheets = new List<SharedSheetSyncStateEntry>();
        }

        return state;
    }

    private static string GetSharedSheetBaseHash(Excel.Workbook workbook, string sheetId)
    {
        if (string.IsNullOrWhiteSpace(sheetId))
        {
            return null;
        }

        SharedSheetDocument sharedSheetDocument = GetSharedSheetBaseDocument(workbook, sheetId);
        if (sharedSheetDocument != null && !string.IsNullOrWhiteSpace(sharedSheetDocument.Hash))
        {
            return sharedSheetDocument.Hash;
        }

        SharedSheetSyncState state = GetSharedSheetSyncState(workbook);
        SharedSheetSyncStateEntry entry = state.Sheets.FirstOrDefault(x => string.Equals(x.SheetId, sheetId, StringComparison.Ordinal));
        if (entry != null && !string.IsNullOrWhiteSpace(entry.BaseHash))
        {
            return entry.BaseHash;
        }

        return null;
    }

    private static void SetSharedSheetBaseHash(Excel.Workbook workbook, string sheetId, string baseHash)
    {
        if (workbook == null || string.IsNullOrWhiteSpace(sheetId))
        {
            return;
        }

        SharedSheetSyncState state = GetSharedSheetSyncState(workbook);
        SharedSheetSyncStateEntry entry = state.Sheets.FirstOrDefault(x => string.Equals(x.SheetId, sheetId, StringComparison.Ordinal));
        if (entry == null)
        {
            entry = new SharedSheetSyncStateEntry
            {
                SheetId = sheetId
            };
            state.Sheets.Add(entry);
        }

        entry.BaseHash = baseHash;
        workbook.SetCustomProperty(sharedSheetSyncStateCustomPropertyName, state);
    }

    private static SheetViewState CaptureActiveWorkbookViewState(Excel.Workbook workbook)
    {
        if (workbook == null)
        {
            return null;
        }

        Excel.Application excelApp;
        try
        {
            excelApp = workbook.Application;
        }
        catch
        {
            return null;
        }

        Excel.Worksheet activeSheet = excelApp.ActiveSheet as Excel.Worksheet;
        if (activeSheet == null)
        {
            return null;
        }

        return new SheetViewState
        {
            SheetId = activeSheet.GetCustomProperty(sheetIdCustomPropertyName),
            SheetName = activeSheet.Name,
            ActiveCellPosition = excelApp.GetActiveCellPosition(),
            ScrollPosition = excelApp.GetScrollPosition(),
            Zoom = excelApp.GetActiveSheetZoom(),
        };
    }

    private static void RestoreActiveWorkbookViewState(Excel.Workbook workbook, SheetViewState viewState)
    {
        if (workbook == null || viewState == null)
        {
            return;
        }

        Excel.Application excelApp;
        try
        {
            excelApp = workbook.Application;
        }
        catch
        {
            return;
        }

        Excel.Worksheet targetSheet = null;

        if (!string.IsNullOrWhiteSpace(viewState.SheetId))
        {
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (string.Equals(sheet.GetCustomProperty(sheetIdCustomPropertyName), viewState.SheetId, StringComparison.Ordinal))
                {
                    targetSheet = sheet;
                    break;
                }
            }
        }

        if (targetSheet == null && !string.IsNullOrWhiteSpace(viewState.SheetName))
        {
            try
            {
                targetSheet = workbook.Worksheets[viewState.SheetName] as Excel.Worksheet;
            }
            catch
            {
            }
        }

        if (targetSheet == null)
        {
            return;
        }

        ApplyViewState(excelApp, targetSheet, viewState);
    }

    private static ExcelUiSuspendScope TryCreateExcelUiSuspendScope(Excel.Workbook workbook)
    {
        Excel.Application excelApp = null;

        try
        {
            excelApp = workbook == null ? null : workbook.Application;
        }
        catch
        {
        }

        if (excelApp == null)
        {
            try
            {
                excelApp = (Excel.Application)ExcelDnaUtil.Application;
            }
            catch
            {
                return null;
            }
        }

        try
        {
            return new ExcelUiSuspendScope(excelApp);
        }
        catch
        {
            return null;
        }
    }

    private static void EnsureSharedSheetBaseStorePrepared(Excel.Workbook workbook, Action<string> progressReporter = null)
    {
        if (workbook == null)
        {
            return;
        }

        if (GetSharedSheetBaseStoreSheet(workbook, createIfMissing: false) != null)
        {
            return;
        }

        progressReporter?.Invoke("共有状態シートを準備しています");
        ExcelUiSuspendScope uiSuspendScope = TryCreateExcelUiSuspendScope(workbook);
        if (uiSuspendScope == null)
        {
            GetSharedSheetBaseStoreSheet(workbook, createIfMissing: true);
            return;
        }

        using (uiSuspendScope)
        {
            GetSharedSheetBaseStoreSheet(workbook, createIfMissing: true);
        }
    }

    private static void TryHideSharedSheetBaseStoreSheet(Excel.Worksheet worksheet)
    {
        if (worksheet == null)
        {
            return;
        }

        try
        {
            Excel.XlSheetVisibility currentVisibility = (Excel.XlSheetVisibility)worksheet.Visible;
            if (currentVisibility != Excel.XlSheetVisibility.xlSheetVisible)
            {
                return;
            }
        }
        catch
        {
            return;
        }

        Excel.Workbook workbook = null;
        try
        {
            workbook = ExecuteExcelComWithRetry(() => worksheet.Parent as Excel.Workbook);
        }
        catch
        {
        }

        if (workbook == null)
        {
            return;
        }

        int visibleSheetCount = 0;
        int worksheetCount;
        try
        {
            worksheetCount = ExecuteExcelComWithRetry(() => workbook.Worksheets.Count);
        }
        catch
        {
            return;
        }

        for (int i = 1; i <= worksheetCount; i++)
        {
            Excel.Worksheet sheet;
            try
            {
                sheet = ExecuteExcelComWithRetry(() => workbook.Worksheets[i] as Excel.Worksheet);
            }
            catch
            {
                continue;
            }

            try
            {
                if ((Excel.XlSheetVisibility)sheet.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                {
                    visibleSheetCount++;
                }
            }
            catch
            {
            }
        }

        if (visibleSheetCount <= 1)
        {
            return;
        }

        try
        {
            worksheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
        }
        catch
        {
        }
    }

    private static Excel.Worksheet GetSharedSheetBaseStoreSheet(Excel.Workbook workbook, bool createIfMissing)
    {
        if (workbook == null)
        {
            return null;
        }

        int worksheetCount = ExecuteExcelComWithRetry(() => workbook.Worksheets.Count);
        for (int i = 1; i <= worksheetCount; i++)
        {
            Excel.Worksheet worksheet = ExecuteExcelComWithRetry(() => workbook.Worksheets[i] as Excel.Worksheet);
            if (string.Equals(worksheet.Name, sharedSheetBaseStoreSheetName, StringComparison.OrdinalIgnoreCase))
            {
                TryHideSharedSheetBaseStoreSheet(worksheet);
                EnsureSharedSheetBaseStoreHeader(worksheet);
                return worksheet;
            }
        }

        if (!createIfMissing)
        {
            return null;
        }

        SheetViewState originalViewState = CaptureActiveWorkbookViewState(workbook);
        Excel.Worksheet newWorksheet = ExecuteExcelComWithRetry(
            () => (Excel.Worksheet)workbook.Worksheets.Add(
                After: workbook.Worksheets[worksheetCount]));
        newWorksheet.Name = sharedSheetBaseStoreSheetName;
        newWorksheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
        EnsureSharedSheetBaseStoreHeader(newWorksheet);
        RestoreActiveWorkbookViewState(workbook, originalViewState);
        return newWorksheet;
    }

    private static void EnsureSharedSheetBaseStoreHeader(Excel.Worksheet worksheet)
    {
        if (worksheet == null)
        {
            return;
        }

        object[,] headerValues = (object[,])Array.CreateInstance(typeof(object), new int[] { 1, 4 }, new int[] { 1, 1 });
        headerValues[1, 1] = "SheetId";
        headerValues[1, 2] = "SheetName";
        headerValues[1, 3] = "BaseHash";
        headerValues[1, 4] = "ChunkCount";
        worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 4]].Value2 = headerValues;
    }

    private static int FindSharedSheetBaseStoreRow(Excel.Worksheet worksheet, string sheetId)
    {
        if (worksheet == null || string.IsNullOrWhiteSpace(sheetId))
        {
            return 0;
        }

        Excel.Range usedRange = worksheet.UsedRange;
        int startRow = usedRange.Row;
        int rowCount = usedRange.Rows.Count;
        int lastRow = Math.Max(1, startRow + rowCount - 1);
        Excel.Range idColumnRange = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[lastRow, 1]];
        object[,] idValues = ExcelExtensions.GetValuesAs2DArray(idColumnRange.Value2);

        for (int offset = 1; offset <= idValues.GetLength(0); offset++)
        {
            object value = idValues[offset, 1];
            string existingSheetId = value == null ? null : value.ToString();
            if (string.Equals(existingSheetId, sheetId, StringComparison.Ordinal))
            {
                return offset + 1;
            }
        }

        return 0;
    }

    private static int GetNextSharedSheetBaseStoreRow(Excel.Worksheet worksheet)
    {
        if (worksheet == null)
        {
            return 2;
        }

        Excel.Range usedRange = worksheet.UsedRange;
        int startRow = usedRange.Row;
        int rowCount = usedRange.Rows.Count;
        int lastRow = Math.Max(1, startRow + rowCount - 1);
        return Math.Max(2, lastRow + 1);
    }

    private static List<string> SplitSharedSheetBaseJson(string jsonText, int chunkLength = 30000)
    {
        var result = new List<string>();
        string normalizedJsonText = jsonText ?? string.Empty;

        if (normalizedJsonText.Length == 0)
        {
            result.Add(string.Empty);
            return result;
        }

        for (int i = 0; i < normalizedJsonText.Length; i += chunkLength)
        {
            int length = Math.Min(chunkLength, normalizedJsonText.Length - i);
            result.Add(normalizedJsonText.Substring(i, length));
        }

        return result;
    }

    private static void SaveSharedSheetBaseDocument(Excel.Workbook workbook, SharedSheetDocument sharedSheetDocument)
    {
        if (workbook == null || sharedSheetDocument == null || string.IsNullOrWhiteSpace(sharedSheetDocument.SheetId))
        {
            return;
        }

        string jsonText = CreateSharedSheetJsonText(sharedSheetDocument);
        List<string> chunks = SplitSharedSheetBaseJson(jsonText);
        Excel.Worksheet worksheet = GetSharedSheetBaseStoreSheet(workbook, createIfMissing: true);
        int row = FindSharedSheetBaseStoreRow(worksheet, sharedSheetDocument.SheetId);
        if (row == 0)
        {
            row = GetNextSharedSheetBaseStoreRow(worksheet);
        }

        int previousChunkCount = 0;
        object previousChunkCountValue = worksheet.Cells[row, 4].Value2;
        if (previousChunkCountValue != null)
        {
            int.TryParse(previousChunkCountValue.ToString(), out previousChunkCount);
        }

        int totalColumns = Math.Max(4 + Math.Max(chunks.Count, previousChunkCount), 4);
        object[,] rowValues = (object[,])Array.CreateInstance(typeof(object), new int[] { 1, totalColumns }, new int[] { 1, 1 });
        rowValues[1, 1] = sharedSheetDocument.SheetId;
        rowValues[1, 2] = sharedSheetDocument.SheetName ?? string.Empty;
        rowValues[1, 3] = sharedSheetDocument.Hash ?? string.Empty;
        rowValues[1, 4] = chunks.Count;

        for (int i = 0; i < chunks.Count; i++)
        {
            rowValues[1, 5 + i] = chunks[i];
        }

        worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, totalColumns]].Value2 = rowValues;
    }

    private static SharedSheetDocument GetSharedSheetBaseDocument(Excel.Workbook workbook, string sheetId)
    {
        if (workbook == null || string.IsNullOrWhiteSpace(sheetId))
        {
            return null;
        }

        Excel.Worksheet worksheet = GetSharedSheetBaseStoreSheet(workbook, createIfMissing: false);
        int row = FindSharedSheetBaseStoreRow(worksheet, sheetId);
        if (row == 0)
        {
            return null;
        }

        int chunkCount;
        object chunkCountValue = worksheet.Cells[row, 4].Value2;
        if (chunkCountValue == null || !int.TryParse(chunkCountValue.ToString(), out chunkCount) || chunkCount <= 0)
        {
            return null;
        }

        var builder = new StringBuilder();
        for (int i = 0; i < chunkCount; i++)
        {
            object chunkValue = worksheet.Cells[row, 5 + i].Value2;
            if (chunkValue != null)
            {
                builder.Append(chunkValue.ToString());
            }
        }

        SharedSheetDocument sharedSheetDocument = ParseSharedSheetDocument(builder.ToString());
        if (sharedSheetDocument != null)
        {
            sharedSheetDocument.SheetId = string.IsNullOrWhiteSpace(sharedSheetDocument.SheetId)
                ? sheetId
                : sharedSheetDocument.SheetId;

            object storedSheetName = worksheet.Cells[row, 2].Value2;
            if (string.IsNullOrWhiteSpace(sharedSheetDocument.SheetName) && storedSheetName != null)
            {
                sharedSheetDocument.SheetName = storedSheetName.ToString();
            }

            object storedBaseHash = worksheet.Cells[row, 3].Value2;
            if (string.IsNullOrWhiteSpace(sharedSheetDocument.Hash) && storedBaseHash != null)
            {
                sharedSheetDocument.Hash = storedBaseHash.ToString();
            }
        }

        return sharedSheetDocument;
    }

    private static void SetSharedSheetBaseHashes(
        Excel.Workbook workbook,
        IEnumerable<SharedSheetSelectionItem> items)
    {
        if (workbook == null)
        {
            return;
        }

        SharedSheetSyncState state = GetSharedSheetSyncState(workbook);
        foreach (SharedSheetSelectionItem item in items ?? Enumerable.Empty<SharedSheetSelectionItem>())
        {
            if (item == null || item.Document == null || string.IsNullOrWhiteSpace(item.SheetId))
            {
                continue;
            }

            SharedSheetSyncStateEntry entry = state.Sheets.FirstOrDefault(x => string.Equals(x.SheetId, item.SheetId, StringComparison.Ordinal));
            if (entry == null)
            {
                entry = new SharedSheetSyncStateEntry
                {
                    SheetId = item.SheetId
                };
                state.Sheets.Add(entry);
            }

            entry.BaseHash = item.Document.Hash;
        }

        state.Sheets = state.Sheets
            .Where(x => x != null && !string.IsNullOrWhiteSpace(x.SheetId))
            .OrderBy(x => x.SheetId, StringComparer.Ordinal)
            .ToList();

        workbook.SetCustomProperty(sharedSheetSyncStateCustomPropertyName, state);
    }

    private static void SetSharedSheetBaseHashes(Excel.Workbook workbook, IEnumerable<SharedSheetDocument> sheetDocuments)
    {
        if (workbook == null)
        {
            return;
        }

        SharedSheetSyncState state = GetSharedSheetSyncState(workbook);
        state.Sheets = (sheetDocuments ?? Enumerable.Empty<SharedSheetDocument>())
            .Where(x => x != null && !string.IsNullOrWhiteSpace(x.SheetId))
            .GroupBy(x => x.SheetId)
            .Select(x => new SharedSheetSyncStateEntry
            {
                SheetId = x.Key,
                BaseHash = x.Last().Hash
            })
            .OrderBy(x => x.SheetId, StringComparer.Ordinal)
            .ToList();

        workbook.SetCustomProperty(sharedSheetSyncStateCustomPropertyName, state);
    }

    private static string BuildSharedProjectManifestPath(string projectId)
    {
        return GitLabPathResolver.NormalizeGitLabRelativePath(projectId + "/_manifest.json");
    }

    private static string GetNormalizedShareRefName(GitLabShareInfo shareInfo)
    {
        if (shareInfo == null || string.IsNullOrWhiteSpace(shareInfo.RefName))
        {
            return "main";
        }

        return shareInfo.RefName.Trim();
    }

    private async Task<string> EnsureValidatedShareRefNameAsync(
        GitLabShareInfo shareInfo,
        string token,
        Excel.Workbook workbook = null)
    {
        if (shareInfo == null)
        {
            throw new ArgumentNullException(nameof(shareInfo));
        }

        string configuredRefName = GetNormalizedShareRefName(shareInfo);

        try
        {
            await GitLabClient.GetCommitIdAsync(
                shareInfo.BaseUrl,
                shareInfo.ProjectId,
                configuredRefName,
                token).ConfigureAwait(true);
            return configuredRefName;
        }
        catch (Exception originalException)
        {
            GitLabProjectInfo projectInfo = await GitLabClient.GetProjectInfoAsync(
                shareInfo.BaseUrl,
                shareInfo.ProjectId,
                token).ConfigureAwait(true);

            string defaultBranch = projectInfo == null ? null : projectInfo.DefaultBranch;
            if (string.IsNullOrWhiteSpace(defaultBranch) ||
                string.Equals(defaultBranch, configuredRefName, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "共有先 GitLab の branch が見つかりません。\n" +
                    "Project ID: " + shareInfo.ProjectId + "\n" +
                    "Ref: " + configuredRefName + "\n\n" +
                    originalException.Message);
            }

            await GitLabClient.GetCommitIdAsync(
                shareInfo.BaseUrl,
                shareInfo.ProjectId,
                defaultBranch,
                token).ConfigureAwait(true);

            shareInfo.RefName = defaultBranch;
            GitLabShareInfoStore.Save(shareInfo);

            if (workbook != null)
            {
                workbook.SetCustomProperty(gitLabShareInfoCustomPropertyName, shareInfo);
            }

            FileLogger.Info("[SharedRefFallback] " + configuredRefName + " -> " + defaultBranch);
            return defaultBranch;
        }
    }

    private async Task<bool> HasSharedUpdatesAsync(
        Excel.Workbook workbook,
        GitLabShareInfo shareInfo,
        string token)
    {
        if (workbook == null || shareInfo == null || string.IsNullOrWhiteSpace(token))
        {
            return false;
        }

        string projectId = workbook.GetCustomProperty(ssProjectIdCustomPropertyName);
        if (string.IsNullOrWhiteSpace(projectId))
        {
            return false;
        }

        await EnsureValidatedShareRefNameAsync(shareInfo, token, workbook);
        string configuredRefName = GetNormalizedShareRefName(shareInfo);

        GitLabProjectInfo projectInfo = await GitLabClient.GetProjectInfoAsync(
            shareInfo.BaseUrl,
            shareInfo.ProjectId,
            token).ConfigureAwait(true);
        string defaultBranch = projectInfo == null ? null : projectInfo.DefaultBranch;

        var candidateRefs = new List<string>();
        foreach (string candidate in new[] { configuredRefName, defaultBranch })
        {
            if (string.IsNullOrWhiteSpace(candidate))
            {
                continue;
            }

            if (candidateRefs.Any(x => string.Equals(x, candidate, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            candidateRefs.Add(candidate);
        }

        byte[] manifestBytes = null;
        string refName = configuredRefName;
        foreach (string candidateRef in candidateRefs)
        {
            manifestBytes = await TryDownloadSharedProjectManifestBytesAsync(
                shareInfo,
                projectId,
                candidateRef,
                token).ConfigureAwait(true);

            if (manifestBytes != null && manifestBytes.Length > 0)
            {
                refName = candidateRef;
                break;
            }
        }

        if (manifestBytes == null)
        {
            return false;
        }

        if (!string.Equals(refName, configuredRefName, StringComparison.OrdinalIgnoreCase))
        {
            shareInfo.RefName = refName;
            GitLabShareInfoStore.Save(shareInfo);
            workbook.SetCustomProperty(gitLabShareInfoCustomPropertyName, shareInfo);
        }

        SharedProjectManifest manifest = ParseSharedProjectManifest(Encoding.UTF8.GetString(manifestBytes));
        if (manifest == null || manifest.Sheets == null || manifest.Sheets.Count == 0)
        {
            return false;
        }

        foreach (SharedProjectManifestEntry entry in manifest.Sheets)
        {
            if (entry == null || string.IsNullOrWhiteSpace(entry.SheetId))
            {
                continue;
            }

            string baseHash = GetSharedSheetBaseHash(workbook, entry.SheetId);
            if (!string.Equals(baseHash, entry.Hash, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static string BuildSharedSheetPath(string projectId, string sheetId)
    {
        return GitLabPathResolver.NormalizeGitLabRelativePath(projectId + "/" + sheetId + ".json");
    }

    private static object ReadSharedJsonValue(JsonNode node)
    {
        if (node == null)
        {
            return null;
        }

        JsonValue jsonValue = node as JsonValue;
        if (jsonValue == null)
        {
            return node.ToJsonString();
        }

        JsonElement element;
        if (jsonValue.TryGetValue<JsonElement>(out element))
        {
            switch (element.ValueKind)
            {
                case JsonValueKind.Null:
                case JsonValueKind.Undefined:
                    return null;
                case JsonValueKind.String:
                    return element.GetString();
                case JsonValueKind.True:
                case JsonValueKind.False:
                    return element.GetBoolean();
                case JsonValueKind.Number:
                    return element.GetDouble();
            }
        }

        string rawText = jsonValue.ToJsonString();
        if (string.Equals(rawText, "null", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        double numberValue;
        if (double.TryParse(rawText, out numberValue))
        {
            return numberValue;
        }

        bool boolValue;
        if (bool.TryParse(rawText, out boolValue))
        {
            return boolValue;
        }

        return rawText.Trim('"');
    }

    private static object[][] ParseSharedSheetValues(JsonNode valuesNode)
    {
        JsonArray valuesArray = valuesNode as JsonArray;
        if (valuesArray == null)
        {
            return new object[0][];
        }

        var result = new object[valuesArray.Count][];

        for (int row = 0; row < valuesArray.Count; row++)
        {
            JsonArray rowArray = valuesArray[row] as JsonArray;
            if (rowArray == null)
            {
                result[row] = new object[0];
                continue;
            }

            result[row] = new object[rowArray.Count];
            for (int col = 0; col < rowArray.Count; col++)
            {
                result[row][col] = ReadSharedJsonValue(rowArray[col]);
            }
        }

        return result;
    }

    private static object[] ParseSharedSheetRowIds(JsonNode rowIdsNode)
    {
        JsonArray rowIdsArray = rowIdsNode as JsonArray;
        if (rowIdsArray == null)
        {
            return new object[0];
        }

        var result = new object[rowIdsArray.Count];
        for (int i = 0; i < rowIdsArray.Count; i++)
        {
            result[i] = ReadSharedJsonValue(rowIdsArray[i]);
        }

        return result;
    }

    private static SharedRangeInfo ParseSharedRangeInfo(JsonNode rangeInfoNode)
    {
        JsonObject rangeInfoObject = rangeInfoNode as JsonObject;
        if (rangeInfoObject == null)
        {
            return null;
        }

        var ignoreColumnOffsets = new HashSet<int>();
        JsonArray ignoreOffsetsNode = rangeInfoObject["ignoreColumnOffsets"] as JsonArray;
        if (ignoreOffsetsNode != null)
        {
            foreach (JsonNode ignoreNode in ignoreOffsetsNode)
            {
                if (ignoreNode == null)
                {
                    continue;
                }

                int value;
                if (int.TryParse(ignoreNode.ToJsonString(), out value))
                {
                    ignoreColumnOffsets.Add(value);
                }
            }
        }

        int? idColumnOffset = null;
        JsonNode idColumnOffsetNode = rangeInfoObject["idColumnOffset"];
        if (idColumnOffsetNode != null)
        {
            int value;
            if (int.TryParse(idColumnOffsetNode.ToJsonString(), out value))
            {
                idColumnOffset = value;
            }
        }

        return new SharedRangeInfo
        {
            IdColumnOffset = idColumnOffset,
            IgnoreColumnOffsets = ignoreColumnOffsets
        };
    }

    private static SharedSheetDocument ParseSharedSheetDocument(string jsonText)
    {
        if (string.IsNullOrWhiteSpace(jsonText))
        {
            return null;
        }

        JsonObject root = JsonNode.Parse(jsonText) as JsonObject;
        if (root == null)
        {
            return null;
        }

        return new SharedSheetDocument
        {
            Project = root["project"]?.GetValue<string>(),
            SheetId = root["sheetId"]?.GetValue<string>(),
            SheetName = root["sheetName"]?.GetValue<string>(),
            RangeAddress = root["rangeAddress"]?.GetValue<string>(),
            RangeInfo = ParseSharedRangeInfo(root["rangeInfo"]),
            RowIds = ParseSharedSheetRowIds(root["rowIds"]),
            Values = ParseSharedSheetValues(root["values"]),
            Hash = root["hash"]?.GetValue<string>()
        };
    }

    private static SharedProjectManifest ParseSharedProjectManifest(string jsonText)
    {
        if (string.IsNullOrWhiteSpace(jsonText))
        {
            return null;
        }

        JsonObject root = JsonNode.Parse(jsonText) as JsonObject;
        if (root == null)
        {
            return null;
        }

        var sheets = new List<SharedProjectManifestEntry>();
        JsonArray sheetsNode = root["sheets"] as JsonArray;
        if (sheetsNode != null)
        {
            foreach (JsonNode sheetNode in sheetsNode)
            {
                JsonObject sheetObject = sheetNode as JsonObject;
                if (sheetObject == null)
                {
                    continue;
                }

                sheets.Add(new SharedProjectManifestEntry
                {
                    SheetId = sheetObject["sheetId"]?.GetValue<string>(),
                    SheetName = sheetObject["sheetName"]?.GetValue<string>(),
                    Hash = sheetObject["hash"]?.GetValue<string>()
                });
            }
        }

        return new SharedProjectManifest
        {
            Project = root["project"]?.GetValue<string>(),
            UpdatedAt = root["updatedAt"]?.GetValue<string>(),
            Sheets = sheets
        };
    }

    private static object[,] ConvertJaggedArrayTo2DArray(object[][] values)
    {
        if (values == null || values.Length == 0)
        {
            return (object[,])Array.CreateInstance(typeof(object), new int[] { 1, 1 }, new int[] { 1, 1 });
        }

        int rowCount = values.Length;
        int columnCount = values.Max(row => row == null ? 0 : row.Length);
        columnCount = Math.Max(columnCount, 1);

        var result = (object[,])Array.CreateInstance(typeof(object), new int[] { rowCount, columnCount }, new int[] { 1, 1 });

        for (int row = 0; row < rowCount; row++)
        {
            object[] sourceRow = values[row];
            if (sourceRow == null)
            {
                continue;
            }

            for (int col = 0; col < sourceRow.Length; col++)
            {
                result[row + 1, col + 1] = sourceRow[col];
            }
        }

        return result;
    }

    private static IEnumerable<object> GetIdsFromSharedSheetDocument(SharedSheetDocument sheetDocument)
    {
        if (sheetDocument == null ||
            sheetDocument.RowIds == null ||
            sheetDocument.RowIds.Length == 0)
        {
            return Enumerable.Empty<object>();
        }

        return sheetDocument.RowIds;
    }

    private static bool HasAnyNonEmptySharedIds(IEnumerable<object> ids)
    {
        if (ids == null)
        {
            return false;
        }

        foreach (object id in ids)
        {
            if (id != null && !string.IsNullOrWhiteSpace(id.ToString()))
            {
                return true;
            }
        }

        return false;
    }

    private static bool AreSharedCellValuesEqual(object left, object right)
    {
        if (left == null && right == null)
        {
            return true;
        }

        if (left == null || right == null)
        {
            return false;
        }

        double leftNumber;
        double rightNumber;
        if (TryConvertToDouble(left, out leftNumber) && TryConvertToDouble(right, out rightNumber))
        {
            return leftNumber.Equals(rightNumber);
        }

        return string.Equals(left.ToString(), right.ToString(), StringComparison.Ordinal);
    }

    private static bool TryConvertToDouble(object value, out double result)
    {
        if (value == null)
        {
            result = 0;
            return false;
        }

        if (value is double doubleValue)
        {
            result = doubleValue;
            return true;
        }

        if (value is float floatValue)
        {
            result = floatValue;
            return true;
        }

        if (value is decimal decimalValue)
        {
            result = (double)decimalValue;
            return true;
        }

        if (value is int intValue)
        {
            result = intValue;
            return true;
        }

        if (value is long longValue)
        {
            result = longValue;
            return true;
        }

        if (value is short shortValue)
        {
            result = shortValue;
            return true;
        }

        if (value is byte byteValue)
        {
            result = byteValue;
            return true;
        }

        return double.TryParse(value.ToString(), out result);
    }

    private static string FormatSharedCellValueForDiff(object value)
    {
        if (value == null || value == DBNull.Value)
        {
            return "(null)";
        }

        if (value is DateTime dateTimeValue)
        {
            return dateTimeValue.ToString("yyyy-MM-dd HH:mm:ss");
        }

        return value.ToString();
    }

    private static int? TryGetSharedSheetStartRow(string rangeAddress)
    {
        if (string.IsNullOrWhiteSpace(rangeAddress))
        {
            return null;
        }

        Match match = Regex.Match(rangeAddress, @"\$?[A-Za-z]+\$?(\d+)");
        if (!match.Success)
        {
            return null;
        }

        int startRow;
        if (!int.TryParse(match.Groups[1].Value, out startRow))
        {
            return null;
        }

        return startRow;
    }

    private static int? TryGetSharedSheetStartColumn(string rangeAddress)
    {
        if (string.IsNullOrWhiteSpace(rangeAddress))
        {
            return null;
        }

        Match match = Regex.Match(rangeAddress, @"\$?([A-Za-z]+)\$?\d+");
        if (!match.Success)
        {
            return null;
        }

        string columnText = match.Groups[1].Value;
        int column = 0;
        foreach (char ch in columnText.ToUpperInvariant())
        {
            if (ch < 'A' || ch > 'Z')
            {
                return null;
            }

            column = (column * 26) + (ch - 'A' + 1);
        }

        return column > 0 ? (int?)column : null;
    }

    private static string GetExcelColumnLetter(int column)
    {
        if (column <= 0)
        {
            return "?";
        }

        string result = "";
        while (column > 0)
        {
            int mod = (column - 1) % 26;
            result = (char)('A' + mod) + result;
            column = (column - mod - 1) / 26;
        }

        return result;
    }

    private static Dictionary<string, int> CreateSharedSheetDisplayRowMap(SharedSheetDocument localDocument)
    {
        var result = new Dictionary<string, int>(StringComparer.Ordinal);
        if (!CanMergeSharedSheetByRowIds(localDocument))
        {
            return result;
        }

        int? startRow = TryGetSharedSheetStartRow(localDocument.RangeAddress);
        if (!startRow.HasValue)
        {
            return result;
        }

        for (int i = 0; i < localDocument.RowIds.Length; i++)
        {
            string rowId = NormalizeSharedRowId(localDocument.RowIds[i]);
            if (string.IsNullOrWhiteSpace(rowId) || result.ContainsKey(rowId))
            {
                continue;
            }

            result[rowId] = startRow.Value + i;
        }

        return result;
    }

    private static string BuildSharedDiffStateLabel(object baseValue, object localValue, object remoteValue)
    {
        if (AreSharedCellValuesEqual(localValue, baseValue) &&
            AreSharedCellValuesEqual(remoteValue, baseValue))
        {
            return "変更なし";
        }

        if (AreSharedCellValuesEqual(localValue, baseValue))
        {
            return "共有先変更";
        }

        if (AreSharedCellValuesEqual(remoteValue, baseValue))
        {
            return "ローカル変更";
        }

        if (AreSharedCellValuesEqual(localValue, remoteValue))
        {
            return "同一変更";
        }

        return "競合";
    }

    private static string BuildSharedSheetDiffText(
        SharedSheetDocument baseDocument,
        SharedSheetDocument localDocument,
        SharedSheetDocument remoteDocument)
    {
        if (localDocument == null)
        {
            return "差分はありません。";
        }

        if (!CanMergeSharedSheetByRowIds(localDocument))
        {
            return "rowIds が無いため差分を表示できません。";
        }

        int columnCount = Math.Max(
            GetSharedSheetColumnCount(localDocument),
            Math.Max(GetSharedSheetColumnCount(baseDocument), GetSharedSheetColumnCount(remoteDocument)));

        var ignoreColumnOffsets = localDocument.RangeInfo == null || localDocument.RangeInfo.IgnoreColumnOffsets == null
            ? new HashSet<int>()
            : new HashSet<int>(localDocument.RangeInfo.IgnoreColumnOffsets);

        Dictionary<string, object[]> localRows = CreateSharedSheetRowMap(localDocument);
        Dictionary<string, object[]> remoteRows = CreateSharedSheetRowMap(remoteDocument);
        Dictionary<string, object[]> baseRows = CreateSharedSheetRowMap(baseDocument);
        List<string> rowOrder = BuildSharedSheetRowOrder(localDocument, remoteDocument, baseDocument);
        Dictionary<string, int> displayRows = CreateSharedSheetDisplayRowMap(localDocument);
        int? startColumn = TryGetSharedSheetStartColumn(localDocument.RangeAddress);

        var lines = new List<string>();
        lines.Add("sheetName: " + (localDocument.SheetName ?? ""));
        lines.Add("sheetId: " + (localDocument.SheetId ?? ""));
        lines.Add("range: " + (localDocument.RangeAddress ?? ""));
        lines.Add("");

        int diffCount = 0;
        foreach (string rowId in rowOrder)
        {
            object[] localRow;
            localRows.TryGetValue(rowId, out localRow);
            if (localRow == null)
            {
                continue;
            }

            object[] remoteRow;
            remoteRows.TryGetValue(rowId, out remoteRow);

            object[] baseRow;
            baseRows.TryGetValue(rowId, out baseRow);

            for (int col = 0; col < columnCount; col++)
            {
                if (ignoreColumnOffsets.Contains(col))
                {
                    continue;
                }

                object baseValue = GetSharedSheetCellValue(baseRow, col);
                object localValue = GetSharedSheetCellValue(localRow, col);
                object remoteValue = remoteRow == null ? baseValue : GetSharedSheetCellValue(remoteRow, col);

                if (AreSharedCellValuesEqual(baseValue, localValue) &&
                    AreSharedCellValuesEqual(localValue, remoteValue))
                {
                    continue;
                }

                string stateLabel = BuildSharedDiffStateLabel(baseValue, localValue, remoteValue);
                int displayRow;
                bool hasDisplayRow = displayRows.TryGetValue(rowId, out displayRow);
                string displayAddress = hasDisplayRow && startColumn.HasValue
                    ? GetExcelColumnLetter(startColumn.Value + col) + displayRow
                    : "?";
                lines.Add(
                    "addr=" + displayAddress +
                    "\tstate=" + stateLabel +
                    "\tbase=" + FormatSharedCellValueForDiff(baseValue) +
                    "\tlocal=" + FormatSharedCellValueForDiff(localValue) +
                    "\tremote=" + FormatSharedCellValueForDiff(remoteValue));
                diffCount++;
            }
        }

        if (diffCount == 0)
        {
            lines.Add("差分はありません。");
        }
        else
        {
            lines.Insert(4, "changedCells: " + diffCount);
        }

        return string.Join(Environment.NewLine, lines);
    }

    private static void LogSharedOverwriteDifferences(Excel.Worksheet sheet, SheetValuesInfo currentSheetValuesInfo, object[,] mergedValues)
    {
        if (sheet == null || currentSheetValuesInfo == null || mergedValues == null)
        {
            return;
        }

        var ids = currentSheetValuesInfo.Ids == null
            ? null
            : currentSheetValuesInfo.Ids.Select(x => x == null ? null : x.ToString()).ToArray();
        int rangeStartRow = currentSheetValuesInfo.Range?.Row ?? 1;

        int rowCount = currentSheetValuesInfo.Values.GetLength(0);
        int columnCount = currentSheetValuesInfo.Values.GetLength(1);
        for (int row = 1; row <= rowCount; row++)
        {
            string rowId = ids != null && row - 1 < ids.Length ? ids[row - 1] : null;
            int displayRow = rangeStartRow + row - 1;

            for (int col = 1; col <= columnCount; col++)
            {
                if (currentSheetValuesInfo.IgnoreColumnOffsets.Contains(col - 1))
                {
                    continue;
                }

                object currentValue = currentSheetValuesInfo.Values[row, col];
                object mergedValue = mergedValues[row, col];

                if (AreSharedCellValuesEqual(currentValue, mergedValue))
                {
                    continue;
                }

                Excel.Range cell = currentSheetValuesInfo.Range.Cells[row, col];
                FileLogger.Info(
                    "[SharedReceiveOverwrite] " +
                    "sheetId=" + sheet.GetCustomProperty(sheetIdCustomPropertyName) +
                    " sheetName=" + sheet.Name +
                    " rowId=" + (rowId ?? "") +
                    " row=" + displayRow +
                    " cell=" + cell.Address[false, false] +
                    " local=" + (currentValue ?? "") +
                    " remote=" + (mergedValue ?? ""));
            }
        }
    }

    private static object[,] BuildSharedSheetRevertValues(SheetValuesInfo currentSheetValuesInfo, SharedSheetDocument baseDocument)
    {
        if (currentSheetValuesInfo == null)
        {
            throw new ArgumentNullException(nameof(currentSheetValuesInfo));
        }

        object[,] result = (object[,])currentSheetValuesInfo.Values.Clone();

        if (currentSheetValuesInfo.Ids == null || !HasAnyNonEmptySharedIds(currentSheetValuesInfo.Ids))
        {
            throw new InvalidOperationException("Shared sheet rowIds are required for revert.");
        }

        Dictionary<string, object[]> baseRows = CreateSharedSheetRowMap(baseDocument);
        int rowCount = currentSheetValuesInfo.Values.GetLength(0);
        int columnCount = currentSheetValuesInfo.Values.GetLength(1);

        var ids = currentSheetValuesInfo.Ids.ToArray();
        for (int row = 1; row <= rowCount; row++)
        {
            string rowId = row - 1 < ids.Length ? NormalizeSharedRowId(ids[row - 1]) : null;
            object[] baseRow = null;
            bool hasBaseRow = !string.IsNullOrWhiteSpace(rowId) && baseRows.TryGetValue(rowId, out baseRow);

            for (int col = 1; col <= columnCount; col++)
            {
                if (currentSheetValuesInfo.IgnoreColumnOffsets.Contains(col - 1))
                {
                    continue;
                }

                result[row, col] = hasBaseRow
                    ? GetSharedSheetCellValue(baseRow, col - 1)
                    : null;
            }
        }

        return result;
    }

    private static int LogSharedRevertDifferences(Excel.Worksheet sheet, SheetValuesInfo currentSheetValuesInfo, object[,] revertedValues)
    {
        if (sheet == null || currentSheetValuesInfo == null || revertedValues == null)
        {
            return 0;
        }

        var ids = currentSheetValuesInfo.Ids == null
            ? null
            : currentSheetValuesInfo.Ids.Select(x => x == null ? null : x.ToString()).ToArray();
        int rangeStartRow = currentSheetValuesInfo.Range?.Row ?? 1;
        int rowCount = currentSheetValuesInfo.Values.GetLength(0);
        int columnCount = currentSheetValuesInfo.Values.GetLength(1);
        int diffCount = 0;

        for (int row = 1; row <= rowCount; row++)
        {
            string rowId = ids != null && row - 1 < ids.Length ? ids[row - 1] : null;
            int displayRow = rangeStartRow + row - 1;

            for (int col = 1; col <= columnCount; col++)
            {
                if (currentSheetValuesInfo.IgnoreColumnOffsets.Contains(col - 1))
                {
                    continue;
                }

                object currentValue = currentSheetValuesInfo.Values[row, col];
                object revertedValue = revertedValues[row, col];
                if (AreSharedCellValuesEqual(currentValue, revertedValue))
                {
                    continue;
                }

                Excel.Range cell = currentSheetValuesInfo.Range.Cells[row, col];
                FileLogger.Info(
                    "[SharedRevert] " +
                    "sheetId=" + sheet.GetCustomProperty(sheetIdCustomPropertyName) +
                    " sheetName=" + sheet.Name +
                    " rowId=" + (rowId ?? "") +
                    " row=" + displayRow +
                    " cell=" + cell.Address[false, false] +
                    " local=" + (currentValue ?? "") +
                    " base=" + (revertedValue ?? ""));
                diffCount++;
            }
        }

        return diffCount;
    }

    private static void ApplySharedSheetDocumentToWorksheet(Excel.Worksheet sheet, SharedSheetDocument sharedSheetDocument)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }

        if (sharedSheetDocument == null)
        {
            throw new ArgumentNullException(nameof(sharedSheetDocument));
        }

        SheetValuesInfo currentSheetValuesInfo = SheetValuesInfo.CreateFromSheet(sheet);
        if (currentSheetValuesInfo == null)
        {
            throw new InvalidOperationException("SS_SHEET is not defined. sheet='" + sheet.Name + "'.");
        }

        object[,] mergedValues;
        object[,] remoteValues = ConvertJaggedArrayTo2DArray(sharedSheetDocument.Values);
        object[] remoteIds = GetIdsFromSharedSheetDocument(sharedSheetDocument).ToArray();
        bool canMergeById =
            remoteIds.Length == remoteValues.GetLength(0) &&
            currentSheetValuesInfo.Ids != null &&
            HasAnyNonEmptySharedIds(currentSheetValuesInfo.Ids) &&
            HasAnyNonEmptySharedIds(remoteIds);

        if (!canMergeById)
        {
            throw new InvalidOperationException(
                "Shared sheet rowIds are required. " +
                "sheetId='" + sharedSheetDocument.SheetId + "', " +
                "sheetName='" + sheet.Name + "'.");
        }

        Dictionary<string, List<object>> remoteDictionary = CreateRowDictionaryWithIDKeys(remoteValues, remoteIds);
        mergedValues = CopyValuesById(
            currentSheetValuesInfo.Values,
            currentSheetValuesInfo.Ids,
            remoteDictionary,
            currentSheetValuesInfo.IgnoreColumnOffsets);

        LogSharedOverwriteDifferences(sheet, currentSheetValuesInfo, mergedValues);
        currentSheetValuesInfo.Range.Value2 = mergedValues;
    }

    private async Task<byte[]> TryDownloadSharedProjectManifestBytesAsync(
        GitLabShareInfo shareInfo,
        string projectId,
        string refName,
        string token)
    {
        string manifestPath = BuildSharedProjectManifestPath(projectId);
        byte[] manifestBytes = await GitLabClient.TryDownloadFileRawByPathAsync(
            shareInfo.BaseUrl,
            shareInfo.ProjectId,
            manifestPath,
            refName,
            token).ConfigureAwait(false);

        if (manifestBytes == null)
        {
            manifestBytes = await GitLabClient.TryDownloadFileViaTreeAsync(
                shareInfo.BaseUrl,
                shareInfo.ProjectId,
                projectId,
                "_manifest.json",
                refName,
                token).ConfigureAwait(false);
        }

        if (manifestBytes == null || manifestBytes.Length == 0)
        {
            return null;
        }

        return manifestBytes;
    }

    private async Task<SharedProjectManifest> TryDownloadSharedProjectManifestAsync(
        GitLabShareInfo shareInfo,
        string projectId,
        string token)
    {
        string refName = GetNormalizedShareRefName(shareInfo);
        byte[] manifestBytes = await TryDownloadSharedProjectManifestBytesAsync(
            shareInfo,
            projectId,
            refName,
            token).ConfigureAwait(false);

        if (manifestBytes == null)
        {
            return null;
        }

        return ParseSharedProjectManifest(Encoding.UTF8.GetString(manifestBytes));
    }

    private async Task<SharedSheetDocument> TryDownloadSharedSheetDocumentAsync(
        GitLabShareInfo shareInfo,
        string projectId,
        string sheetId,
        string token)
    {
        string refName = GetNormalizedShareRefName(shareInfo);
        string sheetPath = BuildSharedSheetPath(projectId, sheetId);
        byte[] sheetBytes = await GitLabClient.TryDownloadFileRawByPathAsync(
            shareInfo.BaseUrl,
            shareInfo.ProjectId,
            sheetPath,
            refName,
            token).ConfigureAwait(false);

        if (sheetBytes == null)
        {
            sheetBytes = await GitLabClient.TryDownloadFileViaTreeAsync(
                shareInfo.BaseUrl,
                shareInfo.ProjectId,
                projectId,
                sheetId + ".json",
                refName,
                token).ConfigureAwait(false);
        }

        if (sheetBytes == null || sheetBytes.Length == 0)
        {
            return null;
        }

        return ParseSharedSheetDocument(Encoding.UTF8.GetString(sheetBytes));
    }

    private static void ShowSharedReceiveConflictDialogIfNeeded(SharedReceiveResult result)
    {
        if (result == null || result.ConflictAppliedSheetCount <= 0)
        {
            return;
        }

        string detail = string.Empty;
        if (result.ConflictSheetNames != null && result.ConflictSheetNames.Count > 0)
        {
            int maxNames = Math.Min(5, result.ConflictSheetNames.Count);
            detail = Environment.NewLine +
                     "対象シート: " +
                     string.Join(", ", result.ConflictSheetNames.Take(maxNames));

            if (result.ConflictSheetNames.Count > maxNames)
            {
                detail += " ほか " + (result.ConflictSheetNames.Count - maxNames) + " シート";
            }
        }

        DialogResult dialogResult = MessageBox.Show(
            "共有値の競合があり、共有先の値で上書きされました。" + Environment.NewLine +
            "競合シート数: " + result.ConflictAppliedSheetCount + Environment.NewLine +
            "競合セル数: " + result.ConflictAppliedCellCount +
            detail + Environment.NewLine + Environment.NewLine +
            "保存せずにブックを閉じると、取得前の状態に戻せます。" + Environment.NewLine + Environment.NewLine +
            "ログファイルを開きますか？",
            "最新版取得",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);

        if (dialogResult == DialogResult.Yes)
        {
            try
            {
                FileLogger.OpenLog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "ログファイルを開けませんでした");
            }
        }
    }

    private async Task<SharedReceiveResult> ReceiveSharedSheetsAsync(
        Excel.Workbook workbook,
        GitLabShareInfo shareInfo,
        Action<string> progressReporter = null)
    {
        SharedReceiveResult result = new SharedReceiveResult();

        if (workbook == null || shareInfo == null)
        {
            progressReporter?.Invoke("共有設定が無いため共有値確認をスキップしました");
            return result;
        }

        string projectId = workbook.GetCustomProperty(ssProjectIdCustomPropertyName);
        if (string.IsNullOrWhiteSpace(projectId))
        {
            progressReporter?.Invoke("SSProjectId が無いため共有値確認をスキップしました");
            return result;
        }

        InitializeLoggerForSharedReceive(workbook);

        string token = GitLabAuth.GetOrPromptToken(shareInfo.BaseUrl, shareInfo.ProjectId);
        if (string.IsNullOrWhiteSpace(token))
        {
            MessageBox.Show("共有値の取得をキャンセルしました（トークン未入力）", "最新版取得");
            return result;
        }

        await EnsureValidatedShareRefNameAsync(shareInfo, token, workbook);
        string configuredRefName = GetNormalizedShareRefName(shareInfo);
        string manifestPath = BuildSharedProjectManifestPath(projectId);

        progressReporter?.Invoke("共有値を確認しています");
        progressReporter?.Invoke("共有先 Project ID: " + (shareInfo.ProjectId ?? ""));
        progressReporter?.Invoke("共有先 Ref: " + configuredRefName);
        progressReporter?.Invoke("共有マニフェスト Path: " + manifestPath);
        GitLabProjectInfo projectInfo = await GitLabClient.GetProjectInfoAsync(
            shareInfo.BaseUrl,
            shareInfo.ProjectId,
            token).ConfigureAwait(false);
        string defaultBranch = projectInfo == null ? null : projectInfo.DefaultBranch;

        var candidateRefs = new List<string>();
        foreach (string candidate in new[] { configuredRefName, defaultBranch })
        {
            if (string.IsNullOrWhiteSpace(candidate))
            {
                continue;
            }

            if (candidateRefs.Any(x => string.Equals(x, candidate, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            candidateRefs.Add(candidate);
        }

        byte[] manifestBytes = null;
        string refName = configuredRefName;
        foreach (string candidateRef in candidateRefs)
        {
            progressReporter?.Invoke("共有マニフェストを確認します: " + candidateRef);
            manifestBytes = await TryDownloadSharedProjectManifestBytesAsync(
                shareInfo,
                projectId,
                candidateRef,
                token).ConfigureAwait(false);

            if (manifestBytes != null && manifestBytes.Length > 0)
            {
                refName = candidateRef;
                break;
            }
        }

        if (manifestBytes == null)
        {
            FileLogger.Info(
                "[SharedReceive] shared manifest not found. project=" + projectId +
                " path=" + manifestPath +
                " refs=" + string.Join(",", candidateRefs ?? new List<string>()) +
                " shareRepo=" + shareInfo.ProjectId);
            progressReporter?.Invoke("共有値はまだありません");
            return result;
        }

        if (!string.Equals(refName, configuredRefName, StringComparison.OrdinalIgnoreCase))
        {
            shareInfo.RefName = refName;
            GitLabShareInfoStore.Save(shareInfo);
            workbook.SetCustomProperty(gitLabShareInfoCustomPropertyName, shareInfo);
            progressReporter?.Invoke("共有 branch を切り替えました: " + refName);
            FileLogger.Info("[SharedManifestRefFallback] manifest found on branch: " + refName);
        }

        progressReporter?.Invoke("共有マニフェスト取得サイズ: " + manifestBytes.Length + " bytes");
        SharedProjectManifest manifest = ParseSharedProjectManifest(Encoding.UTF8.GetString(manifestBytes));

        if (manifest == null || manifest.Sheets == null || manifest.Sheets.Count == 0)
        {
            FileLogger.Info("[SharedReceive] shared manifest not found or empty. project=" + projectId + " path=" + manifestPath + " ref=" + refName + " shareRepo=" + shareInfo.ProjectId);
            progressReporter?.Invoke("共有マニフェストが見つからないか空です");
            return result;
        }

        EnsureSharedSheetBaseStorePrepared(workbook, progressReporter);
        progressReporter?.Invoke("共有マニフェストを読み込みました: " + manifest.Sheets.Count + " シート");

        var sheetsById = new Dictionary<string, Excel.Worksheet>(StringComparer.Ordinal);
        var sheetsByName = new Dictionary<string, List<Excel.Worksheet>>(StringComparer.Ordinal);
        foreach (Excel.Worksheet sheet in workbook.Sheets)
        {
            string sheetId = sheet.GetCustomProperty(sheetIdCustomPropertyName);
            if (!string.IsNullOrWhiteSpace(sheetId))
            {
                sheetsById[sheetId] = sheet;
            }

             List<Excel.Worksheet> sameNameSheets;
             if (!sheetsByName.TryGetValue(sheet.Name, out sameNameSheets))
             {
                 sameNameSheets = new List<Excel.Worksheet>();
                 sheetsByName[sheet.Name] = sameNameSheets;
             }

             sameNameSheets.Add(sheet);
        }

        progressReporter?.Invoke("ローカルシートを確認しました: " + sheetsById.Count + " シート");

        int appliedCount = 0;
        int skippedLatestCount = 0;
        int missingLocalSheetCount = 0;
        int missingSharedDocumentCount = 0;
        int fallbackMatchedCount = 0;
        int conflictAppliedCount = 0;

        foreach (SharedProjectManifestEntry entry in manifest.Sheets)
        {
            if (entry == null || string.IsNullOrWhiteSpace(entry.SheetId))
            {
                continue;
            }

            Excel.Worksheet sheet;
            if (!sheetsById.TryGetValue(entry.SheetId, out sheet))
            {
                List<Excel.Worksheet> sameNameSheets;
                if (!string.IsNullOrWhiteSpace(entry.SheetName) &&
                    sheetsByName.TryGetValue(entry.SheetName, out sameNameSheets) &&
                    sameNameSheets.Count == 1)
                {
                    sheet = sameNameSheets[0];
                    string currentSheetId = sheet.GetCustomProperty(sheetIdCustomPropertyName);
                    if (string.IsNullOrWhiteSpace(currentSheetId))
                    {
                        sheet.SetCustomProperty(sheetIdCustomPropertyName, entry.SheetId);
                    }
                    sheetsById[entry.SheetId] = sheet;
                    fallbackMatchedCount++;
                    progressReporter?.Invoke("シート名で対応付けました: " + sheet.Name + " -> " + entry.SheetId);
                }
                else
                {
                    missingLocalSheetCount++;
                    progressReporter?.Invoke("ローカルに対象シートがありません: " + entry.SheetId);
                    continue;
                }
            }

            string baseHash = GetSharedSheetBaseHash(workbook, entry.SheetId);
            if (!string.IsNullOrWhiteSpace(baseHash) &&
                !string.IsNullOrWhiteSpace(entry.Hash) &&
                string.Equals(baseHash, entry.Hash, StringComparison.OrdinalIgnoreCase))
            {
                skippedLatestCount++;
                progressReporter?.Invoke("共有値は最新です: " + sheet.Name);
                continue;
            }

            progressReporter?.Invoke("共有値を取得しています: " + sheet.Name);
            SharedSheetDocument sharedSheetDocument = await TryDownloadSharedSheetDocumentAsync(
                shareInfo,
                projectId,
                entry.SheetId,
                token);

            if (sharedSheetDocument == null)
            {
                missingSharedDocumentCount++;
                progressReporter?.Invoke("共有JSONが見つかりません: " + entry.SheetId);
                continue;
            }

            if (string.IsNullOrWhiteSpace(sharedSheetDocument.Hash))
            {
                sharedSheetDocument.Hash = ComputeSharedSheetHash(sharedSheetDocument);
            }

            if (!string.IsNullOrWhiteSpace(baseHash) &&
                string.Equals(baseHash, sharedSheetDocument.Hash, StringComparison.OrdinalIgnoreCase))
            {
                skippedLatestCount++;
                progressReporter?.Invoke("共有値は最新です: " + sheet.Name);
                continue;
            }

            SharedSheetDocument localSheetDocument = CreateSharedSheetDocument(sheet);
            SharedSheetDocument baseSheetDocument = GetSharedSheetBaseDocument(workbook, entry.SheetId);
            SharedSheetMergeResult mergeResult = TryMergeSharedSheetDocumentForReceive(
                baseSheetDocument,
                localSheetDocument,
                sharedSheetDocument);

            if (mergeResult == null || mergeResult.MergedDocument == null)
            {
                throw new InvalidOperationException(
                    "Shared receive merge failed. " +
                    "sheetId='" + entry.SheetId + "', " +
                    "sheetName='" + sheet.Name + "'.");
            }

            ExcelUiSuspendScope uiSuspendScope = TryCreateExcelUiSuspendScope(workbook);
            if (uiSuspendScope == null)
            {
                ApplySharedSheetDocumentToWorksheet(sheet, mergeResult.MergedDocument);
                SaveSharedSheetBaseDocument(workbook, sharedSheetDocument);
            }
            else
            {
                using (uiSuspendScope)
                {
                    ApplySharedSheetDocumentToWorksheet(sheet, mergeResult.MergedDocument);
                    SaveSharedSheetBaseDocument(workbook, sharedSheetDocument);
                }
            }
            FileLogger.Info("[SharedReceive] applied sheetId=" + entry.SheetId + " hash=" + sharedSheetDocument.Hash);
            if (mergeResult.ConflictCount > 0)
            {
                conflictAppliedCount++;
                result.ConflictAppliedCellCount += mergeResult.ConflictCount;
                result.ConflictSheetNames.Add(sheet.Name);
                progressReporter?.Invoke("共有値の競合を反映しました: " + sheet.Name + " (" + mergeResult.ConflictCount + "セル)");
            }
            progressReporter?.Invoke("共有値を反映しました: " + sheet.Name);
            appliedCount++;
        }

        if (appliedCount == 0)
        {
            progressReporter?.Invoke(
                "共有値の反映対象はありませんでした"
                + " (manifest=" + manifest.Sheets.Count
                + ", upToDate=" + skippedLatestCount
                + ", localMissing=" + missingLocalSheetCount
                + ", jsonMissing=" + missingSharedDocumentCount
                + ", fallbackMatched=" + fallbackMatchedCount
                + ")");
        }
        else if (conflictAppliedCount > 0)
        {
            progressReporter?.Invoke("共有値の競合上書きがありました: " + conflictAppliedCount + " シート");
        }

        result.AppliedCount = appliedCount;
        result.ConflictAppliedSheetCount = conflictAppliedCount;
        return result;
    }

    private static void InitializeLoggerForWorkbookSession(Excel.Workbook workbook, string baseName)
    {
        string fallbackDirectoryPath = Path.Combine(Path.GetTempPath(), "SheetRenderer", "SharedReceive");
        try
        {
            string workbookPath = workbook == null ? null : workbook.FullName;
            string directoryPath = null;

            if (!string.IsNullOrWhiteSpace(workbookPath))
            {
                directoryPath = Path.GetDirectoryName(workbookPath);
            }

            if (string.IsNullOrWhiteSpace(directoryPath))
            {
                directoryPath = fallbackDirectoryPath;
            }

            FileLogger.InitializeForSession(directoryPath, baseName, timestamped: false);
        }
        catch
        {
            try
            {
                FileLogger.InitializeForSession(fallbackDirectoryPath, baseName, timestamped: false);
            }
            catch
            {
            }
        }
    }

    private static void InitializeLoggerForSharedReceive(Excel.Workbook workbook)
    {
        InitializeLoggerForWorkbookSession(workbook, "shared-receive");
    }

    private static string CreateSharedSheetJsonText(SharedSheetDocument sharedSheetDocument)
    {
        JsonNode jsonNode = CreateSharedSheetJsonNode(sharedSheetDocument, includeHash: true);
        return jsonNode == null
            ? "{}"
            : jsonNode.ToJsonString();
    }

    private static string CreateSharedSheetUploadJsonText(SharedSheetDocument sharedSheetDocument)
    {
        return JsonSerializer.Serialize(
            sharedSheetDocument ?? new SharedSheetDocument(),
            new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            });
    }

    private static string CreateSharedProjectManifestJsonText(SharedProjectManifest manifest)
    {
        return JsonSerializer.Serialize(
            manifest ?? new SharedProjectManifest(),
            new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            });
    }

    private static bool IsEmptySharedCellValue(object value)
    {
        if (value == null || value == DBNull.Value)
        {
            return true;
        }

        string stringValue = value as string;
        if (stringValue != null)
        {
            return string.IsNullOrWhiteSpace(stringValue);
        }

        return false;
    }

    private static bool IsSharedSheetEmpty(SharedSheetDocument sharedSheetDocument)
    {
        if (sharedSheetDocument == null || sharedSheetDocument.Values == null)
        {
            return true;
        }

        foreach (object[] row in sharedSheetDocument.Values)
        {
            if (row == null)
            {
                continue;
            }

            foreach (object value in row)
            {
                if (!IsEmptySharedCellValue(value))
                {
                    return false;
                }
            }
        }

        return true;
    }

    private static bool IsSharedSheetRowEmpty(SharedSheetDocument sharedSheetDocument, object[] rowValues)
    {
        if (rowValues == null || rowValues.Length == 0)
        {
            return true;
        }

        var ignoreColumnOffsets = sharedSheetDocument == null ||
                                  sharedSheetDocument.RangeInfo == null ||
                                  sharedSheetDocument.RangeInfo.IgnoreColumnOffsets == null
            ? new HashSet<int>()
            : new HashSet<int>(sharedSheetDocument.RangeInfo.IgnoreColumnOffsets);

        for (int col = 0; col < rowValues.Length; col++)
        {
            if (ignoreColumnOffsets.Contains(col))
            {
                continue;
            }

            if (!IsEmptySharedCellValue(rowValues[col]))
            {
                return false;
            }
        }

        return true;
    }

    private static HashSet<string> CollectSharedExistingRowIds(params SharedSheetDocument[] documents)
    {
        var result = new HashSet<string>(StringComparer.Ordinal);

        foreach (SharedSheetDocument document in documents ?? new SharedSheetDocument[0])
        {
            if (!CanMergeSharedSheetByRowIds(document))
            {
                continue;
            }

            foreach (object rowIdValue in document.RowIds)
            {
                string rowId = NormalizeSharedRowId(rowIdValue);
                if (!string.IsNullOrWhiteSpace(rowId))
                {
                    result.Add(rowId);
                }
            }
        }

        return result;
    }

    private static SharedSheetDocument CreateCommitReadySharedSheetDocument(
        SharedSheetDocument localDocument,
        SharedSheetDocument baseDocument,
        SharedSheetDocument remoteDocument)
    {
        if (localDocument == null)
        {
            return null;
        }

        if (!CanMergeSharedSheetByRowIds(localDocument))
        {
            return localDocument;
        }

        HashSet<string> existingRowIds = CollectSharedExistingRowIds(baseDocument, remoteDocument);
        var keptRowIds = new List<object>();
        var keptRows = new List<object[]>();

        for (int i = 0; i < localDocument.RowIds.Length; i++)
        {
            string rowId = NormalizeSharedRowId(localDocument.RowIds[i]);
            object[] rowValues = i < localDocument.Values.Length
                ? (localDocument.Values[i] ?? new object[0])
                : new object[0];

            bool rowExists = !string.IsNullOrWhiteSpace(rowId) && existingRowIds.Contains(rowId);
            bool rowEmpty = IsSharedSheetRowEmpty(localDocument, rowValues);

            if (rowEmpty && !rowExists)
            {
                continue;
            }

            keptRowIds.Add(localDocument.RowIds[i]);
            keptRows.Add(rowValues);
        }

        if (keptRows.Count == 0)
        {
            return null;
        }

        var document = new SharedSheetDocument
        {
            Project = localDocument.Project,
            SheetId = localDocument.SheetId,
            SheetName = localDocument.SheetName,
            RangeAddress = localDocument.RangeAddress,
            RangeInfo = localDocument.RangeInfo == null ? null : new SharedRangeInfo
            {
                IdColumnOffset = localDocument.RangeInfo.IdColumnOffset,
                IgnoreColumnOffsets = localDocument.RangeInfo.IgnoreColumnOffsets == null
                    ? new HashSet<int>()
                    : new HashSet<int>(localDocument.RangeInfo.IgnoreColumnOffsets)
            },
            RowIds = keptRowIds.ToArray(),
            Values = keptRows.ToArray()
        };
        document.Hash = ComputeSharedSheetHash(document);
        return document;
    }

    private sealed class SharedSheetMergeResult
    {
        public SharedSheetDocument MergedDocument { get; set; }
        public int ConflictCount { get; set; }
    }

    private static bool CanMergeSharedSheetByRowIds(SharedSheetDocument document)
    {
        if (document == null ||
            document.Values == null ||
            document.RowIds == null)
        {
            return false;
        }

        if (document.RowIds.Length != document.Values.Length)
        {
            return false;
        }

        return HasAnyNonEmptySharedIds(document.RowIds);
    }

    private static string NormalizeSharedRowId(object rowId)
    {
        if (rowId == null)
        {
            return null;
        }

        string text = rowId.ToString();
        return string.IsNullOrWhiteSpace(text) ? null : text;
    }

    private static int GetSharedSheetColumnCount(SharedSheetDocument document)
    {
        if (document == null || document.Values == null || document.Values.Length == 0)
        {
            return 0;
        }

        return document.Values.Max(row => row == null ? 0 : row.Length);
    }

    private static object GetSharedSheetCellValue(object[] row, int columnIndex)
    {
        if (row == null || columnIndex < 0 || columnIndex >= row.Length)
        {
            return null;
        }

        return row[columnIndex];
    }

    private static Dictionary<string, object[]> CreateSharedSheetRowMap(SharedSheetDocument document)
    {
        var result = new Dictionary<string, object[]>(StringComparer.Ordinal);
        if (!CanMergeSharedSheetByRowIds(document))
        {
            return result;
        }

        for (int i = 0; i < document.RowIds.Length; i++)
        {
            string rowId = NormalizeSharedRowId(document.RowIds[i]);
            if (string.IsNullOrWhiteSpace(rowId))
            {
                continue;
            }

            result[rowId] = document.Values[i] ?? new object[0];
        }

        return result;
    }

    private static List<string> BuildSharedSheetRowOrder(
        SharedSheetDocument localDocument,
        SharedSheetDocument remoteDocument,
        SharedSheetDocument baseDocument)
    {
        var result = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        Action<SharedSheetDocument> append = document =>
        {
            if (!CanMergeSharedSheetByRowIds(document))
            {
                return;
            }

            foreach (object rowIdValue in document.RowIds)
            {
                string rowId = NormalizeSharedRowId(rowIdValue);
                if (string.IsNullOrWhiteSpace(rowId) || !seen.Add(rowId))
                {
                    continue;
                }

                result.Add(rowId);
            }
        };

        append(localDocument);
        append(remoteDocument);
        append(baseDocument);

        return result;
    }

    private static SharedSheetMergeResult TryMergeSharedSheetDocuments(
        SharedSheetDocument baseDocument,
        SharedSheetDocument localDocument,
        SharedSheetDocument remoteDocument)
    {
        if (localDocument == null)
        {
            return null;
        }

        int columnCount = Math.Max(
            GetSharedSheetColumnCount(localDocument),
            Math.Max(GetSharedSheetColumnCount(baseDocument), GetSharedSheetColumnCount(remoteDocument)));

        var ignoreColumnOffsets = localDocument.RangeInfo == null || localDocument.RangeInfo.IgnoreColumnOffsets == null
            ? new HashSet<int>()
            : new HashSet<int>(localDocument.RangeInfo.IgnoreColumnOffsets);

        bool useRowIdMerge =
            CanMergeSharedSheetByRowIds(localDocument) &&
            (baseDocument == null || CanMergeSharedSheetByRowIds(baseDocument)) &&
            (remoteDocument == null || CanMergeSharedSheetByRowIds(remoteDocument));

        if (!useRowIdMerge)
        {
            return new SharedSheetMergeResult
            {
                MergedDocument = localDocument,
                ConflictCount = 1
            };
        }

        var mergedRows = new List<object[]>();
        var mergedRowIds = new List<object>();
        int conflictCount = 0;

        Dictionary<string, object[]> localRows = CreateSharedSheetRowMap(localDocument);
        Dictionary<string, object[]> remoteRows = CreateSharedSheetRowMap(remoteDocument);
        Dictionary<string, object[]> baseRows = CreateSharedSheetRowMap(baseDocument);
        List<string> rowOrder = BuildSharedSheetRowOrder(localDocument, remoteDocument, baseDocument);

        foreach (string rowId in rowOrder)
        {
            object[] localRow;
            localRows.TryGetValue(rowId, out localRow);

            object[] remoteRow;
            remoteRows.TryGetValue(rowId, out remoteRow);

            object[] baseRow;
            baseRows.TryGetValue(rowId, out baseRow);

            var mergedRow = new object[columnCount];
            for (int col = 0; col < columnCount; col++)
            {
                object baseValue = GetSharedSheetCellValue(baseRow, col);
                object localValue = localRow == null ? baseValue : GetSharedSheetCellValue(localRow, col);
                object remoteValue = remoteRow == null ? baseValue : GetSharedSheetCellValue(remoteRow, col);

                if (ignoreColumnOffsets.Contains(col))
                {
                    mergedRow[col] = NormalizeSharedCellValue(localValue);
                    continue;
                }

                if (AreSharedCellValuesEqual(localValue, baseValue))
                {
                    mergedRow[col] = NormalizeSharedCellValue(remoteValue);
                }
                else if (AreSharedCellValuesEqual(remoteValue, baseValue))
                {
                    mergedRow[col] = NormalizeSharedCellValue(localValue);
                }
                else if (AreSharedCellValuesEqual(localValue, remoteValue))
                {
                    mergedRow[col] = NormalizeSharedCellValue(localValue);
                }
                else
                {
                    mergedRow[col] = NormalizeSharedCellValue(localValue);
                    conflictCount++;
                }
            }

            mergedRows.Add(mergedRow);
            mergedRowIds.Add(rowId);
        }

        if (mergedRows.Count == 0 && localDocument.Values != null && localDocument.Values.Length > 0)
        {
            return new SharedSheetMergeResult
            {
                MergedDocument = localDocument,
                ConflictCount = 1
            };
        }

        var mergedDocument = new SharedSheetDocument
        {
            Project = localDocument.Project,
            SheetId = localDocument.SheetId,
            SheetName = localDocument.SheetName,
            RangeAddress = localDocument.RangeAddress,
            RangeInfo = localDocument.RangeInfo == null ? null : new SharedRangeInfo
            {
                IdColumnOffset = localDocument.RangeInfo.IdColumnOffset,
                IgnoreColumnOffsets = localDocument.RangeInfo.IgnoreColumnOffsets == null
                    ? new HashSet<int>()
                    : new HashSet<int>(localDocument.RangeInfo.IgnoreColumnOffsets)
            },
            RowIds = mergedRowIds.Count == 0
                ? (localDocument.RowIds ?? new object[0])
                : mergedRowIds.Cast<object>().ToArray(),
            Values = mergedRows.ToArray()
        };
        mergedDocument.Hash = ComputeSharedSheetHash(mergedDocument);

        return new SharedSheetMergeResult
        {
            MergedDocument = mergedDocument,
            ConflictCount = conflictCount
        };
    }

    private static SharedSheetMergeResult TryMergeSharedSheetDocumentForReceive(
        SharedSheetDocument baseDocument,
        SharedSheetDocument localDocument,
        SharedSheetDocument remoteDocument)
    {
        if (localDocument == null || remoteDocument == null)
        {
            return null;
        }

        if (!CanMergeSharedSheetByRowIds(localDocument) ||
            !CanMergeSharedSheetByRowIds(remoteDocument))
        {
            return new SharedSheetMergeResult
            {
                MergedDocument = localDocument,
                ConflictCount = 1
            };
        }

        if (baseDocument != null && !CanMergeSharedSheetByRowIds(baseDocument))
        {
            baseDocument = null;
        }

        int columnCount = Math.Max(
            GetSharedSheetColumnCount(localDocument),
            Math.Max(GetSharedSheetColumnCount(baseDocument), GetSharedSheetColumnCount(remoteDocument)));

        var ignoreColumnOffsets = localDocument.RangeInfo == null || localDocument.RangeInfo.IgnoreColumnOffsets == null
            ? new HashSet<int>()
            : new HashSet<int>(localDocument.RangeInfo.IgnoreColumnOffsets);

        Dictionary<string, object[]> localRows = CreateSharedSheetRowMap(localDocument);
        Dictionary<string, object[]> remoteRows = CreateSharedSheetRowMap(remoteDocument);
        Dictionary<string, object[]> baseRows = CreateSharedSheetRowMap(baseDocument);
        var mergedRows = new List<object[]>();
        var mergedRowIds = new List<object>();
        int conflictCount = 0;

        foreach (object rowIdValue in localDocument.RowIds)
        {
            string rowId = NormalizeSharedRowId(rowIdValue);
            if (string.IsNullOrWhiteSpace(rowId))
            {
                continue;
            }

            object[] localRow;
            localRows.TryGetValue(rowId, out localRow);

            object[] remoteRow;
            bool hasRemoteRow = remoteRows.TryGetValue(rowId, out remoteRow);

            object[] baseRow;
            baseRows.TryGetValue(rowId, out baseRow);

            var mergedRow = new object[columnCount];
            for (int col = 0; col < columnCount; col++)
            {
                object baseValue = GetSharedSheetCellValue(baseRow, col);
                object localValue = GetSharedSheetCellValue(localRow, col);
                object remoteValue = hasRemoteRow
                    ? GetSharedSheetCellValue(remoteRow, col)
                    : baseValue;

                if (ignoreColumnOffsets.Contains(col))
                {
                    mergedRow[col] = NormalizeSharedCellValue(localValue);
                    continue;
                }

                if (AreSharedCellValuesEqual(localValue, baseValue))
                {
                    mergedRow[col] = NormalizeSharedCellValue(remoteValue);
                }
                else if (AreSharedCellValuesEqual(remoteValue, baseValue))
                {
                    mergedRow[col] = NormalizeSharedCellValue(localValue);
                }
                else if (AreSharedCellValuesEqual(localValue, remoteValue))
                {
                    mergedRow[col] = NormalizeSharedCellValue(localValue);
                }
                else
                {
                    mergedRow[col] = NormalizeSharedCellValue(remoteValue);
                    conflictCount++;
                }
            }

            mergedRows.Add(mergedRow);
            mergedRowIds.Add(rowId);
        }

        var mergedDocument = new SharedSheetDocument
        {
            Project = localDocument.Project,
            SheetId = localDocument.SheetId,
            SheetName = localDocument.SheetName,
            RangeAddress = localDocument.RangeAddress,
            RangeInfo = localDocument.RangeInfo == null ? null : new SharedRangeInfo
            {
                IdColumnOffset = localDocument.RangeInfo.IdColumnOffset,
                IgnoreColumnOffsets = localDocument.RangeInfo.IgnoreColumnOffsets == null
                    ? new HashSet<int>()
                    : new HashSet<int>(localDocument.RangeInfo.IgnoreColumnOffsets)
            },
            RowIds = mergedRowIds.ToArray(),
            Values = mergedRows.ToArray()
        };
        mergedDocument.Hash = ComputeSharedSheetHash(mergedDocument);

        return new SharedSheetMergeResult
        {
            MergedDocument = mergedDocument,
            ConflictCount = conflictCount
        };
    }

    private static string GetRemoteSharedHash(SharedProjectManifest manifest, string sheetId)
    {
        if (manifest == null || manifest.Sheets == null || string.IsNullOrWhiteSpace(sheetId))
        {
            return null;
        }

        SharedProjectManifestEntry entry = manifest.Sheets.FirstOrDefault(x => string.Equals(x.SheetId, sheetId, StringComparison.Ordinal));
        return entry == null ? null : entry.Hash;
    }

    private static List<SharedSheetSelectionItem> CollectSharedSheetSelectionItems(
        Excel.Workbook workbook,
        SharedProjectManifest remoteManifest,
        string targetSheetName = null)
    {
        var items = new List<SharedSheetSelectionItem>();
        foreach (SharedSheetDocument document in CollectSharedSheetDocuments(workbook))
        {
            if (!string.IsNullOrWhiteSpace(targetSheetName) &&
                !string.Equals(document.SheetName, targetSheetName, StringComparison.CurrentCulture))
            {
                continue;
            }

            SharedSheetDocument baseDocument = GetSharedSheetBaseDocument(workbook, document.SheetId);
            SharedSheetDocument commitDocument = CreateCommitReadySharedSheetDocument(document, baseDocument, null);
            if (commitDocument == null)
            {
                continue;
            }

            string baseHash = GetSharedSheetBaseHash(workbook, document.SheetId);
            string remoteHash = GetRemoteSharedHash(remoteManifest, document.SheetId);

            if (string.Equals(commitDocument.Hash, baseHash, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(remoteHash) &&
                string.IsNullOrWhiteSpace(baseHash) &&
                IsSharedSheetEmpty(commitDocument))
            {
                continue;
            }

            items.Add(new SharedSheetSelectionItem
            {
                Selected = true,
                SheetName = commitDocument.SheetName,
                SheetId = commitDocument.SheetId,
                ActionLabel = string.IsNullOrWhiteSpace(remoteHash) ? "新規" : "更新",
                StatusDetail = string.IsNullOrWhiteSpace(remoteHash)
                    ? "共有先にまだありません"
                    : "ローカル変更があります",
                DiffText = BuildSharedSheetDiffText(
                    baseDocument,
                    commitDocument,
                    string.IsNullOrWhiteSpace(remoteHash) ? null : baseDocument),
                Document = commitDocument
            });
        }

        return items
            .OrderBy(x => x.SheetName, StringComparer.CurrentCulture)
            .ToList();
    }

    private static List<string> CollectStaleSharedSheetNames(
        Excel.Workbook workbook,
        IEnumerable<SharedSheetSelectionItem> items,
        SharedProjectManifest remoteManifest)
    {
        var staleSheetNames = new List<string>();

        foreach (SharedSheetSelectionItem item in items ?? Enumerable.Empty<SharedSheetSelectionItem>())
        {
            if (item == null || item.Document == null || string.IsNullOrWhiteSpace(item.SheetId))
            {
                continue;
            }

            string baseHash = GetSharedSheetBaseHash(workbook, item.SheetId);
            string remoteHash = GetRemoteSharedHash(remoteManifest, item.SheetId);

            if (string.IsNullOrWhiteSpace(remoteHash))
            {
                continue;
            }

            if (!string.Equals(baseHash, remoteHash, StringComparison.OrdinalIgnoreCase))
            {
                staleSheetNames.Add(item.SheetName);
            }
        }

        return staleSheetNames
            .Distinct(StringComparer.CurrentCulture)
            .OrderBy(x => x, StringComparer.CurrentCulture)
            .ToList();
    }

    private async Task<List<string>> MergeSharedSelectionItemsWithRemoteAsync(
        Excel.Workbook workbook,
        GitLabShareInfo shareInfo,
        string token,
        IEnumerable<SharedSheetSelectionItem> items,
        SharedProjectManifest remoteManifest)
    {
        var conflictSheetNames = new List<string>();

        foreach (SharedSheetSelectionItem item in items ?? Enumerable.Empty<SharedSheetSelectionItem>())
        {
            if (item == null || item.Document == null || string.IsNullOrWhiteSpace(item.SheetId))
            {
                continue;
            }

            string baseHash = GetSharedSheetBaseHash(workbook, item.SheetId);
            string remoteHash = GetRemoteSharedHash(remoteManifest, item.SheetId);

            if (string.IsNullOrWhiteSpace(remoteHash) ||
                string.Equals(baseHash, remoteHash, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            SharedSheetDocument localDocument = item.Document;
            SharedSheetDocument baseDocument = GetSharedSheetBaseDocument(workbook, item.SheetId);
            SharedSheetDocument remoteDocument = await TryDownloadSharedSheetDocumentAsync(
                shareInfo,
                workbook.GetCustomProperty(ssProjectIdCustomPropertyName),
                item.SheetId,
                token).ConfigureAwait(true);

            if (remoteDocument == null)
            {
                item.HasConflict = true;
                item.ActionLabel = "競合";
                item.StatusDetail = "共有先シートを取得できません";
                item.DiffText = BuildSharedSheetDiffText(baseDocument, item.Document, null);
                conflictSheetNames.Add(item.SheetName);
                continue;
            }

            SharedSheetMergeResult mergeResult = TryMergeSharedSheetDocuments(baseDocument, localDocument, remoteDocument);
            if (mergeResult == null || mergeResult.ConflictCount > 0 || mergeResult.MergedDocument == null)
            {
                item.HasConflict = true;
                item.ActionLabel = "競合";
                item.StatusDetail = mergeResult == null
                    ? "競合判定に失敗しました"
                    : ("競合セル " + mergeResult.ConflictCount + " 件");
                item.DiffText = BuildSharedSheetDiffText(baseDocument, localDocument, remoteDocument);
                conflictSheetNames.Add(item.SheetName);
                continue;
            }

            SharedSheetDocument commitDocument = CreateCommitReadySharedSheetDocument(
                mergeResult.MergedDocument,
                baseDocument,
                remoteDocument);

            if (commitDocument == null ||
                string.Equals(commitDocument.Hash, remoteHash, StringComparison.OrdinalIgnoreCase))
            {
                item.Selected = false;
                item.ActionLabel = "対象外";
                item.StatusDetail = "共有先と同じため送信しません";
                item.DiffText = BuildSharedSheetDiffText(baseDocument, localDocument, remoteDocument);
                item.Document = null;
                continue;
            }

            item.Document = commitDocument;
            item.ActionLabel = "マージ";
            item.StatusDetail = "共有先変更を取り込みます";
            item.DiffText = BuildSharedSheetDiffText(baseDocument, localDocument, remoteDocument);
        }

        return conflictSheetNames
            .Distinct(StringComparer.CurrentCulture)
            .OrderBy(x => x, StringComparer.CurrentCulture)
            .ToList();
    }

    private static SharedProjectManifest MergeSharedProjectManifest(
        SharedProjectManifest remoteManifest,
        string projectId,
        IEnumerable<SharedSheetDocument> updatedDocuments)
    {
        var entries = new Dictionary<string, SharedProjectManifestEntry>(StringComparer.Ordinal);

        if (remoteManifest != null && remoteManifest.Sheets != null)
        {
            foreach (SharedProjectManifestEntry entry in remoteManifest.Sheets)
            {
                if (entry == null || string.IsNullOrWhiteSpace(entry.SheetId))
                {
                    continue;
                }

                entries[entry.SheetId] = new SharedProjectManifestEntry
                {
                    SheetId = entry.SheetId,
                    SheetName = entry.SheetName,
                    Hash = entry.Hash
                };
            }
        }

        foreach (SharedSheetDocument document in updatedDocuments ?? Enumerable.Empty<SharedSheetDocument>())
        {
            if (document == null || string.IsNullOrWhiteSpace(document.SheetId))
            {
                continue;
            }

            entries[document.SheetId] = new SharedProjectManifestEntry
            {
                SheetId = document.SheetId,
                SheetName = document.SheetName,
                Hash = document.Hash
            };
        }

        return new SharedProjectManifest
        {
            Project = projectId,
            UpdatedAt = DateTime.UtcNow.ToString("o"),
            Sheets = entries.Values
                .OrderBy(x => x.SheetId, StringComparer.Ordinal)
                .ToList()
        };
    }

    private async Task UploadSharedSheetsAsync(
        Excel.Workbook workbook,
        GitLabShareInfo shareInfo,
        string token,
        SharedProjectManifest remoteManifest,
        IEnumerable<SharedSheetSelectionItem> selectedItems,
        Action<string> progressReporter = null)
    {
        if (workbook == null)
        {
            throw new ArgumentNullException(nameof(workbook));
        }

        if (shareInfo == null)
        {
            throw new ArgumentNullException(nameof(shareInfo));
        }

        List<SharedSheetSelectionItem> items = (selectedItems ?? Enumerable.Empty<SharedSheetSelectionItem>())
            .Where(x => x != null && x.Document != null)
            .ToList();
        if (items.Count == 0)
        {
            return;
        }

        string projectId = workbook.GetCustomProperty(ssProjectIdCustomPropertyName);
        if (string.IsNullOrWhiteSpace(token))
        {
            MessageBox.Show("共有をキャンセルしました（トークン未入力）", "変更共有");
            return;
        }
        string refName = GetNormalizedShareRefName(shareInfo);

        var actions = new List<object>();

        foreach (SharedSheetSelectionItem item in items)
        {
            SharedSheetDocument document = item.Document;
            string filePath = BuildSharedSheetPath(projectId, document.SheetId);
            string actionName = string.IsNullOrWhiteSpace(GetRemoteSharedHash(remoteManifest, document.SheetId))
                ? "create"
                : "update";
            progressReporter?.Invoke("共有しています: " + item.SheetName);

            actions.Add(new Dictionary<string, object>
            {
                { "action", actionName },
                { "file_path", filePath },
                { "content", CreateSharedSheetUploadJsonText(document) },
                { "encoding", "text" }
            });
        }

        SharedProjectManifest updatedManifest = MergeSharedProjectManifest(
            remoteManifest,
            projectId,
            items.Select(x => x.Document));

        progressReporter?.Invoke("共有マニフェストを更新しています");
        actions.Add(new Dictionary<string, object>
        {
            { "action", remoteManifest == null ? "create" : "update" },
            { "file_path", BuildSharedProjectManifestPath(projectId) },
            { "content", CreateSharedProjectManifestJsonText(updatedManifest) },
            { "encoding", "text" }
        });

        progressReporter?.Invoke("GitLab の応答を待っています（数秒かかることがあります）");
        await GitLabClient.CreateCommitAsync(
            shareInfo.BaseUrl,
            shareInfo.ProjectId,
            refName,
            token,
            "Update shared sheets: " + projectId,
            actions).ConfigureAwait(true);

        progressReporter?.Invoke("ローカル base を更新しています");
        EnsureSharedSheetBaseStorePrepared(workbook, progressReporter);
        ExcelUiSuspendScope uiSuspendScope = TryCreateExcelUiSuspendScope(workbook);
        if (uiSuspendScope == null)
        {
            foreach (SharedSheetSelectionItem item in items ?? Enumerable.Empty<SharedSheetSelectionItem>())
            {
                if (item == null || item.Document == null || string.IsNullOrWhiteSpace(item.SheetId))
                {
                    continue;
                }

                SaveSharedSheetBaseDocument(workbook, item.Document);
            }
        }
        else
        {
            using (uiSuspendScope)
            {
                foreach (SharedSheetSelectionItem item in items ?? Enumerable.Empty<SharedSheetSelectionItem>())
                {
                    if (item == null || item.Document == null || string.IsNullOrWhiteSpace(item.SheetId))
                    {
                        continue;
                    }

                    SaveSharedSheetBaseDocument(workbook, item.Document);
                }
            }
        }

        foreach (SharedSheetSelectionItem item in items)
        {
            FileLogger.Info("[SharedCommit] uploaded sheetId=" + item.SheetId + " hash=" + item.Document.Hash);
        }
    }

    class SheetViewState
    {
        public string SheetId { get; set; }
        public string SheetName { get; set; }
        public (int row, int column)? ActiveCellPosition { get; set; }
        public (int horizontalScroll, int verticalScroll)? ScrollPosition { get; set; }
        public double Zoom { get; set; }
    }

    async Task ForceUpdateSheet(Excel.Worksheet sheet, string txtFilePathOverride = null)
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
        string storedSourceFilePath = NormalizeSourceFilePath(lastRenderLog.SourceFilePath);
        bool isPullWorkSourceFile = IsPullWorkSourceFilePath(storedSourceFilePath);
        bool sourceFileExists = CanReuseStoredSourceFilePath(storedSourceFilePath);
        string txtFilePath = txtFilePathOverride;

        if (txtFilePath == null)
        {
            if (!isSameUser || !sourceFileExists)
            {
                string message = null;
                if (!isSameUser)
                {
                    message = "最後に更新された環境と異なります。";
                }
                else if (isPullWorkSourceFile)
                {
                    message = "Pull で取得した一時ファイルはローカル更新に使用できません。";
                }
                else if (!sourceFileExists)
                {
                    message = "ソースファイルが見つかりません。";
                }

                DialogResult fileSelectionResult = MessageBox.Show($"{message}\nProject ID が「{projectId}」の TXT を選択し直してください。", "確認", MessageBoxButtons.OKCancel);
                if (fileSelectionResult != DialogResult.OK)
                {
                    return;
                }

                txtFilePath = OpenSourceFile();
                // キャンセルされたら何もしない
                if (txtFilePath == null)
                {
                    return;
                }
            }
            else
            {
                txtFilePath = storedSourceFilePath;
            }
        }

        string jsonFilePath = TxtToJsonPath(txtFilePath);
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
            SourceFilePath = txtFilePath,
            User = Environment.UserName
        };
        workbook.SetCustomProperty("RenderLog", renderLog);
    }

    private static bool IsExcelBusyComException(COMException ex)
    {
        return ex != null && ex.HResult == unchecked((int)0x8001010A);
    }

    private static T ExecuteExcelComWithRetry<T>(Func<T> action)
    {
        COMException lastException = null;

        for (int attempt = 0; attempt < 8; attempt++)
        {
            try
            {
                return action();
            }
            catch (COMException ex) when (IsExcelBusyComException(ex))
            {
                lastException = ex;
                Thread.Sleep(100);
            }
        }

        throw new InvalidOperationException("セルの編集を確定してから実行してください。", lastException);
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

        string txtFilePath = SelectSourceFileForParse(false);
        if (txtFilePath == null)
        {
            return;
        }

        FileLogger.InitializeForInput(txtFilePath, timestamped: false);
        bool parseSucceeded = RunParsePipeline(txtFilePath, true);
        if (!parseSucceeded)
        {
            return;
        }

        excelApp.ScreenUpdating = false;
        excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
        excelApp.EnableEvents = false;
        MacroControl.DisableMacros(excelApp);

        await ForceUpdateSheet(sheet, txtFilePath);

        excelApp.StatusBar = false;
        excelApp.ScreenUpdating = true;
        excelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
        excelApp.EnableEvents = true;
        MacroControl.EnableMacros(excelApp);

    }

    Dictionary<string, SheetViewState> CaptureViewStatesForUpdate(
        Excel.Application excelApp,
        Excel.Workbook workbook,
        JsonArray sheetNodes,
        Dictionary<string, string> sheetNamesById,
        Excel.Worksheet originalActiveSheet,
        (int row, int column)? originalActiveCellPosition,
        (int horizontalScroll, int verticalScroll)? originalScrollPosition,
        double originalZoom)
    {
        var viewStates = new Dictionary<string, SheetViewState>();

        foreach (JsonNode sheetNode in sheetNodes)
        {
            string id = sheetNode["id"].ToString();

            if (!sheetNamesById.ContainsKey(id))
            {
                continue;
            }

            string sheetName = sheetNamesById[id];
            Excel.Worksheet sheet = workbook.Sheets[sheetName];

            try
            {
                sheet.Activate();

                var viewState = new SheetViewState
                {
                    SheetId = id,
                    SheetName = sheetName,
                    ActiveCellPosition = excelApp.GetActiveCellPosition(),
                    ScrollPosition = excelApp.GetScrollPosition(),
                    Zoom = excelApp.GetActiveSheetZoom(),
                };

                viewStates[id] = viewState;
            }
            catch (Exception ex)
            {
                FileLogger.Warn($"Failed to capture view state for sheet '{sheetName}': {ex}");
            }
        }

        try
        {
            if (originalActiveSheet != null)
            {
                originalActiveSheet.Activate();
                excelApp.SetActiveCellPosition(originalActiveCellPosition);
                excelApp.SetActiveSheetZoom(originalZoom);
                excelApp.SetScrollPosition(originalScrollPosition);
            }
        }
        catch (Exception ex)
        {
            FileLogger.Warn($"Failed to restore original active sheet after view state capture: {ex}");
        }

        return viewStates;
    }

    SheetViewState FindViewState(Dictionary<string, SheetViewState> viewStates, string sheetId, string sheetName)
    {
        if (sheetId != null)
        {
            if (viewStates.TryGetValue(sheetId, out SheetViewState viewStateById))
            {
                return viewStateById;
            }
        }

        if (sheetName != null)
        {
            return viewStates.Values.FirstOrDefault(state => state.SheetName == sheetName);
        }

        return null;
    }

    static void ApplyViewState(Excel.Application excelApp, Excel.Worksheet sheet, SheetViewState viewState)
    {
        if (viewState == null)
        {
            return;
        }

        try
        {
            sheet.Activate();

            excelApp.SetActiveCellPosition(viewState.ActiveCellPosition);
            excelApp.SetActiveSheetZoom(viewState.Zoom);
            excelApp.SetScrollPosition(viewState.ScrollPosition);

            var nudgeZoom = excelApp.ActiveWindow.Zoom + 1;
            excelApp.ActiveWindow.Zoom = nudgeZoom;
            excelApp.ActiveWindow.Zoom = viewState.Zoom;
        }
        catch (Exception ex)
        {
            FileLogger.Warn($"Failed to apply view state for sheet '{sheet.Name}': {ex}");
        }
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

    string ComputeConfHash(JsonNode jsonObject)
    {
        JsonNode variablesNode = jsonObject?["variables"];

        if (variablesNode == null)
        {
            return string.Empty;
        }

        return variablesNode.ComputeSha256();
    }

    async Task UpdateAllSheets(Excel.Workbook workbook, string txtFilePathOverride = null, string jsonFilePathOverride = null)
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

        string txtFilePath = txtFilePathOverride;
        if (txtFilePath == null)
        {
            var lastRenderLog = workbookInfo.LastRenderLog;
            bool isSameUser = lastRenderLog.User == Environment.UserName;
            string storedSourceFilePath = NormalizeSourceFilePath(lastRenderLog.SourceFilePath);
            bool isPullWorkSourceFile = IsPullWorkSourceFilePath(storedSourceFilePath);
            bool sourceFileExists = CanReuseStoredSourceFilePath(storedSourceFilePath);

            if (!isSameUser || !sourceFileExists)
            {
                string message = !isSameUser
                    ? "最後に更新された環境と異なります。"
                    : isPullWorkSourceFile
                        ? "Pull で取得した一時ファイルはローカル更新に使用できません。"
                        : "ソースファイルが見つかりません。";

                DialogResult fileSelectionResult = MessageBox.Show(
                    $"{message}\nProject ID が「{projectId}」の TXT を選択し直してください。", "確認",
                    MessageBoxButtons.OKCancel);

                if (fileSelectionResult != DialogResult.OK)
                {
                    return;
                }

                txtFilePath = OpenSourceFile();
                if (txtFilePath == null)
                {
                    return;
                }
            }
            else
            {
                txtFilePath = storedSourceFilePath;
            }
        }

        string jsonFilePath = jsonFilePathOverride ?? TxtToJsonPath(txtFilePath);
        string jsonString = File.ReadAllText(jsonFilePath);
        JsonNode jsonObject = JsonNode.Parse(jsonString);
        var confData = GetPropertiesFromJsonNode(jsonObject, "variables");

        if (confData["project"] != projectId)
        {
            MessageBox.Show($"Project ID({projectId})が異なります。");
            return;
        }

        string newConfHash = ComputeConfHash(jsonObject);
        string storedConfHash = workbook.GetCustomProperty(confHashCustomPropertyName);
        bool forceRenderAllSheets = storedConfHash != newConfHash;

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

        var indexSheet = workbook.Sheets[workbookInfo.IndexSheetName] as Excel.Worksheet;
        var sheetIds = GetSheetIdsFromIndexSheet(indexSheet).Select(item => item.ToString()).ToList();
        var sheetNameRange = GetSheetNamesRangeFromIndexSheet(indexSheet);
        var sheetNames = sheetNameRange.GetColumnValues(0).Select(item => item.ToString()).ToList();

        // 特定のプロパティ（items）を配列としてアクセス
        JsonArray sheetNodes = jsonObject["children"].AsArray();

        int currentSheetCount = sheetNames.Count;
        int newSheetCount = sheetNodes.Count;

        if (currentSheetCount == 1 && newSheetCount > 1)
        {
            FileLogger.Info($"Auto-switch to regeneration: 1 -> {newSheetCount}");
            await RegenerateWorkbook(workbook, workbookInfo, txtFilePath, jsonFilePath);
            return;
        }

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
        Dictionary<string, string> originalSheetNamesById = sheetIds.Zip(sheetNames, (id, name) => new { id, name })
                                                                        .ToDictionary(x => x.id, x => x.name);

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

        var sheetViewStates = CaptureViewStatesForUpdate(
            excelApp,
            workbook,
            sheetNodes,
            originalSheetNamesById,
            activeSheet,
            activeCellPosition,
            scrollPosition,
            activeSheetZoom);

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

                    // confData が変わっていなければ、元データと画像の差分を見て生成要否を判断する
                    string newSheetHash = await newSheetHashTasks[id];
                    if (!forceRenderAllSheets && newSheetHash == sheetHash)
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
                    string sheetName = newSheet.Name;
                    var viewState = FindViewState(sheetViewStates, id, sheetName);
                    ApplyViewState(excelApp, newSheet, viewState);
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

        // RenderLog を更新して次回以降に今回選択したソースを利用できるようにする
        RenderLog renderLog = new RenderLog
        {
            SourceFilePath = txtFilePath,
            User = Environment.UserName
        };
        workbook.SetCustomProperty("RenderLog", renderLog);
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

    async Task<List<(string filePath, string sheetName, string address)>> RenderWorkbook(
        Excel.Workbook workbook,
        JsonNode jsonObject,
        Dictionary<string, string> confData,
        string jsonFilePath,
        Dictionary<string, SheetValuesInfo> sheetValuesById)
    {
        Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;

        string indexSheetName = workbook.GetCustomProperty(indexSheetNameCustomPropertyName);
        string templateSheetName = workbook.GetCustomProperty(templateSheetNameCustomPropertyName);

        if (indexSheetName == null)
        {
            MessageBox.Show($"カスタムプロパティに {indexSheetNameCustomPropertyName} が設定されていません。");
            return new List<(string filePath, string sheetName, string address)>();
        }
        if (templateSheetName == null)
        {
            MessageBox.Show($"カスタムプロパティに {templateSheetNameCustomPropertyName} が設定されていません。");
            return new List<(string filePath, string sheetName, string address)>();
        }

        excelApp.DisplayAlerts = false;
        excelApp.ScreenUpdating = false;
        excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
        excelApp.EnableEvents = false;
        excelApp.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityForceDisable;

        JsonArray sheetNodes = jsonObject["children"].AsArray();

        Excel.Worksheet templateSheet = workbook.Sheets[templateSheetName];
        var templateSheetVisible = templateSheet.Visible;
        templateSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

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
                string sheetId = sheetNode["id"].ToString();

                // プログレスバーを更新
                progressBarForm.Invoke(new Action<string>(progressBarForm.UpdateSheetName), newSheetName);

                // シートをコピーしてリネーム
                templateSheet.Copy(After: workbook.Sheets[workbook.Sheets.Count]);
                Excel.Worksheet newSheet = workbook.Sheets[workbook.Sheets.Count];
                newSheet.Name = newSheetName;

                SheetValuesInfo sheetValuesInfo = null;
                sheetValuesById?.TryGetValue(sheetId, out sheetValuesInfo);

                var missingImagePathsInSheet = RenderSheet(sheetNode, confData, jsonFilePath, newSheet, sheetValuesInfo);

                // シートの JsonNode の hash をカスタムプロパティに保存
                string sheetHash = await sheetHashTask;
                newSheet.SetCustomProperty(sheetHashCustomPropertyName, sheetHash);

                var newSheetImageHash = await newSheetImageHashTask;
                newSheet.SetCustomProperty(sheetImageHashCustomPropertyName, newSheetImageHash);

                missingImagePaths.AddRange(missingImagePathsInSheet);

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

            string confHash = ComputeConfHash(jsonObject);
            workbook.SetCustomProperty(confHashCustomPropertyName, confHash);

            // 最後にindexシートを選択状態にしておく
            indexSheet.Activate();

            templateSheet.Visible = templateSheetVisible;

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

        return missingImagePaths;
    }

    static string CreateBackupFilePath(string originalPath)
    {
        string directory = Path.GetDirectoryName(originalPath);
        string bakDirectory = Path.Combine(directory, "bak");

        // bak フォルダがなければ作成（既にあっても例外は出ない）
        Directory.CreateDirectory(bakDirectory);

        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(originalPath);
        string extension = Path.GetExtension(originalPath);
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

        return Path.Combine(
            bakDirectory,
            $"{fileNameWithoutExtension}.backup_{timestamp}{extension}"
        );
    }

    static string CreateTemporaryRegenerationPath(string originalPath)
    {
        string directory = Path.GetDirectoryName(originalPath);
        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(originalPath);
        string extension = Path.GetExtension(originalPath);

        return Path.Combine(directory, $"{fileNameWithoutExtension}.__regen__{extension}");
    }

    static Dictionary<string, SheetValuesInfo> SnapshotSheetValuesById(Excel.Workbook workbook)
    {
        var result = new Dictionary<string, SheetValuesInfo>();

        foreach (Excel.Worksheet sheet in workbook.Sheets)
        {
            string sheetId = sheet.GetCustomProperty(sheetIdCustomPropertyName);
            if (string.IsNullOrEmpty(sheetId))
            {
                continue;
            }

            var sheetValuesInfo = SheetValuesInfo.CreateFromSheet(sheet);
            result[sheetId] = sheetValuesInfo;
        }

        return result;
    }

    async Task RegenerateWorkbook(Excel.Workbook originalWorkbook, WorkbookInfo workbookInfo, string txtFilePath, string jsonFilePath)
    {
        Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;

        if (originalWorkbook.Saved == false)
        {
            originalWorkbook.Save();
        }

        string jsonString = File.ReadAllText(jsonFilePath);
        JsonNode jsonObject = JsonNode.Parse(jsonString);
        var confData = GetPropertiesFromJsonNode(jsonObject, "variables");

        if (confData["project"] != workbookInfo.ProjectId)
        {
            MessageBox.Show($"Project ID({workbookInfo.ProjectId})が異なります。");
            return;
        }

        var sheetValuesById = SnapshotSheetValuesById(originalWorkbook);

        string originalPath = originalWorkbook.FullName;
        string tempPath = CreateTemporaryRegenerationPath(originalPath);

        if (File.Exists(tempPath))
        {
            File.Delete(tempPath);
        }

        string templateFilePath = GetAbsolutePathFromExecutingDirectory(templateFileName);
        Excel.Workbook newWorkbook = null;

        try
        {
            newWorkbook = CreateCopiedWorkbook(excelApp, templateFilePath, tempPath);

            var missingImagePaths = await RenderWorkbook(newWorkbook, jsonObject, confData, jsonFilePath, sheetValuesById);

            // RenderLog は TXT を保存
            RenderLog renderLog = new RenderLog
            {
                SourceFilePath = txtFilePath,   // TXT を保存
                User = Environment.UserName
            };
            newWorkbook.SetCustomProperty("RenderLog", renderLog);

            string projectId = confData["project"];
            newWorkbook.SetCustomProperty(ssProjectIdCustomPropertyName, projectId);

            newWorkbook.Save();

            newWorkbook.Close(false);

            string backupPath = CreateBackupFilePath(originalPath);

            originalWorkbook.Close(false);

            File.Move(originalPath, backupPath);

            try
            {
                File.Move(tempPath, originalPath);
            }
            catch
            {
                File.Move(backupPath, originalPath);
                throw;
            }

            Excel.Workbook reopenedWorkbook = excelApp.Workbooks.Open(originalPath);
            reopenedWorkbook.Activate();

            if (missingImagePaths.Any())
            {
                ShowMissingImageFilesDialog(missingImagePaths);
            }
        }
        catch (Exception ex)
        {
            FileLogger.Error($"RegenerateWorkbook failed: {ex}");
            MessageBox.Show($"再生成に失敗しました: {ex.Message}", "再生成", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (newWorkbook != null)
            {
                try
                {
                    newWorkbook.Close(false);
                }
                catch
                {
                }
            }

            if (File.Exists(tempPath))
            {
                try
                {
                    File.Delete(tempPath);
                }
                catch
                {
                }
            }
        }
    }

    private static string SanitizeFolderName(string folderName, string fallbackName)
    {
        string name = string.IsNullOrWhiteSpace(folderName) ? fallbackName : folderName.Trim();
        char[] invalidChars = Path.GetInvalidFileNameChars();
        var builder = new StringBuilder(name.Length);

        foreach (char c in name)
        {
            if (invalidChars.Contains(c))
            {
                builder.Append('_');
            }
            else
            {
                builder.Append(c);
            }
        }

        string sanitized = builder.ToString().Trim();
        return string.IsNullOrWhiteSpace(sanitized) ? fallbackName : sanitized;
    }

    private static string GetPullWorkbookOutputDirectory(string projectId, string projectFolderName = null)
    {
        string documentsDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        string safeProjectId = string.IsNullOrWhiteSpace(projectId) ? "UnknownProject" : projectId.Trim();
        string safeProjectFolderName = SanitizeFolderName(projectFolderName, safeProjectId);
        return Path.Combine(documentsDirectory, "SheetRenderer", safeProjectFolderName);
    }

    private static bool ConfirmOverwriteForPullNewWorkbook(string outputFilePath)
    {
        if (string.IsNullOrWhiteSpace(outputFilePath))
        {
            return true;
        }

        string fullPath = GetNormalizedWorkbookOutputPath(outputFilePath);
        if (!File.Exists(fullPath))
        {
            return true;
        }

        DialogResult overwriteResult = MessageBox.Show(
            "同名のファイルが既に存在します。上書きしますか？\n\n" + fullPath,
            "Pull 新規作成",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);
        return overwriteResult == DialogResult.Yes;
    }

    private static string GetNormalizedWorkbookOutputPath(string outputFilePath)
    {
        string fullPath = Path.GetFullPath(outputFilePath);
        string directoryPath = Path.GetDirectoryName(fullPath);
        string fileName = NormalizeWorkbookOutputFileName(Path.GetFileName(fullPath));

        fullPath = string.IsNullOrWhiteSpace(directoryPath)
            ? fileName
            : Path.Combine(directoryPath, fileName);

        return fullPath;
    }

    private static string NormalizeWorkbookOutputFileName(string fileName)
    {
        string normalizedFileName = Path.GetFileName(fileName ?? string.Empty);
        if (string.IsNullOrWhiteSpace(normalizedFileName))
        {
            normalizedFileName = "output";
        }

        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(normalizedFileName);
        if (string.IsNullOrWhiteSpace(fileNameWithoutExtension))
        {
            fileNameWithoutExtension = "output";
        }

        char[] invalidChars = Path.GetInvalidFileNameChars();
        var builder = new StringBuilder(fileNameWithoutExtension.Length);
        foreach (char c in fileNameWithoutExtension)
        {
            builder.Append(invalidChars.Contains(c) ? '_' : c);
        }

        string safeFileNameWithoutExtension = builder.ToString().Trim();
        if (string.IsNullOrWhiteSpace(safeFileNameWithoutExtension))
        {
            safeFileNameWithoutExtension = "output";
        }

        return safeFileNameWithoutExtension + Path.GetExtension(templateFileName);
    }

    private static string GetOutputWorkbookFileNameFromJson(string jsonFilePath)
    {
        string jsonString = File.ReadAllText(jsonFilePath);
        JsonNode jsonObject = JsonNode.Parse(jsonString);
        var confData = GetPropertiesFromJsonNode(jsonObject, "variables");

        const string outputFilenameConfName = "outputFilename";
        string rawFileName = confData.ContainsKey(outputFilenameConfName)
            ? confData[outputFilenameConfName]
            : Path.GetFileNameWithoutExtension(jsonFilePath);
        return NormalizeWorkbookOutputFileName(rawFileName);
    }

    private static async Task<string> TryGetProjectFolderNameAsync(
        string baseUrl,
        string projectId,
        string token)
    {
        try
        {
            GitLabProjectInfo projectInfo = await GitLabClient.GetProjectInfoAsync(
                baseUrl,
                projectId,
                token).ConfigureAwait(false);

            return projectInfo == null ? null : projectInfo.Name;
        }
        catch (Exception ex)
        {
            FileLogger.Warn("Failed to resolve GitLab project name. " + ex.Message);
            return null;
        }
    }

    private static void SavePullStateToWorkbook(
        Excel.Workbook workbook,
        GitLabLastInput pullInfo,
        string pullCommitId,
        GitLabShareInfo shareInfo = null)
    {
        if (workbook == null)
        {
            return;
        }

        if (pullInfo != null)
        {
            workbook.SetCustomProperty(gitLabPullInfoCustomPropertyName, pullInfo);
        }

        if (!string.IsNullOrWhiteSpace(pullCommitId))
        {
            workbook.SetCustomProperty(gitLabPullCommitIdCustomPropertyName, pullCommitId);
        }

        if (shareInfo != null)
        {
            workbook.SetCustomProperty(gitLabShareInfoCustomPropertyName, shareInfo);
        }
    }

    private static bool ConfirmSaveBeforeLatestFetch(Excel.Workbook workbook)
    {
        if (workbook == null || workbook.Saved)
        {
            return true;
        }

        DialogResult yesNoCancel = MessageBox.Show(
            "最新版取得の前にファイルの変更内容を保存しますか？",
            "確認",
            MessageBoxButtons.YesNoCancel,
            MessageBoxIcon.Exclamation);

        switch (yesNoCancel)
        {
            case DialogResult.Yes:
                workbook.Save();
                return true;
            case DialogResult.No:
                return true;
            default:
                return false;
        }
    }

    async Task<bool> CreateNewWorkbook(
        string txtFilePath = null,
        string jsonFilePath = null,
        string newFilePathOverride = null,
        bool failIfExists = false,
        bool confirmOverwriteIfExists = false,
        GitLabLastInput pullInfo = null,
        string pullCommitId = null,
        GitLabShareInfo shareInfo = null)
    {
        // jsonFilePath を読み、confData 等を取得してテンプレから生成
        string jsonString = File.ReadAllText(jsonFilePath);
        JsonNode jsonObject = JsonNode.Parse(jsonString);
        var confData = GetPropertiesFromJsonNode(jsonObject, "variables");
        string newFileName = GetOutputWorkbookFileNameFromJson(jsonFilePath);

        if (IsSameNameWorkbookOpen(newFileName))
        {
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(newFileName);
            MessageBox.Show($"'{fileNameWithoutExtension}'と同じ名前のファイルが既に開かれています。\nファイルを閉じてから再度実行してください。");
            return false;
        }

        string jsonFileDirectory = Path.GetDirectoryName(jsonFilePath);
        string newFilePath = string.IsNullOrWhiteSpace(newFilePathOverride)
            ? Path.Combine(jsonFileDirectory, newFileName)
            : GetNormalizedWorkbookOutputPath(newFilePathOverride);

        if (File.Exists(newFilePath))
        {
            if (failIfExists)
            {
                MessageBox.Show(
                    "同名のファイルが既に存在するため新規作成できません。\n" +
                    "更新したい場合はそのファイルを開いてから Pull を実行してください。\n\n" +
                    newFilePath,
                    "Pull 新規作成");
                return false;
            }

            if (confirmOverwriteIfExists)
            {
                DialogResult overwriteResult = MessageBox.Show(
                    "同名のファイルが既に存在します。上書きしますか？\n\n" + newFilePath,
                    "Pull 新規作成",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (overwriteResult != DialogResult.Yes)
                {
                    return false;
                }
            }
        }

        EnsureDirectoryForLocalPath(newFilePath);

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
            return false;
        }
        if (templateSheetName == null)
        {
            // ブックを保存せずに閉じる
            workbook.Close(false);
            MessageBox.Show($"{templateFileName} のカスタムプロパティに {templateSheetNameCustomPropertyName} が設定されていません。");
            return false;
        }

        var missingImagePaths = await RenderWorkbook(workbook, jsonObject, confData, jsonFilePath, null);

        if (missingImagePaths.Any())
        {
            ShowMissingImageFilesDialog(missingImagePaths);
        }

        // RenderLog は TXT を保存
        RenderLog renderLog = new RenderLog
        {
            SourceFilePath = txtFilePath,   // TXT を保存
            User = Environment.UserName
        };
        workbook.SetCustomProperty("RenderLog", renderLog);

        //var lastRenderLog = workbook.GetCustomProperty<RenderLog>("RenderLog");

        string projectId = confData["project"];
        workbook.SetCustomProperty(ssProjectIdCustomPropertyName, projectId);

        SavePullStateToWorkbook(workbook, pullInfo, pullCommitId, shareInfo);

        workbook.Save();
        return true;
    }

    public async void OnCreateNewButtonPressed(IRibbonControl control)
    {
        Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
        string txtFilePath = OpenSourceFile();
        if (txtFilePath == null)
        {
            return;
        }

        FileLogger.InitializeForInput(txtFilePath, timestamped: false);

        bool parseSucceeded = RunParsePipeline(txtFilePath, true);
        if (!parseSucceeded)
        {
            return;
        }

        string jsonFilePath = TxtToJsonPath(txtFilePath);

        await CreateNewWorkbook(txtFilePath, jsonFilePath);

        excelApp.EnableEvents = true;
        if (excelApp.ActiveWorkbook != null)
        {
            excelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
        }
        excelApp.ScreenUpdating = true;
        excelApp.DisplayAlerts = true;
        excelApp.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityByUI;
    }

    public async void OnRegenerateWorkbookPressed(IRibbonControl control)
    {
        Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
        Excel.Workbook workbook = excelApp.ActiveWorkbook as Excel.Workbook;

        if (workbook == null)
        {
            MessageBox.Show("アクティブなブックが見つかりません。");
            return;
        }

        WorkbookInfo workbookInfo = WorkbookInfo.CreateFromWorkbook(workbook);

        if (workbookInfo == null)
        {
            string projectName = Assembly.GetExecutingAssembly().GetName().Name;
            MessageBox.Show($"{projectName} で生成されたブックではありません。");
            return;
        }

        string txtFilePath = SelectSourceFileForRender(workbookInfo);
        if (txtFilePath == null)
        {
            return;
        }

        FileLogger.InitializeForInput(txtFilePath, timestamped: false);

        string jsonFilePath = TxtToJsonPath(txtFilePath);

        if (File.Exists(jsonFilePath))
        {
            try
            {
                File.Delete(jsonFilePath);
            }
            catch (Exception ex)
            {
                FileLogger.Error($"json の削除に失敗しました: {ex}");
                MessageBox.Show($"json の削除に失敗しました。\n{ex.Message}");
                return;
            }
        }

        bool parseSucceeded = RunParsePipeline(txtFilePath, true);
        if (!parseSucceeded)
        {
            return;
        }

        if (!File.Exists(jsonFilePath))
        {
            FileLogger.Warn($"jsonファイルが見つかりません: {jsonFilePath}");
            return;
        }

        await RegenerateWorkbook(workbook, workbookInfo, txtFilePath, jsonFilePath);
    }

    public async void OnRenderButtonPressed(IRibbonControl control)
    {
        Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
        var sheet = excelApp.ActiveSheet as Excel.Worksheet;
        Excel.Workbook workbook = sheet == null ? null : sheet.Parent as Excel.Workbook;
        string txtFilePath = null;
        string jsonFilePath = null;
        bool isNewWorkbook = sheet == null;

        try
        {
            excelApp.StatusBar = "Parsing...";

            if (isNewWorkbook)
            {
                DialogResult fileSelectionResult = MessageBox.Show(
                    "ファイルを新規作成します。\nソースの TXT を選択してください。",
                    "確認", MessageBoxButtons.OKCancel);
                if (fileSelectionResult != DialogResult.OK)
                {
                    return;
                }

                txtFilePath = OpenSourceFile();
                if (txtFilePath == null)
                {
                    return;
                }
            }
            else
            {
                WorkbookInfo workbookInfo = WorkbookInfo.CreateFromWorkbook(workbook);

                if (workbookInfo == null)
                {
                    string projectName = Assembly.GetExecutingAssembly().GetName().Name;
                    MessageBox.Show($"{projectName} で生成されたブックではありません。");
                    return;
                }

                txtFilePath = SelectSourceFileForRender(workbookInfo);
                if (txtFilePath == null)
                {
                    return;
                }
            }

            FileLogger.InitializeForInput(txtFilePath, timestamped: false);

            bool parseSucceeded = RunParsePipeline(txtFilePath, true);
            if (!parseSucceeded)
            {
                return;
            }

            jsonFilePath = TxtToJsonPath(txtFilePath);
            if (!File.Exists(jsonFilePath))
            {
                FileLogger.Warn($"jsonファイルが見つかりません: {jsonFilePath}");
                return;
            }

            excelApp.StatusBar = "Rendering...";

            Stopwatch renderStopwatch = Stopwatch.StartNew();
            FileLogger.Info("Render started.");

            if (isNewWorkbook)
            {
                await CreateNewWorkbook(txtFilePath, jsonFilePath);
            }
            else
            {
                await UpdateAllSheets(workbook, txtFilePath, jsonFilePath);
            }

            renderStopwatch.Stop();
            FileLogger.Info($"Render finished ({renderStopwatch.ElapsedMilliseconds} ms).");
        }
        finally
        {
            excelApp.StatusBar = false;
            excelApp.EnableEvents = true;
            if (excelApp.ActiveWorkbook != null)
            {
                excelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            }
            excelApp.ScreenUpdating = true;
            excelApp.DisplayAlerts = true;
            excelApp.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityByUI;
        }
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

    static string[,] ConvertTo2DArray_GroupAligned(
        List<List<NodeData>> rows,
        out int alignedDepth,
        out List<Dictionary<int,int>> indexMap /* j -> aligned column c per row */)
    {
        // 1) 各 group の幅（= max(depthInGroup)+1）を集計
        var groupWidth = new Dictionary<int, int>();
        foreach (var path in rows)
        {
            foreach (var node in path)
            {
                if (node == null) continue;

                int g = node.group;
                int d = node.depthInGroup;
                if (d < 0) d = 0;

                int w = d + 1;
                if (!groupWidth.TryGetValue(g, out int cur) || w > cur)
                {
                    groupWidth[g] = w;
                }
            }
        }

        // 2) group を昇順に並べ、各 group の先頭オフセット（base）を決める
        List<int> groups = groupWidth.Keys.OrderBy(x => x).ToList();
        if (groups.Count == 0)
        {
            alignedDepth = 0;
            indexMap = new List<Dictionary<int,int>>();
            return new string[rows.Count, 0];
        }

        var groupBase = new Dictionary<int, int>();
        int offset = 0;
        foreach (int g in groups)
        {
            groupBase[g] = offset;
            offset += groupWidth[g];
        }
        alignedDepth = offset;

        // 3) 配列本体と j->c の対応表（行ごと）を用意
        int rcount = rows.Count;
        string[,] array = new string[rcount, alignedDepth];
        indexMap = Enumerable.Range(0, rcount)
                            .Select(_ => new Dictionary<int,int>())
                            .ToList();

        // 4) 配置: c = groupBase[g] + depthInGroup
        for (int r = 0; r < rcount; r++)
        {
            var path = rows[r];

            for (int j = 0; j < path.Count; j++)
            {
                var node = path[j];
                if (node == null) continue;

                int g = node.group;
                int d = node.depthInGroup;
                if (d < 0) d = 0;

                if (!groupBase.TryGetValue(g, out int b))
                {
                    continue;
                }

                int c = b + d;
                if (0 <= c && c < alignedDepth)
                {
                    array[r, c] = node.text;
                    indexMap[r][j] = c; // j(元の列) -> c(整列後の実列)
                }
            }
        }

        return array;
    }

    // 各行の「連続する末尾の空欄」を、最右 1 セルだけ leader にし、それ以外は空欄のままにする
    static void ReplaceTailWithSingleLeader(string[,] array, string leader)
    {
        int numRows = array.GetLength(0);
        int numCols = array.GetLength(1);
        if (numCols == 0) return;
        for (int i = 0; i < numRows; i++)
        {
            // 行末から最初の非 null を探す
            int lastNonNull = -1;
            for (int j = numCols - 1; j >= 0; j--)
            {
                if (array[i, j] != null)
                {
                    lastNonNull = j;
                    break;
                }
            }
            if (lastNonNull == -1)
            {
                // 行全体が空なら何もしない（記号も置かない）
                continue;
            }
            int lastCol = numCols - 1;
            // 末尾領域（lastNonNull+1 .. lastCol-1）は空欄(null)のままに統一
            for (int j = lastNonNull + 1; j < lastCol; j++)
            {
                array[i, j] = null;
            }
            // 最右 1 セルだけリーダー記号
            if (lastNonNull < lastCol)
            {
                array[i, lastCol] = leader;
            }
        }
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

        // group列揃えで2次元配列に変換
        int alignedDepth;
        List<Dictionary<int,int>> indexMap;
        string[,] arrayResult = ConvertTo2DArray_GroupAligned(result, out alignedDepth, out indexMap);
        maxDepth = alignedDepth;

        // 右端の“リーダー記号”のみ付与（途中は空欄）
        string leaderSymbol = (confData != null && confData.ContainsKey("LEADER_SYMBOL") && !string.IsNullOrEmpty(confData["LEADER_SYMBOL"]))
            ? confData["LEADER_SYMBOL"]
            : " ";
        ReplaceTailWithSingleLeader(arrayResult, leaderSymbol);

        int startColumn = 3;
        const int endColumn = 6;
        int columnWidth = endColumn - startColumn + 1;
        const int startRow = 25;
        const int endRow = 102;
        int rowHeight = endRow - startRow + 1;
        // ヘッダー行（本文開始行の 1 行上）
        const int headerRow = startRow - 1;

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

        // 深度が列より少ない場合は、余った列を列全体削除で詰める
        // 右端の実データ列の書式を維持するため、余剰列は左から削除する
        if (maxDepth < columnWidth)
        {
            int deleteCount = columnWidth - maxDepth;

            Excel.Range deleteRange = dstSheet.Range[
                dstSheet.Columns[startColumn],
                dstSheet.Columns[startColumn + deleteCount - 1]
            ];

            deleteRange.Delete();

            columnWidth = maxDepth;
        }

        // ─────────────────────────────────────────────────────────────
        // ヘッダー行の描画（tableHeadersNonInputArea）
        //  - JSON 上は常に存在するが、空配列のこともある
        //  - パディングは空セル（= 名前はグループの先頭列にのみ置く）
        //  - グループの列範囲は本文と同じ group 揃え規則で再構成する
        // ─────────────────────────────────────────────────────────────
        JsonArray thnia = sheetNode["tableHeadersNonInputArea"] as JsonArray;
        if (thnia != null && thnia.Count > 0)
        {
            // group -> width (= max(depthInGroup)+1)
            var groupWidthForHeader = new Dictionary<int, int>();
            foreach (var path in result)
            {
                foreach (var n in path)
                {
                    if (n == null) continue;
                    int g = n.group;
                    int d = n.depthInGroup;
                    if (d < 0) d = 0;
                    int w = d + 1;
                    if (!groupWidthForHeader.TryGetValue(g, out int cur) || w > cur)
                    {
                        groupWidthForHeader[g] = w;
                    }
                }
            }
            // group の並びと base オフセット
            var groupsForHeader = groupWidthForHeader.Keys.OrderBy(x => x).ToList();
            var groupBaseForHeader = new Dictionary<int, int>();
            int off = 0;
            foreach (var g in groupsForHeader)
            {
                groupBaseForHeader[g] = off;
                off += groupWidthForHeader[g];
            }
            // 1 行分のヘッダー配列（null 埋め = パディング空欄）
            string[,] headerArray = new string[1, maxDepth];
            foreach (JsonNode h in thnia)
            {
                // { "group": <int>, "name": <string>, "size": <int> } が想定
                if (h == null) continue;
                int g = 0;
                try { g = h["group"].GetValue<int>(); } catch { continue; }
                string name = null;
                try { name = h["name"]?.GetValue<string>(); } catch { }
                if (string.IsNullOrEmpty(name)) continue;

                if (!groupBaseForHeader.TryGetValue(g, out int b)) continue;
                if (0 <= b && b < maxDepth)
                {
                    headerArray[0, b] = name; // 先頭列のみに名前、右側は空欄（パディングしない）
                }
            }
            // ヘッダーを書き込み（本文は startRow から、ヘッダーは 1 行上）
            dstSheet.SetValueInSheet(headerRow, startColumn, headerArray);
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

                // 整列後の列位置（j -> col）を解決。無ければこの j はスキップ。
                if (!indexMap[i].TryGetValue(j, out int col)) { continue; }

                void ApplyCommentCell(int cellColorIndex, int fontColorIndex)
                {
                    // ここより右のセルの色を変える
                    var cells = dstSheet.Cells[startRow + i, startColumn + col];
                    cells = cells.Resize(1, maxDepth - col);

                    cells.Interior.ColorIndex = cellColorIndex;
                    cells.Font.ColorIndex = fontColorIndex;

                    // チェック予定日欄を空欄にする
                    var dateCell = dstSheet.Cells[startRow + i, dateColumn];
                    dateCell.Value = null;
                }

                void ApplyCommentCellColor(Color cellColor, Color fontColor)
                {
                    // ここより右のセルの色を変える
                    var cells = dstSheet.Cells[startRow + i, startColumn + col];
                    cells = cells.Resize(1, maxDepth - col);
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
                var tagColors = new Dictionary<string, (Color cellColor, Color fontColor, string emoji)>
                {
                    { "NOTE", (ColorTranslator.FromHtml("#cce5ff"), ColorTranslator.FromHtml("#004085"), "ℹ️") },    // 📝
                    { "TIP", (ColorTranslator.FromHtml("#d4edda"), ColorTranslator.FromHtml("#155724"), "💡") },
                    //{ "[!IMPORTANT]", (ColorTranslator.FromHtml("#d1ecf1"), ColorTranslator.FromHtml("#0c5460")) },
                    { "IMPORTANT", (ColorTranslator.FromHtml("#e2dbff"), ColorTranslator.FromHtml("#5936bb"), "📌") },
                    { "WARNING", (ColorTranslator.FromHtml("#fff3cd"), ColorTranslator.FromHtml("#856404"), "⚠️") },
                    { "CAUTION", (ColorTranslator.FromHtml("#f8d7da"), ColorTranslator.FromHtml("#721c24"), "⛔") },  // 🚨🚫
                };

                bool applied = false;

                var match = Regex.Match(text, @"^\[\!(\w+)(-)?\]\s*(.*)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                if (match.Success)
                {
                    var tagName = match.Groups[1].Value.ToUpper(); // 例: "NOTE"
                    var noEmoji = match.Groups[2].Success;         // "-" があれば true
                    var body = match.Groups[3].Value;              // 本文

                    if (tagColors.TryGetValue(tagName, out var style))
                    {
                        if (!noEmoji)
                        {
                            body = style.emoji + " " + body;
                        }

                        ApplyCommentCellColor(style.cellColor, style.fontColor);
                        dstSheet.Cells[startRow + i, startColumn + col].Value = body;
                        dstSheet.Cells[startRow + i, resultColumn].Value = "-";
                        applied = true;
                    }
                }

                if (applied)
                {
                    //break;
                    continue;
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

                // 整列後の列位置（j -> col2）。無ければスキップ。
                if (!indexMap[i].TryGetValue(j, out int col2)) { continue; }

                if (node.imageFilePath != null)
                {
                    string path = GetAbsolutePathFromBasePath(jsonFilePath, node.imageFilePath);
                    var cell = dstSheet.Cells[startRow + i, startColumn + col2];

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
        if (File.Exists(newFilePath))
        {
            File.Delete(newFilePath);
        }

        File.Copy(filePath, newFilePath, false);

        FileAttributes copiedFileAttributes = File.GetAttributes(newFilePath);
        if ((copiedFileAttributes & FileAttributes.ReadOnly) != 0)
        {
            File.SetAttributes(newFilePath, copiedFileAttributes & ~FileAttributes.ReadOnly);
        }

        Excel.Workbook workbook = excelApp.Workbooks.Open(newFilePath, ReadOnly: false);

        // Excelを表示
        excelApp.Visible = true;

        return workbook;
    }

    string SelectSourceFileForParse(bool forceSelectNew)
    {
        if (!forceSelectNew)
        {
            Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
            var workbook = excelApp.ActiveWorkbook as Excel.Workbook;
            string storedSourceFilePath = GetStoredSourceFilePathFromWorkbook(workbook);
            if (storedSourceFilePath != null)
            {
                return storedSourceFilePath;
            }
        }

        return OpenSourceFile();
    }

    void ApplyPostProcess(string jsonPath)
    {
        TimeAssigner.Assign(jsonPath);
    }

    bool RunParsePipeline(string txtFilePath, bool runPostProcess)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        FileLogger.Info("Parse started.");

        try
        {
            currentRootDirectory = Path.GetDirectoryName(Path.GetFullPath(txtFilePath));
            currentGitLabBaseFileRelativePath = ResolveInitialGitLabBaseFileRelativePath(txtFilePath);
            JsHost.SetFilePathResolveHook((requestedPath, baseFilePath) => ResolveAndEnsureLocalFilePathForJs(requestedPath, baseFilePath));
            JsHost.SetFileReadTraceHook(message => AddFileReadTrace(message));
            var result = JsHost.Call("parse", txtFilePath);
            if (IsQuitResult(result))
            {
                return false;
            }
        }
        catch (Microsoft.ClearScript.ScriptEngineException ex)
        {
            var iex = ex.InnerException;
            while (iex != null)
            {
                if (iex is JsQuitException q)
                {
                    FileLogger.Error(ex.ToString());
                    Notifier.Error("エラー", "中断されましました。クリックでログを開きます。");

                    MessageBox.Show(q.Message, "中断");
                }
                iex = iex.InnerException;
            }

            string details = ex.ErrorDetails;

            FileLogger.Error(ex.ToString());
            Notifier.Error("エラー", "パースでエラーが発生しました。クリックでログを開きます。");

            MessageBox.Show(details, "JS実行エラー");
            return false;
        }
        finally
        {
            JsHost.ClearFileReadTraceHook();
            JsHost.ClearFilePathResolveHook();
            currentGitLabBaseFileRelativePath = null;
            currentRootDirectory = null;
        }

        string jsonPath = Path.ChangeExtension(txtFilePath, ".json");
        if (runPostProcess)
        {
            if (!File.Exists(jsonPath))
            {
                FileLogger.Warn($"jsonファイルが見つかりません: {jsonPath}");
                return false;
            }

            try
            {
                ApplyPostProcess(jsonPath);
            }
            catch (Exception ex)
            {
                FileLogger.Error(ex.ToString());
                Notifier.Error("エラー", "後処理でエラーが発生しました。クリックでログを開きます。");
                MessageBox.Show(ex.Message, "後処理エラー");
                return false;
            }
        }

        stopwatch.Stop();
        FileLogger.Info($"Parse finished ({stopwatch.ElapsedMilliseconds} ms).");

        if (!runPostProcess)
        {
            Notifier.Info("正常終了", $"jsonファイルを出力しました\n{jsonPath}");
        }
        return true;
    }

    private static string ResolveAndEnsureLocalFilePathForJs(string requestedPath, string baseFilePath)
    {
        if (string.IsNullOrWhiteSpace(requestedPath))
        {
            throw new ArgumentException("requestedPath is empty.", "requestedPath");
        }

        if (currentPullSession == null ||
            string.IsNullOrEmpty(currentPullSession.WorkRoot) ||
            string.IsNullOrEmpty(currentPullSession.EntryGitLabRelativePath))
        {
            return Path.GetFullPath(requestedPath);
        }

        string workRoot = Path.GetFullPath(currentPullSession.WorkRoot)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        string baseFileRelativePath = ResolveBaseFileRelativePath(baseFilePath);
        string resolvedGitLabRelativePath;

        if (Path.IsPathRooted(requestedPath))
        {
            string fullPath = Path.GetFullPath(requestedPath);
            if (!fullPath.StartsWith(workRoot, StringComparison.OrdinalIgnoreCase))
            {
                return fullPath;
            }

            resolvedGitLabRelativePath = ToGitLabRelativePath(workRoot, fullPath);
        }
        else
        {
            resolvedGitLabRelativePath = GitLabPathResolver.ResolveGitLabRelativePath(baseFileRelativePath, requestedPath);
        }

        AddFileReadTrace("[resolved] gitlabRelative=" + resolvedGitLabRelativePath);

        string localEnsuredPath = EnsureFileInWorkRootAsync(
            currentPullSession.BaseUrl,
            currentPullSession.ProjectId,
            currentPullSession.RefName,
            currentPullSession.Token,
            currentPullSession.WorkRoot,
            resolvedGitLabRelativePath,
            requestedPath,
            currentPullSession.SessionLog)
            .GetAwaiter()
            .GetResult();

        return localEnsuredPath;
    }

    internal static string ResolveAndEnsureLocalPathForFileSystemCheck(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return path;
        }

        try
        {
            return ResolveAndEnsureLocalFilePathForJs(path, path);
        }
        catch
        {
            return path;
        }
    }

    internal static string NormalizeLazyReadPathIdentity(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        string fullPath = Path.GetFullPath(path);

        if (currentPullSession == null ||
            string.IsNullOrEmpty(currentPullSession.WorkRoot) ||
            string.IsNullOrEmpty(currentPullSession.EntryGitLabRelativePath))
        {
            return fullPath.Replace('\\', '/');
        }

        string normalizedWorkRoot = Path.GetFullPath(currentPullSession.WorkRoot)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        string rootWithSeparator = normalizedWorkRoot + Path.DirectorySeparatorChar;

        if (!string.Equals(fullPath, normalizedWorkRoot, StringComparison.OrdinalIgnoreCase) &&
            !fullPath.StartsWith(rootWithSeparator, StringComparison.OrdinalIgnoreCase))
        {
            return fullPath.Replace('\\', '/');
        }

        return ToGitLabRelativePath(normalizedWorkRoot, fullPath);
    }

    private static void AddFileReadTrace(string message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return;
        }

        FileLogger.Info("[PullLazyReadTrace] " + message);

        if (currentPullSession != null && currentPullSession.SessionLog != null)
        {
            currentPullSession.SessionLog.Add(PullFileActionType.FileReadTrace, message);
        }
    }

    private static void ReportPullProgress(string message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return;
        }

        PullSessionContext sessionContext = currentPullSession;
        if (sessionContext == null || sessionContext.ProgressReporter == null)
        {
            return;
        }

        sessionContext.ProgressReporter(message);
    }

    private static string ResolveBaseFileRelativePath(string baseFilePath)
    {
        if (!string.IsNullOrWhiteSpace(baseFilePath))
        {
            if (Path.IsPathRooted(baseFilePath))
            {
                string fullBasePath = Path.GetFullPath(baseFilePath);

                if (HasActivePullSession())
                {
                    string workRoot = currentPullSession.WorkRoot;

                    if (IsPathInRoot(fullBasePath, workRoot))
                    {
                        return ToGitLabRelativePath(NormalizeRootPath(workRoot), fullBasePath);
                    }

                    throw new InvalidOperationException(
                        "baseFilePath is outside WorkRoot in pull mode. " +
                        "baseFilePath='" + baseFilePath + "', " +
                        "workRoot='" + workRoot + "'.");
                }

                string rootDirectoryOutsidePull = GetBaseFileRootDirectory(false, baseFilePath);
                if (!IsPathInRoot(fullBasePath, rootDirectoryOutsidePull))
                {
                    throw new InvalidOperationException(
                        "baseFilePath is outside rootDirectory. " +
                        "baseFilePath='" + baseFilePath + "', " +
                        "rootDirectory='" + rootDirectoryOutsidePull + "'.");
                }

                return ToGitLabRelativePath(NormalizeRootPath(rootDirectoryOutsidePull), fullBasePath);
            }

            return GitLabPathResolver.CanonicalizeGitLabRelativePath(baseFilePath, "baseFilePath");
        }

        if (!string.IsNullOrEmpty(currentGitLabBaseFileRelativePath))
        {
            return currentGitLabBaseFileRelativePath;
        }

        if (HasActivePullSession())
        {
            return currentPullSession.EntryGitLabRelativePath;
        }

        return string.Empty;
    }

    private static bool HasActivePullSession()
    {
        return currentPullSession != null &&
            !string.IsNullOrEmpty(currentPullSession.WorkRoot) &&
            !string.IsNullOrEmpty(currentPullSession.EntryGitLabRelativePath);
    }

    private static string NormalizeRootPath(string rootPath)
    {
        return Path.GetFullPath(rootPath)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }

    private static bool IsPathInRoot(string fullPath, string rootPath)
    {
        if (string.IsNullOrWhiteSpace(fullPath) || string.IsNullOrWhiteSpace(rootPath))
        {
            return false;
        }

        string normalizedRootPath = NormalizeRootPath(rootPath);
        string rootWithSeparator = normalizedRootPath + Path.DirectorySeparatorChar;

        return string.Equals(fullPath, normalizedRootPath, StringComparison.OrdinalIgnoreCase) ||
            fullPath.StartsWith(rootWithSeparator, StringComparison.OrdinalIgnoreCase);
    }

    private static string GetBaseFileRootDirectory(bool isPullMode, string baseFilePath)
    {
        if (isPullMode)
        {
            return currentPullSession.WorkRoot;
        }

        if (!string.IsNullOrWhiteSpace(currentRootDirectory))
        {
            return currentRootDirectory;
        }

        return Path.GetDirectoryName(Path.GetFullPath(baseFilePath));
    }

    private static string ResolveInitialGitLabBaseFileRelativePath(string txtFilePath)
    {
        if (currentPullSession == null ||
            string.IsNullOrEmpty(currentPullSession.WorkRoot) ||
            string.IsNullOrEmpty(currentPullSession.EntryGitLabRelativePath))
        {
            return null;
        }

        if (string.IsNullOrWhiteSpace(txtFilePath))
        {
            return currentPullSession.EntryGitLabRelativePath;
        }

        string fullTxtPath = Path.GetFullPath(txtFilePath);
        string normalizedWorkRoot = Path.GetFullPath(currentPullSession.WorkRoot)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        string rootWithSeparator = normalizedWorkRoot + Path.DirectorySeparatorChar;

        if (!string.Equals(fullTxtPath, normalizedWorkRoot, StringComparison.OrdinalIgnoreCase) &&
            !fullTxtPath.StartsWith(rootWithSeparator, StringComparison.OrdinalIgnoreCase))
        {
            return currentPullSession.EntryGitLabRelativePath;
        }

        return ToGitLabRelativePath(normalizedWorkRoot, fullTxtPath);
    }

    public void OnDebugParseButtonPressed(IRibbonControl control)
    {
        Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
        string txtPath2 = OpenSourceFile();
        if (txtPath2 == null)
        {
            return;
        }
        FileLogger.InitializeForInput(txtPath2, timestamped: false);
        RunParsePipeline(txtPath2, false);
    }

    public void OnDebugValidateNestedLazyReadButtonPressed(IRibbonControl control)
    {
        Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
        PullSessionContext debugSession = lastSuccessfulPullSession;
        if (debugSession == null ||
            string.IsNullOrEmpty(debugSession.WorkRoot) ||
            string.IsNullOrEmpty(debugSession.EntryGitLabRelativePath))
        {
            MessageBox.Show("先に Pull を実行してください。", "NestedRead");
            return;
        }

        string entryGitLabRelativePath = debugSession.EntryGitLabRelativePath;
        string entryLocalPath = BuildLocalPathInWorkRoot(debugSession.WorkRoot, entryGitLabRelativePath);

        FileLogger.InitializeForSession(debugSession.WorkRoot, "nested-lazy-read", timestamped: false);
        AddFileReadTrace("[validate-setup] entry=" + entryGitLabRelativePath);

        try
        {
            currentGitLabBaseFileRelativePath = ResolveInitialGitLabBaseFileRelativePath(entryLocalPath);
            JsHost.SetFilePathResolveHook((requestedPath, baseFilePath) => ResolveAndEnsureLocalFilePathForJs(requestedPath, baseFilePath));
            JsHost.SetFileReadTraceHook(message => AddFileReadTrace(message));

            JsHost.Call(
                "debugValidateNestedLazyReadChain",
                entryLocalPath,
                "../common/a.txt",
                "./b.txt");

            MessageBox.Show("Nested lazy-read validation completed. ログを確認してください。", "NestedRead");
        }
        catch (Exception ex)
        {
            FileLogger.Error(ex.ToString());
            MessageBox.Show(ex.ToString(), "NestedRead failed");
        }
        finally
        {
            JsHost.ClearFileReadTraceHook();
            JsHost.ClearFilePathResolveHook();
            currentGitLabBaseFileRelativePath = null;
        }
    }

    static string SelectInputFileForRenderOnly()
    {
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            openFileDialog.Filter = "JSON ファイル (*.json)|*.json";
            openFileDialog.Title = "レンダーに使用するファイルを選択してください";

            return (openFileDialog.ShowDialog() == DialogResult.OK)
                ? openFileDialog.FileName
                : null;
        }
    }

    public async void OnRenderOnlyDebugButtonPressed(IRibbonControl control)
    {
        Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
        string selectedPath = SelectInputFileForRenderOnly();
        if (selectedPath == null)
        {
            return;
        }

        string txtFilePath = selectedPath;
        string jsonFilePath = selectedPath;

        if (string.Equals(Path.GetExtension(selectedPath), ".json", StringComparison.OrdinalIgnoreCase))
        {
            txtFilePath = NormalizeSourceFilePath(selectedPath);
        }
        else
        {
            jsonFilePath = TxtToJsonPath(txtFilePath);
        }

        var sheet = excelApp.ActiveSheet as Excel.Worksheet;
        bool isNewWorkbook = sheet == null;
        Excel.Workbook workbook = isNewWorkbook ? null : sheet.Parent as Excel.Workbook;

        if (!isNewWorkbook)
        {
            WorkbookInfo workbookInfo = WorkbookInfo.CreateFromWorkbook(workbook);
            if (workbookInfo == null)
            {
                string projectName = Assembly.GetExecutingAssembly().GetName().Name;
                MessageBox.Show($"{projectName} で生成されたブックではありません。");
                return;
            }
        }

        try
        {
            excelApp.StatusBar = "Rendering...";
            FileLogger.InitializeForInput(txtFilePath, timestamped: false);

            Stopwatch renderStopwatch = Stopwatch.StartNew();
            FileLogger.Info("Render started.");

            if (isNewWorkbook)
            {
                await CreateNewWorkbook(txtFilePath, jsonFilePath);
            }
            else
            {
                await UpdateAllSheets(workbook, txtFilePath, jsonFilePath);
            }

            renderStopwatch.Stop();
            FileLogger.Info($"Render finished ({renderStopwatch.ElapsedMilliseconds} ms).");
        }
        finally
        {
            excelApp.StatusBar = false;
        }
        return;

    }

    private static bool HasPullSourceSettings(GitLabLastInput input)
    {
        return input != null &&
               !string.IsNullOrWhiteSpace(input.BaseUrl) &&
               !string.IsNullOrWhiteSpace(input.ProjectId);
    }

    private static bool IsPullSourceEnabled(GitLabLastInput input)
    {
        return input != null && input.PullEnabled != false;
    }

    private static bool HasShareSettings(GitLabShareInfo shareInfo)
    {
        return shareInfo != null &&
               !string.IsNullOrWhiteSpace(shareInfo.BaseUrl) &&
               !string.IsNullOrWhiteSpace(shareInfo.ProjectId);
    }

    private static GitLabShareInfo CreateInitialShareSettings(GitLabShareInfo shareInfo, GitLabLastInput pullInfo)
    {
        if (shareInfo == null)
        {
            shareInfo = new GitLabShareInfo();
        }

        return new GitLabShareInfo
        {
            BaseUrl = !string.IsNullOrWhiteSpace(shareInfo.BaseUrl)
                ? shareInfo.BaseUrl
                : (pullInfo == null ? null : pullInfo.BaseUrl),
            ProjectId = shareInfo.ProjectId,
            RefName = !string.IsNullOrWhiteSpace(shareInfo.RefName)
                ? shareInfo.RefName
                : (pullInfo == null || string.IsNullOrWhiteSpace(pullInfo.RefName) ? "main" : pullInfo.RefName)
        };
    }

    private static GitLabLastInput GetInitialPullSettingsForDialog(Excel.Workbook workbook)
    {
        WorkbookInfo workbookInfo = workbook == null ? null : WorkbookInfo.CreateFromWorkbook(workbook);
        GitLabLastInput workbookInput = workbookInfo == null ? null : workbookInfo.PullInfo;
        GitLabLastInput storedInput = GitLabLastInputStore.Load() ?? new GitLabLastInput();

        if (workbookInput == null)
        {
            return new GitLabLastInput
            {
                BaseUrl = storedInput.BaseUrl,
                ProjectId = storedInput.ProjectId,
                RefName = storedInput.RefName,
                FilePath = storedInput.FilePath,
                PullEnabled = workbookInfo == null ? true : false
            };
        }

        return new GitLabLastInput
        {
            BaseUrl = !string.IsNullOrWhiteSpace(workbookInput.BaseUrl)
                ? workbookInput.BaseUrl
                : storedInput.BaseUrl,
            ProjectId = !string.IsNullOrWhiteSpace(workbookInput.ProjectId)
                ? workbookInput.ProjectId
                : storedInput.ProjectId,
            RefName = !string.IsNullOrWhiteSpace(workbookInput.RefName)
                ? workbookInput.RefName
                : storedInput.RefName,
            FilePath = !string.IsNullOrWhiteSpace(workbookInput.FilePath)
                ? workbookInput.FilePath
                : storedInput.FilePath,
            PullEnabled = workbookInput.PullEnabled ?? HasPullSourceSettings(workbookInput)
        };
    }

    private static GitLabLastInput CreateStoredPullDefaults(GitLabLastInput input, bool clearFilePath)
    {
        input = input ?? new GitLabLastInput();

        return new GitLabLastInput
        {
            BaseUrl = input.BaseUrl,
            ProjectId = input.ProjectId,
            RefName = input.RefName,
            FilePath = clearFilePath ? "" : input.FilePath,
            PullEnabled = null
        };
    }

    private static GitLabShareInfo GetInitialShareSettingsForDialog(Excel.Workbook workbook, GitLabLastInput pullInfo)
    {
        WorkbookInfo workbookInfo = workbook == null ? null : WorkbookInfo.CreateFromWorkbook(workbook);
        GitLabShareInfo shareInfo = workbookInfo == null ? null : workbookInfo.ShareInfo;
        if (shareInfo == null)
        {
            shareInfo = GitLabShareInfoStore.Load();
        }

        return CreateInitialShareSettings(shareInfo, pullInfo);
    }

    private void SavePullSourceSettings(GitLabLastInput input, Excel.Workbook workbook)
    {
        if (input == null)
        {
            return;
        }

        GitLabLastInputStore.Save(CreateStoredPullDefaults(input, false), false);

        WorkbookInfo workbookInfo = workbook == null ? null : WorkbookInfo.CreateFromWorkbook(workbook);
        if (workbookInfo == null)
        {
            return;
        }

        workbook.SetCustomProperty(gitLabPullInfoCustomPropertyName, input);
        workbook.SetCustomProperty(gitLabPullCommitIdCustomPropertyName, "");
    }

    private void SaveShareSettings(GitLabShareInfo shareInfo, Excel.Workbook workbook)
    {
        if (shareInfo == null)
        {
            return;
        }

        GitLabShareInfoStore.Save(shareInfo);

        WorkbookInfo workbookInfo = workbook == null ? null : WorkbookInfo.CreateFromWorkbook(workbook);
        if (workbookInfo == null)
        {
            return;
        }

        workbook.SetCustomProperty(gitLabShareInfoCustomPropertyName, shareInfo);
    }

    private bool TryEnsurePullSourceSettingsForWorkbook(
        Excel.Workbook workbook,
        string dialogTitle,
        out GitLabLastInput input)
    {
        input = GetInitialPullSettingsForDialog(workbook);
        if (HasPullSourceSettings(input) || !IsPullSourceEnabled(input))
        {
            return true;
        }

        DialogResult setupResult = MessageBox.Show(
            "このブックには取得元の情報が保存されていません。\n今すぐ設定しますか？",
            dialogTitle,
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);
        if (setupResult != DialogResult.Yes)
        {
            return false;
        }

        GitLabLastInput updated;
        if (!GitLabRepoDialog.TryShow(input, out updated))
        {
            return false;
        }

        SavePullSourceSettings(updated, workbook);
        input = updated;
        return true;
    }

    private bool TryEnsureShareSettingsForWorkbook(
        Excel.Workbook workbook,
        GitLabLastInput pullInfo,
        string dialogTitle,
        bool required,
        out GitLabShareInfo shareInfo)
    {
        WorkbookInfo workbookInfo = workbook == null ? null : WorkbookInfo.CreateFromWorkbook(workbook);
        shareInfo = workbookInfo == null ? null : workbookInfo.ShareInfo;

        if (HasShareSettings(shareInfo))
        {
            return true;
        }

        GitLabShareInfo initial = GetInitialShareSettingsForDialog(workbook, pullInfo);
        if (HasShareSettings(initial) && !required)
        {
            shareInfo = initial;
            return true;
        }

        if (required)
        {
            return false;
        }

        GitLabShareInfo updated;
        if (!GitLabShareSettingsDialog.TryShow(initial, pullInfo, out updated))
        {
            return false;
        }

        SaveShareSettings(updated, workbook);
        shareInfo = updated;
        return true;
    }

    private bool TryGetPullNewWorkbookInputs(
        Excel.Workbook workbook,
        out GitLabLastInput input,
        out GitLabShareInfo shareInfo)
    {
        input = GetInitialPullSettingsForDialog(workbook);
        shareInfo = GetInitialShareSettingsForDialog(workbook, input);

        if (!HasPullSourceSettings(input))
        {
            if (!GitLabRepoDialog.TryShow(input, out input))
            {
                shareInfo = null;
                return false;
            }

            SavePullSourceSettings(input, workbook);
        }

        if (!IsPullSourceEnabled(input) || !HasPullSourceSettings(input))
        {
            MessageBox.Show("新規作成では取得元を有効にしてください。", "新規作成");
            shareInfo = null;
            return false;
        }

        if (!HasShareSettings(shareInfo))
        {
            if (!GitLabShareSettingsDialog.TryShow(shareInfo, input, out shareInfo))
            {
                return false;
            }

            SaveShareSettings(shareInfo, workbook);
        }

        string filePath;
        if (!GitLabFilePathDialog.TryShow(input.FilePath, out filePath))
        {
            return false;
        }

        input.FilePath = filePath;
        GitLabLastInputStore.Save(CreateStoredPullDefaults(input, false), false);
        GitLabShareInfoStore.Save(shareInfo);
        return true;
    }

    public void OnPullSourceSettingsButtonPressed(IRibbonControl control)
    {
        Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
        Excel.Workbook workbook = excelApp.ActiveWorkbook as Excel.Workbook;
        GitLabLastInput initial = GetInitialPullSettingsForDialog(workbook);
        GitLabLastInput updated;

        if (!GitLabRepoDialog.TryShow(initial, out updated))
        {
            return;
        }

        SavePullSourceSettings(updated, workbook);
    }

    public void OnPullShareSettingsButtonPressed(IRibbonControl control)
    {
        Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
        Excel.Workbook workbook = excelApp.ActiveWorkbook as Excel.Workbook;
        GitLabLastInput pullInfo = GetInitialPullSettingsForDialog(workbook);
        GitLabShareInfo initial = GetInitialShareSettingsForDialog(workbook, pullInfo);
        GitLabShareInfo updated;

        if (!GitLabShareSettingsDialog.TryShow(initial, pullInfo, out updated))
        {
            return;
        }

        SaveShareSettings(updated, workbook);
    }

    public async void OnPullButtonPressed(IRibbonControl control)
    {
        Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
        var activeSheet = excelApp.ActiveSheet as Excel.Worksheet;
        Excel.Workbook activeWorkbook = activeSheet == null ? null : activeSheet.Parent as Excel.Workbook;
        PullProgressForm progressForm = new PullProgressForm();
        GitLabShareInfo shareInfo = null;

        if (SynchronizationContext.Current == null)
        {
            SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
        }

        try
        {
            ClearPullSessionState();

            GitLabLastInput input;
            bool isPullUpdate = activeWorkbook != null;

            if (isPullUpdate)
            {
                WorkbookInfo workbookInfo = WorkbookInfo.CreateFromWorkbook(activeWorkbook);
                if (workbookInfo == null)
                {
                    MessageBox.Show("このブックには Pull 用の情報が保存されていません。", "最新版取得");
                    return;
                }

                if (!TryEnsurePullSourceSettingsForWorkbook(activeWorkbook, "最新版取得", out input))
                {
                    return;
                }

                if (!TryEnsureShareSettingsForWorkbook(activeWorkbook, input, "最新版取得", required: false, out shareInfo))
                {
                    return;
                }

                bool shouldPersistPullSettings =
                    workbookInfo.PullInfo == null ||
                    workbookInfo.PullInfo.PullEnabled != input.PullEnabled ||
                    (!HasPullSourceSettings(workbookInfo.PullInfo) && (HasPullSourceSettings(input) || !IsPullSourceEnabled(input)));
                bool shouldPersistShareSettings =
                    shareInfo != null && !HasShareSettings(workbookInfo.ShareInfo);

                bool hasPullUpdate = false;
                string token = null;
                string currentCommitId = null;
                bool pullEnabled = IsPullSourceEnabled(input) && HasPullSourceSettings(input);

                if (pullEnabled)
                {
                    token = GitLabAuth.GetOrPromptToken(input.BaseUrl, input.ProjectId);
                    if (string.IsNullOrEmpty(token))
                    {
                        System.Windows.Forms.MessageBox.Show("同期をキャンセルしました（トークン未入力）");
                        return;
                    }

                    currentCommitId = await GitLabClient.GetCommitIdAsync(
                        input.BaseUrl,
                        input.ProjectId,
                        input.RefName,
                        token).ConfigureAwait(true);

                    hasPullUpdate =
                        string.IsNullOrWhiteSpace(workbookInfo.PullCommitId) ||
                        !string.Equals(workbookInfo.PullCommitId, currentCommitId, StringComparison.OrdinalIgnoreCase);
                }

                if (!hasPullUpdate)
                {
                    if (shareInfo == null)
                    {
                        if (shouldPersistPullSettings)
                        {
                            SavePullSourceSettings(input, activeWorkbook);
                        }
                        MessageBox.Show("最新版です。更新はありません。", "最新版取得");
                        return;
                    }

                    string shareToken = GitLabAuth.GetOrPromptToken(shareInfo.BaseUrl, shareInfo.ProjectId);
                    if (string.IsNullOrWhiteSpace(shareToken))
                    {
                        MessageBox.Show("共有値の取得をキャンセルしました（トークン未入力）", "最新版取得");
                        return;
                    }

                    bool hasSharedUpdates = await HasSharedUpdatesAsync(activeWorkbook, shareInfo, shareToken).ConfigureAwait(true);
                    if (!hasSharedUpdates)
                    {
                        if (shouldPersistPullSettings)
                        {
                            SavePullSourceSettings(input, activeWorkbook);
                        }

                        if (shouldPersistShareSettings)
                        {
                            SaveShareSettings(shareInfo, activeWorkbook);
                        }

                        MessageBox.Show("最新版です。更新はありません。", "最新版取得");
                        return;
                    }
                }

                if (!ConfirmSaveBeforeLatestFetch(activeWorkbook))
                {
                    return;
                }

                if (shouldPersistPullSettings)
                {
                    SavePullSourceSettings(input, activeWorkbook);
                }

                if (shouldPersistShareSettings)
                {
                    SaveShareSettings(shareInfo, activeWorkbook);
                }

                if (!hasPullUpdate)
                {
                    InitializeLoggerForSharedReceive(activeWorkbook);
                    progressForm.Show();
                    if (pullEnabled)
                    {
                        progressForm.AppendLine("Pull 元は最新版です");
                    }
                    else
                    {
                        progressForm.AppendLine("取得元更新は無効です");
                    }
                    progressForm.AppendLine("共有値を確認します");
                    SharedReceiveResult sharedReceiveResult = await ReceiveSharedSheetsAsync(activeWorkbook, shareInfo, progressForm.AppendLine);
                    progressForm.ShowContinueButton("閉じる", "共有値確認が完了しました");
                    await progressForm.WaitForContinueAsync();
                    progressForm.CloseForm();
                    progressForm = null;
                    ShowSharedReceiveConflictDialogIfNeeded(sharedReceiveResult);
                    return;
                }
            }
            else
            {
                if (!TryGetPullNewWorkbookInputs(null, out input, out shareInfo))
                {
                    return;
                }
            }

            progressForm.Show();
            progressForm.AppendLine("最新版の取得を開始します");

            PullExecutionResult pullResult = await ExecutePullAsync(input, progressForm.AppendLine);
            if (pullResult == null)
            {
                return;
            }

            if (isPullUpdate)
            {
                progressForm.ShowContinueButton("Excel 更新開始", "ダウンロード完了");
                await progressForm.WaitForContinueAsync();
                await UpdateAllSheets(activeWorkbook, pullResult.EntryLocalPath, pullResult.JsonFilePath);
                if (shareInfo != null)
                {
                    progressForm.AppendLine("共有値を反映しています");
                    SharedReceiveResult sharedReceiveResult = await ReceiveSharedSheetsAsync(activeWorkbook, shareInfo, progressForm.AppendLine);
                    ShowSharedReceiveConflictDialogIfNeeded(sharedReceiveResult);
                }
            }
            else
            {
                progressForm.ShowContinueButton("Excel 作成開始", "ダウンロード完了");
                await progressForm.WaitForContinueAsync();

                string projectFolderName = await TryGetProjectFolderNameAsync(input.BaseUrl, input.ProjectId, currentPullSession.Token);
                string outputDirectory = GetPullWorkbookOutputDirectory(input.ProjectId, projectFolderName);
                string outputFileName = GetOutputWorkbookFileNameFromJson(pullResult.JsonFilePath);
                string outputFilePath = Path.Combine(outputDirectory, outputFileName);

                var pullInfo = new GitLabLastInput
                {
                    BaseUrl = input.BaseUrl,
                    ProjectId = input.ProjectId,
                    RefName = input.RefName,
                    FilePath = pullResult.NormalizedEntryPath
                };

                bool workbookCreated = await CreateNewWorkbook(
                    pullResult.EntryLocalPath,
                    pullResult.JsonFilePath,
                    outputFilePath,
                    confirmOverwriteIfExists: true,
                    pullInfo: pullInfo,
                    pullCommitId: pullResult.RefCommitId,
                    shareInfo: shareInfo);
                if (!workbookCreated)
                {
                    return;
                }

                await ReceiveSharedSheetsAfterNewWorkbookCreatedAsync(shareInfo, progressForm);
            }

            string manifestPath = WritePullManifest(currentPullSession);
            FileLogger.Info("[PullManifest] written: " + manifestPath);

            SavePullStateToWorkbook(activeWorkbook, input, pullResult.RefCommitId, shareInfo);
            // MessageBox.Show(currentPullSession.SessionLog.BuildSummaryText(30), "Pull Result");
            lastSuccessfulPullSession = currentPullSession;
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show(ex.ToString(), "Pull failed");
        }
        finally
        {
            if (progressForm != null)
            {
                progressForm.CloseForm();
            }
            ClearPullSessionState();
        }
    }

    public async void OnPullCreateButtonPressed(IRibbonControl control)
    {
        Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
        if (SynchronizationContext.Current == null)
        {
            SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
        }

        PullProgressForm progressForm = new PullProgressForm();
        try
        {
            ClearPullSessionState();
            GitLabLastInput input;
            GitLabShareInfo shareInfo;

            if (!TryGetPullNewWorkbookInputs(null, out input, out shareInfo))
            {
                return;
            }

            progressForm.Show();
            progressForm.AppendLine("最新版の取得を開始します");

            PullExecutionResult pullResult = await ExecutePullAsync(input, progressForm.AppendLine);
            if (pullResult == null)
            {
                return;
            }

            progressForm.ShowContinueButton("Excel 作成開始", "ダウンロード完了");
            await progressForm.WaitForContinueAsync();

            string projectFolderName = await TryGetProjectFolderNameAsync(input.BaseUrl, input.ProjectId, currentPullSession.Token);
            string outputDirectory = GetPullWorkbookOutputDirectory(input.ProjectId, projectFolderName);
            string outputFileName = GetOutputWorkbookFileNameFromJson(pullResult.JsonFilePath);
            string outputFilePath = Path.Combine(outputDirectory, outputFileName);

            var pullInfo = new GitLabLastInput
            {
                BaseUrl = input.BaseUrl,
                ProjectId = input.ProjectId,
                RefName = input.RefName,
                FilePath = pullResult.NormalizedEntryPath
            };

            bool workbookCreated = await CreateNewWorkbook(
                pullResult.EntryLocalPath,
                pullResult.JsonFilePath,
                outputFilePath,
                confirmOverwriteIfExists: true,
                pullInfo: pullInfo,
                pullCommitId: pullResult.RefCommitId,
                shareInfo: shareInfo);
            if (!workbookCreated)
            {
                return;
            }

            await ReceiveSharedSheetsAfterNewWorkbookCreatedAsync(shareInfo, progressForm);

            string manifestPath = WritePullManifest(currentPullSession);
            FileLogger.Info("[PullManifest] written: " + manifestPath);

            // MessageBox.Show(currentPullSession.SessionLog.BuildSummaryText(30), "Pull Result");
            lastSuccessfulPullSession = currentPullSession;
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show(ex.ToString(), "Pull failed");
        }
        finally
        {
            progressForm.CloseForm();
            ClearPullSessionState();
        }
    }

    private async Task ReceiveSharedSheetsAfterNewWorkbookCreatedAsync(GitLabShareInfo shareInfo, PullProgressForm progressForm)
    {
        if (shareInfo == null)
        {
            return;
        }

        Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
        Excel.Workbook createdWorkbook = excelApp.ActiveWorkbook as Excel.Workbook;
        if (createdWorkbook == null)
        {
            return;
        }

        progressForm.SetStatusText("共有値を反映しています");
        progressForm.AppendLine("共有値を反映しています");
        SharedReceiveResult sharedReceiveResult = await ReceiveSharedSheetsAsync(createdWorkbook, shareInfo, progressForm.AppendLine);
        ShowSharedReceiveConflictDialogIfNeeded(sharedReceiveResult);
        createdWorkbook.Save();
    }

    public void OnShareCurrentSheetButtonPressed(IRibbonControl control)
    {
        Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
        Excel.Worksheet activeSheet = excelApp.ActiveSheet as Excel.Worksheet;
        if (activeSheet == null)
        {
            MessageBox.Show("アクティブなシートがありません。", "シート共有");
            return;
        }

        OnShareButtonPressed(control, activeSheet.Name);
    }

    public async void OnShowCurrentSheetDiffButtonPressed(IRibbonControl control)
    {
        string dialogTitle = "シート差分表示";

        try
        {
            Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook workbook = excelApp.ActiveWorkbook as Excel.Workbook;
            Excel.Worksheet activeSheet = excelApp.ActiveSheet as Excel.Worksheet;

            if (workbook == null || activeSheet == null)
            {
                MessageBox.Show("アクティブなシートがありません。", dialogTitle);
                return;
            }

            WorkbookInfo workbookInfo = WorkbookInfo.CreateFromWorkbook(workbook);
            if (workbookInfo == null)
            {
                string projectName = Assembly.GetExecutingAssembly().GetName().Name;
                MessageBox.Show($"{projectName} で生成されたブックではありません。", dialogTitle);
                return;
            }

            if (workbookInfo.ShareInfo == null)
            {
                MessageBox.Show("このブックには共有先の情報が保存されていません。", dialogTitle);
                return;
            }

            SharedSheetDocument localDocument = CreateSharedSheetDocument(activeSheet);
            if (localDocument == null)
            {
                MessageBox.Show("このシートは共有対象ではありません。", dialogTitle);
                return;
            }

            SharedSheetDocument baseDocument = GetSharedSheetBaseDocument(workbook, localDocument.SheetId);
            SharedSheetDocument remoteDocument = null;

            string shareToken = GitLabAuth.GetOrPromptToken(
                workbookInfo.ShareInfo.BaseUrl,
                workbookInfo.ShareInfo.ProjectId);
            if (!string.IsNullOrWhiteSpace(shareToken))
            {
                await EnsureValidatedShareRefNameAsync(
                    workbookInfo.ShareInfo,
                    shareToken,
                    workbook).ConfigureAwait(true);

                SharedProjectManifest remoteManifest = await TryDownloadSharedProjectManifestAsync(
                    workbookInfo.ShareInfo,
                    workbookInfo.ProjectId,
                    shareToken).ConfigureAwait(true);

                SharedProjectManifestEntry remoteEntry = remoteManifest == null || remoteManifest.Sheets == null
                    ? null
                    : remoteManifest.Sheets.FirstOrDefault(x => x != null && string.Equals(x.SheetId, localDocument.SheetId, StringComparison.Ordinal));

                if (remoteEntry != null)
                {
                    remoteDocument = await TryDownloadSharedSheetDocumentAsync(
                        workbookInfo.ShareInfo,
                        workbookInfo.ProjectId,
                        localDocument.SheetId,
                        shareToken).ConfigureAwait(true);
                }
            }

            var item = new SharedSheetSelectionItem
            {
                SheetName = localDocument.SheetName,
                SheetId = localDocument.SheetId,
                DiffText = BuildSharedSheetDiffText(baseDocument, localDocument, remoteDocument),
                Document = localDocument
            };

            SharedSheetSelectionDialog.ShowDiff(null, item);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString(), dialogTitle);
        }
    }

    public void OnRevertCurrentSheetChangesButtonPressed(IRibbonControl control)
    {
        string dialogTitle = "シート変更を戻す";

        try
        {
            Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook workbook = excelApp.ActiveWorkbook as Excel.Workbook;
            Excel.Worksheet activeSheet = excelApp.ActiveSheet as Excel.Worksheet;

            if (workbook == null || activeSheet == null)
            {
                MessageBox.Show("アクティブなシートがありません。", dialogTitle);
                return;
            }

            WorkbookInfo workbookInfo = WorkbookInfo.CreateFromWorkbook(workbook);
            if (workbookInfo == null)
            {
                string projectName = Assembly.GetExecutingAssembly().GetName().Name;
                MessageBox.Show($"{projectName} で生成されたブックではありません。", dialogTitle);
                return;
            }

            SharedSheetDocument localDocument = CreateSharedSheetDocument(activeSheet);
            if (localDocument == null)
            {
                MessageBox.Show("このシートは共有対象ではありません。", dialogTitle);
                return;
            }

            SharedSheetDocument baseDocument = GetSharedSheetBaseDocument(workbook, localDocument.SheetId);
            if (baseDocument == null)
            {
                MessageBox.Show("このシートを元に戻す共有状態がありません。", dialogTitle);
                return;
            }

            string diffText = BuildSharedSheetDiffText(baseDocument, localDocument, null);
            if (string.Equals(diffText, "差分はありません。", StringComparison.Ordinal))
            {
                MessageBox.Show("戻す変更はありません。", dialogTitle);
                return;
            }

            var item = new SharedSheetSelectionItem
            {
                SheetName = localDocument.SheetName,
                SheetId = localDocument.SheetId,
                DiffText = diffText,
                Document = localDocument
            };

            bool confirmed = SharedSheetSelectionDialog.TryShowRevertConfirmation(
                null,
                item,
                "このシートの未共有の変更を戻します。この操作は元に戻せません。本当に戻しますか？",
                "戻す");
            if (!confirmed)
            {
                return;
            }

            InitializeLoggerForWorkbookSession(workbook, "shared-revert");

            SheetValuesInfo currentSheetValuesInfo = SheetValuesInfo.CreateFromSheet(activeSheet);
            if (currentSheetValuesInfo == null)
            {
                MessageBox.Show("このシートは共有対象ではありません。", dialogTitle);
                return;
            }

            object[,] revertedValues = BuildSharedSheetRevertValues(currentSheetValuesInfo, baseDocument);
            int revertedCellCount = LogSharedRevertDifferences(activeSheet, currentSheetValuesInfo, revertedValues);

            using (ExcelUiSuspendScope uiSuspendScope = TryCreateExcelUiSuspendScope(workbook))
            {
                currentSheetValuesInfo.Range.Value2 = revertedValues;
            }

            DialogResult openLog = MessageBox.Show(
                "シートの変更を取り消しました。" + Environment.NewLine +
                "変更セル数: " + revertedCellCount + Environment.NewLine + Environment.NewLine +
                "ログファイルを開きますか？",
                dialogTitle,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information);
            if (openLog == DialogResult.Yes)
            {
                FileLogger.OpenLog();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString(), dialogTitle);
        }
    }

    public void OnShareButtonPressed(IRibbonControl control)
    {
        OnShareButtonPressed(control, null);
    }

    private async void OnShareButtonPressed(IRibbonControl control, string targetSheetName)
    {
        if (SynchronizationContext.Current == null)
        {
            SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
        }

        Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
        string dialogTitle = string.IsNullOrWhiteSpace(targetSheetName) ? "変更共有" : "シート共有";
        Excel.Workbook workbook = excelApp.ActiveWorkbook as Excel.Workbook;
        PullProgressForm progressForm = null;

        if (workbook == null)
        {
            MessageBox.Show("アクティブなブックがありません。", dialogTitle);
            return;
        }

        try
        {
            WorkbookInfo workbookInfo = WorkbookInfo.CreateFromWorkbook(workbook);
            if (workbookInfo == null)
            {
                string projectName = Assembly.GetExecutingAssembly().GetName().Name;
                MessageBox.Show($"{projectName} で生成されたブックではありません。", dialogTitle);
                return;
            }

            if (!HasShareSettings(workbookInfo.ShareInfo))
            {
                MessageBox.Show("このブックには共有先の情報が保存されていません。先に最新版取得を実行してください。", dialogTitle);
                return;
            }

            GitLabLastInput pullInfo = workbookInfo.PullInfo;
            GitLabShareInfo shareInfo = workbookInfo.ShareInfo;
            bool pullEnabled = IsPullSourceEnabled(pullInfo) && HasPullSourceSettings(pullInfo);

            if (pullEnabled)
            {
                string pullToken = GitLabAuth.GetOrPromptToken(
                    pullInfo.BaseUrl,
                    pullInfo.ProjectId);
                if (string.IsNullOrWhiteSpace(pullToken))
                {
                    MessageBox.Show("共有をキャンセルしました（トークン未入力）", dialogTitle);
                    return;
                }

                string currentCommitId = await GitLabClient.GetCommitIdAsync(
                    pullInfo.BaseUrl,
                    pullInfo.ProjectId,
                    pullInfo.RefName,
                    pullToken).ConfigureAwait(true);

                if (!string.IsNullOrWhiteSpace(currentCommitId) &&
                    !string.Equals(workbookInfo.PullCommitId, currentCommitId, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("Pull 元が最新ではありません。先に最新版取得を実行してください。", dialogTitle);
                    return;
                }
            }

            string shareToken = GitLabAuth.GetOrPromptToken(
                shareInfo.BaseUrl,
                shareInfo.ProjectId);
            if (string.IsNullOrWhiteSpace(shareToken))
            {
                MessageBox.Show("共有をキャンセルしました（トークン未入力）", dialogTitle);
                return;
            }

            await EnsureValidatedShareRefNameAsync(
                shareInfo,
                shareToken,
                workbook).ConfigureAwait(true);

            SharedProjectManifest remoteManifest = await TryDownloadSharedProjectManifestAsync(
                shareInfo,
                workbookInfo.ProjectId,
                shareToken).ConfigureAwait(true);

            List<SharedSheetSelectionItem> selectionItems = CollectSharedSheetSelectionItems(workbook, remoteManifest, targetSheetName);
            if (selectionItems.Count == 0)
            {
                MessageBox.Show("共有する変更はありません。", dialogTitle);
                return;
            }

            List<string> conflictSheetNames = await MergeSharedSelectionItemsWithRemoteAsync(
                workbook,
                workbookInfo.ShareInfo,
                shareToken,
                selectionItems,
                remoteManifest).ConfigureAwait(true);

            if (conflictSheetNames.Count > 0)
            {
                SharedSheetSelectionDialog.ShowConflictReview(null, selectionItems);
                return;
            }

            if (!selectionItems.Any(x => x != null && x.Document != null))
            {
                MessageBox.Show("共有する変更はありません。", dialogTitle);
                return;
            }

            List<SharedSheetSelectionItem> selectedItems;
            if (!SharedSheetSelectionDialog.TryShow(null, selectionItems, out selectedItems))
            {
                return;
            }

            if (selectedItems == null || selectedItems.Count == 0)
            {
                MessageBox.Show("共有するシートが選択されていません。", dialogTitle);
                return;
            }

            InitializeLoggerForWorkbookSession(workbook, "shared-commit");

            progressForm = new PullProgressForm(dialogTitle, "共有中...");
            progressForm.Show();
            Action<string> shareProgressReporter = message =>
            {
                progressForm.SetStatusText(message);
                progressForm.AppendLine(message);
            };

            shareProgressReporter("変更共有を開始します");

            await UploadSharedSheetsAsync(
                workbook,
                shareInfo,
                shareToken,
                remoteManifest,
                selectedItems,
                shareProgressReporter).ConfigureAwait(true);

            workbook.Save();
            shareProgressReporter("共有が完了しました");
            progressForm.ShowContinueButton("閉じる", "共有が完了しました");
            await progressForm.WaitForContinueAsync();
            progressForm.CloseForm();
            progressForm = null;
            return;
        }
        catch (Exception ex)
        {
            if (progressForm != null)
            {
                progressForm.CloseForm();
                progressForm = null;
            }
            MessageBox.Show(ex.ToString(), dialogTitle);
        }
        finally
        {
            if (progressForm != null)
            {
                progressForm.CloseForm();
            }
        }
    }

    private async Task<PullExecutionResult> ExecutePullAsync(GitLabLastInput input, Action<string> progressReporter = null)
    {
        if (input == null)
        {
            throw new ArgumentNullException(nameof(input));
        }

        string baseUrl = input.BaseUrl;
        string projectId = input.ProjectId;
        string refName = input.RefName;
        string filePath = GitLabPathResolver.NormalizeGitLabRelativePath(input.FilePath);

        string token = GitLabAuth.GetOrPromptToken(baseUrl, projectId);
        if (string.IsNullOrEmpty(token))
        {
            System.Windows.Forms.MessageBox.Show("同期をキャンセルしました（トークン未入力）");
            return null;
        }

        string workRoot = CreatePullWorkRoot();

        FileLogger.InitializeForSession(workRoot, "pull", timestamped: false);
        progressReporter?.Invoke("PullWork を作成しました");

        string refCommitId = await GitLabClient.GetCommitIdAsync(
            baseUrl,
            projectId,
            refName,
            token).ConfigureAwait(false);
        progressReporter?.Invoke("取得元の更新状態を確認しました");

        string normalizedEntryPath = string.IsNullOrEmpty(filePath)
            ? null
            : GitLabPathResolver.NormalizeGitLabFilePathStrict(filePath);

        var sessionLog = new PullSessionLog(workRoot, normalizedEntryPath);

        currentPullSession = new PullSessionContext
        {
            BaseUrl = baseUrl,
            ProjectId = projectId,
            RefName = refName,
            Token = token,
            WorkRoot = workRoot,
            EntryGitLabRelativePath = normalizedEntryPath,
            ManifestPath = CreatePullManifestPath(workRoot),
            SessionLog = sessionLog,
            ExpandedArchiveFolders = new HashSet<string>(StringComparer.OrdinalIgnoreCase),
            ProgressReporter = progressReporter
        };

        progressReporter?.Invoke("エントリーファイルを取得しています");
        await EnsureFileInWorkRootAsync(
            baseUrl,
            projectId,
            refName,
            token,
            workRoot,
            filePath,
            null,
            sessionLog);

        string entryLocalPath = BuildLocalPathInWorkRoot(workRoot, normalizedEntryPath);
        progressReporter?.Invoke("parse を実行しています");
        bool parseSucceeded = RunParsePipeline(entryLocalPath, false);
        if (!parseSucceeded)
        {
            throw new InvalidOperationException("Pull dependency discovery failed during parse.");
        }

        progressReporter?.Invoke("画像依存を確認しています");
        await EnsureImageDependenciesInWorkRootAsync(currentPullSession, entryLocalPath);

        string jsonFilePath = TxtToJsonPath(entryLocalPath);
        progressReporter?.Invoke("ダウンロードと準備が完了しました");
        return new PullExecutionResult
        {
            EntryLocalPath = entryLocalPath,
            JsonFilePath = jsonFilePath,
            NormalizedEntryPath = normalizedEntryPath,
            RefCommitId = refCommitId
        };
    }

    private static void ClearPullSessionState()
    {
        currentPullSession = null;
        currentGitLabBaseFileRelativePath = null;
    }

    private static string CreatePullWorkRoot()
    {
        string workRoot = Path.Combine(
            GetPullWorkParentDirectory(),
            DateTime.Now.ToString("yyyyMMdd_HHmmss_fff"));

        Directory.CreateDirectory(workRoot);
        return workRoot;
    }

    private static string GetPullWorkParentDirectory()
    {
        return Path.Combine(
            Path.GetTempPath(),
            "SheetRenderer",
            "PullWork");
    }

    internal static void CleanupOldPullWorkDirectories()
    {
        try
        {
            string parentDirectory = GetPullWorkParentDirectory();
            if (!Directory.Exists(parentDirectory))
            {
                return;
            }

            DateTime cutoff = DateTime.Now.AddDays(-1);
            string[] workDirectories = Directory.GetDirectories(parentDirectory);
            foreach (string workDirectory in workDirectories)
            {
                try
                {
                    DateTime lastWriteTime = Directory.GetLastWriteTime(workDirectory);
                    if (lastWriteTime >= cutoff)
                    {
                        continue;
                    }

                    Directory.Delete(workDirectory, true);
                }
                catch (Exception ex)
                {
                    _ = ex;
                }
            }
        }
        catch (Exception ex)
        {
            _ = ex;
        }
    }

    private static string GetGitLabParentFolder(string gitLabFilePath)
    {
        string normalized = GitLabPathResolver.NormalizeGitLabRelativePath(gitLabFilePath);
        int idx = normalized.LastIndexOf('/');
        if (idx < 0)
        {
            return string.Empty;
        }

        return normalized.Substring(0, idx);
    }

    private static async Task<List<GitLabTreeItem>> ListDirectBlobItemsAsync(
        string baseUrl,
        string projectId,
        string folder,
        string refName,
        string token)
    {
        var items = await GitLabClient.ListTreeItemsAsync(baseUrl, projectId, folder, refName, token).ConfigureAwait(false);
        var blobs = new List<GitLabTreeItem>();
        foreach (var item in items)
        {
            if (item == null)
            {
                continue;
            }

            if (!string.Equals(item.Type, "blob", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            blobs.Add(item);
        }

        return blobs;
    }

    private static async Task<byte[]> DownloadBlobByIdAsync(
        string baseUrl,
        string projectId,
        string blobId,
        string token)
    {
        return await GitLabClient.DownloadBlobRawAsync(baseUrl, projectId, blobId, token).ConfigureAwait(false);
    }

    private static string SaveBlobToWorkRoot(string workRoot, string relativePath, byte[] bytes)
    {
        string normalizedRelativePath = GitLabPathResolver.NormalizeGitLabFilePathStrict(relativePath);
        string fullPath = BuildLocalPathInWorkRoot(workRoot, normalizedRelativePath);
        EnsureDirectoryForLocalPath(fullPath);

        File.WriteAllBytes(fullPath, bytes ?? new byte[0]);
        return normalizedRelativePath;
    }

    private static async Task<string> EnsureFileInWorkRootAsync(
        string baseUrl,
        string projectId,
        string refName,
        string token,
        string workRoot,
        string gitLabRelativePath,
        string requestedPathForError = null,
        PullSessionLog sessionLog = null)
    {
        string normalizedRelativePath = GitLabPathResolver.NormalizeGitLabFilePathStrict(gitLabRelativePath);
        string localPath = BuildLocalPathInWorkRoot(workRoot, normalizedRelativePath);

        if (File.Exists(localPath))
        {
            FileLogger.Info("[PullLazyRead] local cache hit: " + normalizedRelativePath);
            ReportPullProgress("ローカル確認: " + normalizedRelativePath);
            AddFileReadTrace("[local-hit] gitlabRelative=" + normalizedRelativePath);
            if (sessionLog != null)
            {
                sessionLog.Add(PullFileActionType.AlreadyExists, normalizedRelativePath);
            }

            return Path.GetFullPath(localPath);
        }

        PullSessionContext sessionContext = currentPullSession;
        if (ShouldUseTopLevelFolderArchive(normalizedRelativePath) &&
            sessionContext != null &&
            string.Equals(
                Path.GetFullPath(sessionContext.WorkRoot ?? string.Empty),
                Path.GetFullPath(workRoot),
                StringComparison.OrdinalIgnoreCase))
        {
            bool archiveExpanded = await EnsureTopLevelFolderArchiveInWorkRootAsync(
                sessionContext,
                normalizedRelativePath,
                sessionLog).ConfigureAwait(false);

            if (archiveExpanded && File.Exists(localPath))
            {
                return Path.GetFullPath(localPath);
            }

            throw new FileNotFoundException(
                "File was not found after expanding the top-level folder archive.",
                localPath);
        }

        FileLogger.Info("[PullLazyRead] fetching from GitLab: " + normalizedRelativePath);
        ReportPullProgress("ファイル取得: " + normalizedRelativePath);
        AddFileReadTrace("[download] gitlabRelative=" + normalizedRelativePath);

        string parentFolder = GetGitLabParentFolder(normalizedRelativePath);
        string fileName = GetGitLabFileName(normalizedRelativePath);
        var items = await GitLabClient.ListTreeItemsAsync(baseUrl, projectId, parentFolder, refName, token).ConfigureAwait(false);

        GitLabTreeItem fileItem = FindBlobItemByName(
            items,
            fileName,
            requestedPathForError ?? normalizedRelativePath,
            normalizedRelativePath,
            parentFolder);
        byte[] bytes = await DownloadBlobByIdAsync(baseUrl, projectId, fileItem.Id, token).ConfigureAwait(false);
        SaveBlobToWorkRoot(workRoot, normalizedRelativePath, bytes);
        if (sessionLog != null)
        {
            sessionLog.Add(PullFileActionType.LazyFileRead, normalizedRelativePath);
            UpdatePullManifestIfAvailable(currentPullSession);
        }

        return Path.GetFullPath(localPath);
    }

    private static bool ShouldUseTopLevelFolderArchive(string gitLabRelativePath)
    {
        string normalizedRelativePath = GitLabPathResolver.NormalizeGitLabFilePathStrict(gitLabRelativePath);
        return normalizedRelativePath.IndexOf('/') >= 0;
    }

    private static string GetTopLevelFolderName(string gitLabRelativePath)
    {
        string normalizedRelativePath = GitLabPathResolver.NormalizeGitLabFilePathStrict(gitLabRelativePath);
        int idx = normalizedRelativePath.IndexOf('/');
        if (idx < 0)
        {
            return null;
        }

        return normalizedRelativePath.Substring(0, idx);
    }

    private static async Task<bool> EnsureTopLevelFolderArchiveInWorkRootAsync(
        PullSessionContext sessionContext,
        string gitLabRelativePath,
        PullSessionLog sessionLog)
    {
        if (sessionContext == null)
        {
            throw new ArgumentNullException(nameof(sessionContext));
        }

        string topLevelFolder = GetTopLevelFolderName(gitLabRelativePath);
        if (string.IsNullOrWhiteSpace(topLevelFolder))
        {
            return false;
        }

        if (sessionContext.ExpandedArchiveFolders == null)
        {
            sessionContext.ExpandedArchiveFolders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        if (sessionContext.ExpandedArchiveFolders.Contains(topLevelFolder))
        {
            FileLogger.Info("[archive-hit] " + topLevelFolder);
            ReportPullProgress("アーカイブ再利用: " + topLevelFolder);
            return true;
        }

        FileLogger.Info("[archive-download] " + topLevelFolder);
        ReportPullProgress("フォルダ取得: " + topLevelFolder);
        byte[] archiveBytes = await GitLabClient.DownloadArchiveZipAsync(
            sessionContext.BaseUrl,
            sessionContext.ProjectId,
            sessionContext.RefName,
            topLevelFolder,
            sessionContext.Token).ConfigureAwait(false);

        ExtractTopLevelFolderArchiveToWorkRoot(
            sessionContext.WorkRoot,
            topLevelFolder,
            archiveBytes,
            sessionLog);

        sessionContext.ExpandedArchiveFolders.Add(topLevelFolder);
        FileLogger.Info("[archive-ready] " + topLevelFolder);
        ReportPullProgress("フォルダ展開完了: " + topLevelFolder);
        return true;
    }

    private static void ExtractTopLevelFolderArchiveToWorkRoot(
        string workRoot,
        string topLevelFolder,
        byte[] archiveBytes,
        PullSessionLog sessionLog)
    {
        using (var stream = new MemoryStream(archiveBytes ?? new byte[0]))
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false))
        {
            foreach (var entry in archive.Entries)
            {
                string relativePath = TryGetArchiveEntryRelativePath(entry.FullName, topLevelFolder);
                if (string.IsNullOrWhiteSpace(relativePath))
                {
                    continue;
                }

                string destinationPath = BuildLocalPathInWorkRoot(workRoot, relativePath);
                EnsureDirectoryForLocalPath(destinationPath);

                if (string.IsNullOrEmpty(entry.Name))
                {
                    Directory.CreateDirectory(destinationPath);
                    continue;
                }

                using (var entryStream = entry.Open())
                using (var destinationStream = File.Create(destinationPath))
                {
                    entryStream.CopyTo(destinationStream);
                }

                if (sessionLog != null)
                {
                    sessionLog.Add(PullFileActionType.LazyFileRead, relativePath);
                }
            }
        }
    }

    private static string TryGetArchiveEntryRelativePath(string archiveEntryName, string topLevelFolder)
    {
        if (string.IsNullOrWhiteSpace(archiveEntryName) || string.IsNullOrWhiteSpace(topLevelFolder))
        {
            return null;
        }

        string normalizedEntryName = archiveEntryName.Replace('\\', '/').TrimStart('/');
        string normalizedTopLevelFolder = GitLabPathResolver.NormalizeGitLabFilePathStrict(topLevelFolder);
        string folderMarker = normalizedTopLevelFolder + "/";

        if (normalizedEntryName.StartsWith(folderMarker, StringComparison.OrdinalIgnoreCase))
        {
            return GitLabPathResolver.NormalizeGitLabFilePathStrict(normalizedEntryName);
        }

        string marker = "/" + folderMarker;
        int idx = normalizedEntryName.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (idx < 0)
        {
            return null;
        }

        return GitLabPathResolver.NormalizeGitLabFilePathStrict(
            normalizedEntryName.Substring(idx + 1));
    }

    private static PullManifestReuseContext TryLoadReusablePullManifest(
        string currentWorkRoot,
        string baseUrl,
        string projectId,
        string refName,
        string entryFilePath)
    {
        try
        {
            string candidateWorkRoot = FindLatestReusablePullWorkRoot(currentWorkRoot);
            if (string.IsNullOrEmpty(candidateWorkRoot))
            {
                return null;
            }

            string manifestPath = FindPullManifestPath(candidateWorkRoot);
            if (string.IsNullOrEmpty(manifestPath) || !File.Exists(manifestPath))
            {
                FileLogger.Info("[manifest-reuse-miss] manifest not found: " + candidateWorkRoot);
                return null;
            }

            PullManifest manifest = ReadPullManifest(manifestPath);
            if (!IsReusablePullManifestMatch(manifest, baseUrl, projectId, refName, entryFilePath))
            {
                FileLogger.Info("[manifest-reuse-miss] target mismatch: " + manifestPath);
                return null;
            }

            var filesByGitLabRelativePath = BuildManifestFileMap(manifest.Files);
            FileLogger.Info("[manifest-reuse-ready] sourceWorkRoot=" + candidateWorkRoot + " manifest=" + manifestPath);

            return new PullManifestReuseContext
            {
                SourceWorkRoot = candidateWorkRoot,
                ManifestPath = manifestPath,
                Manifest = manifest,
                FilesByGitLabRelativePath = filesByGitLabRelativePath
            };
        }
        catch (Exception ex)
        {
            FileLogger.Info("[manifest-reuse-disabled] " + ex.Message);
            return null;
        }
    }

    private static string FindLatestReusablePullWorkRoot(string currentWorkRoot)
    {
        string parentDirectory = GetPullWorkParentDirectory();
        if (!Directory.Exists(parentDirectory))
        {
            return null;
        }

        string normalizedCurrentWorkRoot = string.IsNullOrWhiteSpace(currentWorkRoot)
            ? null
            : Path.GetFullPath(currentWorkRoot).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var candidates = new List<string>(Directory.GetDirectories(parentDirectory));
        candidates.Sort(StringComparer.OrdinalIgnoreCase);

        for (int i = candidates.Count - 1; i >= 0; i--)
        {
            string candidate = Path.GetFullPath(candidates[i]).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            if (string.Equals(candidate, normalizedCurrentWorkRoot, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            return candidate;
        }

        return null;
    }

    private static string FindPullManifestPath(string workRoot)
    {
        if (string.IsNullOrWhiteSpace(workRoot) || !Directory.Exists(workRoot))
        {
            return null;
        }

        string[] manifestPaths = Directory.GetFiles(workRoot, ".sheetrenderer-pull-manifest*.json", SearchOption.TopDirectoryOnly);
        if (manifestPaths.Length == 0)
        {
            return null;
        }

        Array.Sort(manifestPaths, StringComparer.OrdinalIgnoreCase);
        return manifestPaths[manifestPaths.Length - 1];
    }

    private static PullManifest ReadPullManifest(string manifestPath)
    {
        string json = File.ReadAllText(manifestPath, Encoding.UTF8);
        var manifest = JsonSerializer.Deserialize<PullManifest>(
            json,
            new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });

        if (manifest == null)
        {
            throw new InvalidOperationException("Pull manifest is empty.");
        }

        return manifest;
    }

    private static bool IsReusablePullManifestMatch(
        PullManifest manifest,
        string baseUrl,
        string projectId,
        string refName,
        string entryFilePath)
    {
        if (manifest == null)
        {
            return false;
        }

        string normalizedManifestEntryFilePath = string.IsNullOrWhiteSpace(manifest.EntryFilePath)
            ? null
            : GitLabPathResolver.NormalizeGitLabFilePathStrict(manifest.EntryFilePath);
        string normalizedEntryFilePath = string.IsNullOrWhiteSpace(entryFilePath)
            ? null
            : GitLabPathResolver.NormalizeGitLabFilePathStrict(entryFilePath);

        return string.Equals(manifest.BaseUrl ?? string.Empty, baseUrl ?? string.Empty, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(manifest.ProjectId ?? string.Empty, projectId ?? string.Empty, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(manifest.RefName ?? string.Empty, refName ?? string.Empty, StringComparison.Ordinal) &&
            string.Equals(normalizedManifestEntryFilePath ?? string.Empty, normalizedEntryFilePath ?? string.Empty, StringComparison.OrdinalIgnoreCase);
    }

    private static Dictionary<string, PullManifestFileRecord> BuildManifestFileMap(List<PullManifestFileRecord> files)
    {
        var map = new Dictionary<string, PullManifestFileRecord>(StringComparer.OrdinalIgnoreCase);
        if (files == null)
        {
            return map;
        }

        foreach (var file in files)
        {
            if (file == null || string.IsNullOrWhiteSpace(file.GitLabRelativePath))
            {
                continue;
            }

            string normalizedRelativePath;
            try
            {
                normalizedRelativePath = GitLabPathResolver.NormalizeGitLabFilePathStrict(file.GitLabRelativePath);
            }
            catch
            {
                continue;
            }

            map[normalizedRelativePath] = file;
        }

        return map;
    }

    private static bool TryReuseManifestFileToWorkRoot(
        PullManifestReuseContext reuseContext,
        string targetWorkRoot,
        string gitLabRelativePath,
        out string savedRelativePath)
    {
        savedRelativePath = null;

        if (reuseContext == null || reuseContext.FilesByGitLabRelativePath == null)
        {
            return false;
        }

        string normalizedRelativePath = GitLabPathResolver.NormalizeGitLabFilePathStrict(gitLabRelativePath);
        PullManifestFileRecord fileRecord;
        if (!reuseContext.FilesByGitLabRelativePath.TryGetValue(normalizedRelativePath, out fileRecord))
        {
            return false;
        }

        string sourcePath = ResolveManifestRecordLocalPath(reuseContext.SourceWorkRoot, fileRecord, normalizedRelativePath);
        if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
        {
            FileLogger.Info("[manifest-reuse-miss] missing file: " + normalizedRelativePath);
            return false;
        }

        string destinationPath = BuildLocalPathInWorkRoot(targetWorkRoot, normalizedRelativePath);
        EnsureDirectoryForLocalPath(destinationPath);
        File.Copy(sourcePath, destinationPath, true);

        FileLogger.Info("[manifest-reuse-hit] " + normalizedRelativePath);
        savedRelativePath = normalizedRelativePath;
        return true;
    }

    private static string ResolveManifestRecordLocalPath(
        string sourceWorkRoot,
        PullManifestFileRecord fileRecord,
        string normalizedRelativePath)
    {
        if (fileRecord == null)
        {
            return null;
        }

        string localPath = fileRecord.LocalPath;
        if (string.IsNullOrWhiteSpace(localPath))
        {
            return BuildLocalPathInWorkRoot(sourceWorkRoot, normalizedRelativePath);
        }

        if (Path.IsPathRooted(localPath))
        {
            return Path.GetFullPath(localPath);
        }

        return Path.GetFullPath(Path.Combine(sourceWorkRoot, localPath));
    }

    private static async Task EnsureImageDependenciesInWorkRootAsync(
        PullSessionContext sessionContext,
        string entryLocalPath)
    {
        if (sessionContext == null)
        {
            throw new ArgumentNullException(nameof(sessionContext));
        }

        if (string.IsNullOrWhiteSpace(entryLocalPath))
        {
            throw new ArgumentException("entryLocalPath is required.", nameof(entryLocalPath));
        }

        string jsonFilePath = TxtToJsonPath(entryLocalPath);
        if (!File.Exists(jsonFilePath))
        {
            throw new FileNotFoundException("Parsed json file was not created in PullWork.", jsonFilePath);
        }

        string json = File.ReadAllText(jsonFilePath, Encoding.UTF8);
        JsonNode rootNode = JsonNode.Parse(json);
        if (rootNode == null)
        {
            throw new InvalidOperationException("Parsed json is empty.");
        }

        var imageFilePaths = CollectImageFilePaths(rootNode);
        foreach (string imageFilePath in imageFilePaths)
        {
            if (string.IsNullOrWhiteSpace(imageFilePath))
            {
                continue;
            }

            string gitLabRelativePath = ResolvePullImageGitLabRelativePath(
                sessionContext.EntryGitLabRelativePath,
                imageFilePath);

            FileLogger.Info("[pull-image] image=" + imageFilePath + " gitlabRelative=" + gitLabRelativePath);
            ReportPullProgress("画像取得: " + gitLabRelativePath);

            await EnsureFileInWorkRootAsync(
                sessionContext.BaseUrl,
                sessionContext.ProjectId,
                sessionContext.RefName,
                sessionContext.Token,
                sessionContext.WorkRoot,
                gitLabRelativePath,
                imageFilePath,
                sessionContext.SessionLog);
        }
    }

    private static string ResolvePullImageGitLabRelativePath(
        string entryGitLabRelativePath,
        string imageFilePath)
    {
        if (string.IsNullOrWhiteSpace(entryGitLabRelativePath))
        {
            throw new ArgumentException("entryGitLabRelativePath is required.", nameof(entryGitLabRelativePath));
        }

        if (string.IsNullOrWhiteSpace(imageFilePath))
        {
            throw new ArgumentException("imageFilePath is required.", nameof(imageFilePath));
        }

        string entryProjectFolder = GetEntryProjectFolderFromEntryGitLabRelativePath(entryGitLabRelativePath);
        string combinedPath = string.IsNullOrWhiteSpace(entryProjectFolder)
            ? imageFilePath
            : entryProjectFolder + "/" + imageFilePath;

        return GitLabPathResolver.CanonicalizeGitLabRelativePath(
            combinedPath,
            "requestedPath",
            imageFilePath,
            entryProjectFolder);
    }

    private static string GetEntryProjectFolderFromEntryGitLabRelativePath(string entryGitLabRelativePath)
    {
        string normalizedEntryPath = GitLabPathResolver.NormalizeGitLabFilePathStrict(entryGitLabRelativePath);
        string entryParentFolder = GetGitLabParentFolder(normalizedEntryPath);
        string folderName = GetGitLabFileName(entryParentFolder);

        if (string.Equals(folderName, "source", StringComparison.OrdinalIgnoreCase))
        {
            return GetGitLabParentFolder(entryParentFolder);
        }

        return entryParentFolder;
    }

    private static string WritePullManifest(PullSessionContext sessionContext)
    {
        if (sessionContext == null)
        {
            throw new ArgumentNullException(nameof(sessionContext));
        }

        if (sessionContext.SessionLog == null)
        {
            throw new ArgumentNullException(nameof(sessionContext.SessionLog));
        }

        string workRoot = Path.GetFullPath(sessionContext.WorkRoot ?? string.Empty);
        string manifestPath = sessionContext.ManifestPath;
        if (string.IsNullOrWhiteSpace(manifestPath))
        {
            manifestPath = CreatePullManifestPath(workRoot);
            sessionContext.ManifestPath = manifestPath;
        }
        else
        {
            manifestPath = Path.GetFullPath(manifestPath);
        }

        if (ShouldReallocateManifestPath(sessionContext, workRoot, manifestPath))
        {
            manifestPath = CreatePullManifestPath(workRoot);
            sessionContext.ManifestPath = manifestPath;
            sessionContext.ManifestHasBeenWritten = false;
        }

        var manifest = new PullManifest
        {
            BaseUrl = sessionContext.BaseUrl,
            ProjectId = sessionContext.ProjectId,
            RefName = sessionContext.RefName,
            EntryFilePath = sessionContext.EntryGitLabRelativePath,
            WorkRoot = workRoot,
            CreatedAt = DateTime.UtcNow.ToString("o"),
            Files = BuildPullManifestFiles(sessionContext.SessionLog)
        };

        string json = JsonSerializer.Serialize(
            manifest,
            new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true
            });

        File.WriteAllText(manifestPath, json, new UTF8Encoding(false));
        sessionContext.ManifestHasBeenWritten = true;
        return manifestPath;
    }

    private static bool ShouldReallocateManifestPath(PullSessionContext sessionContext, string workRoot, string manifestPath)
    {
        if (Directory.Exists(manifestPath))
        {
            return true;
        }

        bool overlapsPulledFile = SessionContainsLocalPath(sessionContext.SessionLog, workRoot, manifestPath);
        if (overlapsPulledFile)
        {
            return true;
        }

        if (!sessionContext.ManifestHasBeenWritten && File.Exists(manifestPath))
        {
            return true;
        }

        return false;
    }

    private static bool SessionContainsLocalPath(PullSessionLog sessionLog, string workRoot, string targetPath)
    {
        if (sessionLog == null)
        {
            return false;
        }

        string normalizedTargetPath = Path.GetFullPath(targetPath);
        foreach (var activity in sessionLog.GetActivities())
        {
            string relativePath = activity.RelativePath;
            if (string.IsNullOrWhiteSpace(relativePath))
            {
                continue;
            }

            string activityPath;
            try
            {
                activityPath = BuildLocalPathInWorkRoot(workRoot, relativePath);
            }
            catch (Exception)
            {
                continue;
            }

            if (string.Equals(
                Path.GetFullPath(activityPath),
                normalizedTargetPath,
                StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static void UpdatePullManifestIfAvailable(PullSessionContext sessionContext)
    {
        if (sessionContext == null ||
            string.IsNullOrWhiteSpace(sessionContext.WorkRoot) ||
            sessionContext.SessionLog == null)
        {
            return;
        }

        string manifestPath = WritePullManifest(sessionContext);
        FileLogger.Info("[PullManifest] updated: " + manifestPath);
    }

    private static string CreatePullManifestPath(string workRoot)
    {
        if (string.IsNullOrWhiteSpace(workRoot))
        {
            throw new ArgumentException("workRoot is required.", nameof(workRoot));
        }

        const string manifestPrefix = ".sheetrenderer-pull-manifest";
        string manifestPath = Path.Combine(workRoot, manifestPrefix + ".json");
        if (!File.Exists(manifestPath) && !Directory.Exists(manifestPath))
        {
            return manifestPath;
        }

        for (int i = 1; i < 1000; i++)
        {
            string candidatePath = Path.Combine(workRoot, manifestPrefix + "-" + i + ".json");
            if (!File.Exists(candidatePath) && !Directory.Exists(candidatePath))
            {
                return candidatePath;
            }
        }

        throw new IOException("Unable to allocate a pull manifest path.");
    }

    private static List<PullManifestFileRecord> BuildPullManifestFiles(PullSessionLog sessionLog)
    {
        var files = new List<PullManifestFileRecord>();
        var fileMap = new Dictionary<string, PullManifestFileRecord>(StringComparer.OrdinalIgnoreCase);

        foreach (var activity in sessionLog.GetActivities())
        {
            string sourceKind = ToManifestSourceKind(activity.ActionType);
            if (sourceKind == null)
            {
                continue;
            }

            string relativePath = GitLabPathResolver.NormalizeGitLabFilePathStrict(activity.RelativePath);
            PullManifestFileRecord record;
            if (!fileMap.TryGetValue(relativePath, out record))
            {
                record = new PullManifestFileRecord
                {
                    GitLabRelativePath = relativePath,
                    LocalPath = relativePath.Replace('/', Path.DirectorySeparatorChar),
                    SourceKind = sourceKind
                };

                fileMap.Add(relativePath, record);
                files.Add(record);
                continue;
            }

            if (string.Equals(record.SourceKind, "initial-folder-download", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            record.SourceKind = sourceKind;
        }

        return files;
    }

    private static string ToManifestSourceKind(PullFileActionType actionType)
    {
        switch (actionType)
        {
            case PullFileActionType.InitialFolderDownload:
                return "initial-folder-download";
            case PullFileActionType.LazyFileRead:
                return "lazy-file-read";
            default:
                return null;
        }
    }

    private static string BuildLocalPathInWorkRoot(string workRoot, string gitLabRelativePath)
    {
        string normalizedRoot = Path.GetFullPath(workRoot).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        string normalizedRelativePath = GitLabPathResolver.NormalizeGitLabFilePathStrict(gitLabRelativePath);
        string combinedPath = Path.Combine(normalizedRoot, normalizedRelativePath.Replace('/', Path.DirectorySeparatorChar));
        string fullPath = Path.GetFullPath(combinedPath);

        string rootWithSeparator = normalizedRoot + Path.DirectorySeparatorChar;
        if (!string.Equals(fullPath, normalizedRoot, StringComparison.OrdinalIgnoreCase) &&
            !fullPath.StartsWith(rootWithSeparator, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                "Resolved local path escaped WorkRoot. " +
                "workRoot='" + workRoot + "', " +
                "gitLabRelativePath='" + gitLabRelativePath + "', " +
                "resolvedPath='" + fullPath + "'.");
        }

        return fullPath;
    }

    private static GitLabTreeItem FindBlobItemByName(
        IEnumerable<GitLabTreeItem> items,
        string fileName,
        string requestedPath,
        string resolvedRelativePath,
        string parentFolder)
    {
        GitLabTreeItem foundByName = null;

        foreach (var item in items)
        {
            if (item == null)
            {
                continue;
            }

            if (!string.Equals(item.Name, fileName, StringComparison.Ordinal))
            {
                continue;
            }

            foundByName = item;
            break;
        }

        if (foundByName == null)
        {
            throw new FileNotFoundException(
                "GitLab file not found. " +
                "requestedPath='" + requestedPath + "', " +
                "resolvedRelativePath='" + resolvedRelativePath + "', " +
                "parentFolder='" + parentFolder + "', " +
                "fileName='" + fileName + "'.");
        }

        if (!string.Equals(foundByName.Type, "blob", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                "GitLab item is not a blob. " +
                "requestedPath='" + requestedPath + "', " +
                "resolvedRelativePath='" + resolvedRelativePath + "', " +
                "parentFolder='" + parentFolder + "', " +
                "fileName='" + fileName + "', " +
                "actualType='" + (foundByName.Type ?? string.Empty) + "'.");
        }

        if (string.IsNullOrEmpty(foundByName.Id))
        {
            throw new InvalidOperationException(
                "GitLab blob id is empty. " +
                "requestedPath='" + requestedPath + "', " +
                "resolvedRelativePath='" + resolvedRelativePath + "', " +
                "parentFolder='" + parentFolder + "', " +
                "fileName='" + fileName + "'.");
        }

        return foundByName;
    }

    private static string GetGitLabFileName(string gitLabRelativePath)
    {
        string normalized = GitLabPathResolver.NormalizeGitLabRelativePath(gitLabRelativePath);
        int idx = normalized.LastIndexOf('/');
        if (idx < 0)
        {
            return normalized;
        }

        return normalized.Substring(idx + 1);
    }

    private static void EnsureDirectoryForLocalPath(string fullPath)
    {
        string dir = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrEmpty(dir))
        {
            return;
        }

        Directory.CreateDirectory(dir);
    }

    private static string ToGitLabRelativePath(string workRoot, string localAbsolutePath)
    {
        string normalizedRoot = Path.GetFullPath(workRoot).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        string normalizedLocalPath = Path.GetFullPath(localAbsolutePath);

        if (!normalizedLocalPath.StartsWith(normalizedRoot, StringComparison.OrdinalIgnoreCase))
        {
            return normalizedLocalPath.Replace('\\', '/');
        }

        string relative = normalizedLocalPath.Substring(normalizedRoot.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        return GitLabPathResolver.NormalizeGitLabFilePathStrict(relative);
    }

    public void OnTokenManagerButtonPressed(IRibbonControl control)
    {
        Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
        try
        {
            GitLabTokenManagerDialog.ShowDialogSafe(null);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString(), "Token Manager Error");
        }
    }

}

public class AddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        if (!AuthorizationHelper.EnsureAuthorizedUser())
        {
            return;
        }

        Notifier.Initialize();
        //Notifier.Info("アドイン起動", "準備が完了しました。");

        // ここはExcelのUIスレッド。必ず最初に記録しておく
        ShellBridge.InitializeOnExcelUiThread();

        // アドインの実行ファイルの隣に scripts/ フォルダ置いてる前提
        var baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "scripts");
        JsHost.Init(baseDir);

        // 最初に共通ライブラリを全部チャージ
        JsHost.LoadModule(@"lib\lodash.min.js");
        JsHost.LoadModule(@"lib\js-yaml.min.js");

        JsHost.LoadModule(@"app\utilities.js");
        JsHost.LoadModule(@"app\constants.js");
        JsHost.LoadModule(@"app\CLCommon.js");
        JsHost.LoadModule(@"app\preprocess.js");
        JsHost.LoadModule(@"app\readconf.js");
        try
        {
            JsHost.LoadModule(@"app\txt2json.js");
        }
        catch (Microsoft.ClearScript.ScriptEngineException ex)
        {
            string details = ex.ErrorDetails;

            FileLogger.Error(ex.ToString());
            Notifier.Error("エラー", "パースでエラーが発生しました。クリックでログを開きます。");

            MessageBox.Show(details, "JS実行エラー");
        }

        RibbonController.CleanupOldPullWorkDirectories();
    }

    public void AutoClose()
    {
        //Notifier.Info("アドイン終了", "シャットダウンします。");
        Notifier.Dispose();
    }
}
