function parseOneLineComment(srclines) {
    var lines = [];

    _.forEach(srclines, function(lineObj) {
        var line = lineObj.line;

        // 1) 空白 + // だけの行は完全に破棄
        if (/^\s*\/\//.test(line)) {
            return;
        }

        // 2) 行中コメントは「空白に続く //」だけを対象
        //    先頭の空白も含めて保存するため、\s+ の先頭位置を取得
        var cppCommentIndex = line.search(/\s+\/\//);
        if (cppCommentIndex !== -1) {
            lineObj.comment = line.slice(cppCommentIndex);   // 例: " // note" / "\t// note"
            lineObj.line = line.slice(0, cppCommentIndex);   // 本体側からは空白ごと切り落とし
        }

        lines.push(lineObj);
    });

    return lines;
}

// それぞれ行頭、行末に書かれた <!-- と --> のみ対応
// 入れ子に対応
// C style コメントについてはごくごく簡易的なもの
// 本来 C style コメントは入れ子には対応してないけど、そこまでは対応しない
function parseMultilineComment(srcLines) {
    var lines = [];
    srcLines = new ArrayReader(srcLines);

    var ParseCommentError = function(errorMessage, lineObj) {
        this.errorMessage = errorMessage;
        this.lineObj = lineObj;
    };
    
    function _parseNest(beginRe, endRe, lineObj) {
        if (!beginRe.test(lineObj.line)) {
            return false;
        }

        var beginLine = lineObj;

        for (var commentDepth = 0;;) {
            var line = lineObj.line;

            if (beginRe.test(line)) {
                commentDepth++;
            }
            if (endRe.test(line)) {
                commentDepth--;
            }
            if (commentDepth == 0) {
                break;
            }
            if (srcLines.atEnd) {
                if (commentDepth > 0) {
                    var errorMessage = "コメントが閉じていません";
                    throw new ParseCommentError(errorMessage, beginLine);
                }
            }
            lineObj = srcLines.read();
        }
        return true;
    }

    while (!srcLines.atEnd) {
        var lineObj = srcLines.read();

        if (_parseNest(/^\s*<!--.*/, /.*-->\s*$/, lineObj)) {
            continue;
        }
        if (_parseNest(/^\s*\/\*.*/, /.*\*\/\s*$/, lineObj)) {
            continue;
        }

        lines.push(lineObj);
    }

    return lines;
}


// #define とか #if else endif 的なの
// コメント削除が適用済みのを渡す
// objs: 定義済みマクロ変数
// TODO: 最後の3個の引数（include の parse で使うやつ）を整理する
function preProcessConditionalCompile(lines, defines, currentProjectDirectoryFromRoot, filePathAbs, templateVariables) {
    var srcLines = new ArrayReader(lines);
    var dstLines = [];
    var objs = defines;
    var states = []; // 入れ子対応のためスタックにしておく

    function currentCondtion() {
        return states.every(s => s.cond);
    }

    function evalFormula(formula, objs) {
        // XXX: 改行は処理的には不要だけど、デバッグ中に頻繁に出力する都合上付けておく
        var s = "(function(){\n";

        // 宣言されてない変数名は false 扱い
        var ids = formula.trim().match(/([a-zA-Z_]\w*)/g);
        if (ids) {
            var undefs = _.difference(ids, _.keys(objs));
            undefs = _.difference(undefs, ['true', 'false']);
            _.forEach(undefs, function(name) {
                s += "var " + name + "=false;\n";
            });
        }

        for (var name in objs) {
            s += "var " + name + "=" + JSON.stringify(objs[name]) + ";\n";
        }

        s += "return(" + formula + ");})();";
        return eval(s);
    }

    // parseError にすると ParseError を上書きするようなので parseSharpError にしておく
    function parseSharpError(option, lineObj) {
        if (!currentCondtion()) {
            return;
        }
        var errorMessage = "@error";
        var text = _.trim(option);
        if (text != "") {
            errorMessage += " : '" + text + "'";
        }
        else {
            errorMessage += " が発生しました。";
        }
        throw new ParseError(errorMessage, lineObj);
    }

    function parseDefine(option, lineObj) {
        if (!currentCondtion()) {
            return;
        }

        var name = option.trim();
        if (!/^([a-zA-Z_]\w*)?$/.test(name)) {
            var errorMessage = "@define コマンドの文法が正しくありません。";
            throw new ParseError(errorMessage, lineObj);
        }

        // define の場合は set true 扱い
        parseSet(name + " = true", lineObj);
    }

    function isReservedName(name) {
        var reserved = [
            'break',
            'case',
            'catch',
            'continue',
            'debugger',
            'default',
            'delete',
            'do',
            'else',
            'finally',
            'for',
            'function',
            'if',
            'in',
            'instanceof',
            'new',
            'return',
            'switch',
            'this',
            'throw',
            'try',
            'typeof',
            'var',
            'void',
            'while',
            'with',

            'true',
            'false',
            'undefined',
            'null'
        ];
        return (_.indexOf(reserved, name) != -1);
    }

    function parseUndef(option, lineObj) {
        if (!currentCondtion()) {
            return;
        }

        var name = option.trim();
        if (!/^([a-zA-Z_]\w*)?$/.test(name)) {
            var errorMessage = "@undef コマンドの文法が正しくありません。";
            throw new ParseError(errorMessage, lineObj);
        }

        if (!isReservedName(name)) {
            delete objs[name];
        }

        // undef の場合は set false 扱い、にしようと思ったけど undef の後で define で redefine 扱いになるので
        // 素直に削除だけにしておく
    }

    function parseSet(option, lineObj) {
        if (!currentCondtion()) {
            return;
        }

        var optionMatch = option.trim().match(/^([a-zA-Z_]\w*)\s*=\s*(.+)?$/);
        if (optionMatch === null) {
            var errorMessage = "@set コマンドの文法が正しくありません。";
            throw new ParseError(errorMessage, lineObj);
        }
        var name = optionMatch[1];

        if (isReservedName(name)) {
            var errorMessage = "変数名に予約語が使われています。";
            throw new ParseError(errorMessage, lineObj);
        }
        if (name in objs) {
            var errorMessage = "変数 " + name + " はすでに定義されています。";
            throw new ParseError(errorMessage, lineObj);
        }
        var value = optionMatch[2];
        try {
            objs[name] = evalFormula(value, objs);
        }
        catch (e) {
            var errorMessage = '右辺の式 "' + value + '" が不正です。';
            throw new ParseError(errorMessage, lineObj);
        }
    }

    function parseCondition(cond, lineObj) {
        try {
            return evalFormula(cond, objs);
        }
        catch (e) {
            var errorMessage = "条件式が不正です。";
            throw new ParseError(errorMessage, lineObj);
        }
    }

    function parseIf(option, lineObj) {
        var state = {
            cond: false, // 今のフラグ
            elseApplied: false,
            condDisabled: true,  // これが立ってたらつねに false 扱い

            lineObj: lineObj
        };
        if (currentCondtion()) {
            var cond = parseCondition(option.trim(), lineObj);
            state.cond = cond;
            state.condDisabled = cond;
        }

        states.push(state);
    }
    function parseElif(option, lineObj) {
        // いきなり elif 出現エラー
        if (states.length === 0) {
            var errorMessage = "対応する if がありません。";
            throw new ParseError(errorMessage, lineObj);
        }

        var state = _.last(states);
        // 今の階層ですでに else が処理済みならエラー
        if (state.elseApplied) {
            var errorMessage = "elif が else の後に存在します。";
            throw new ParseError(errorMessage, lineObj);
        }

        if (state.condDisabled) {
            state.cond = false;
            return;
        }

        var cond = parseCondition(option.trim(), lineObj);
        if (cond) {
            state.cond = true;
            state.condDisabled = true;
        }
    }
    function parseElse(option, lineObj) {
        // いきなり else 出現エラー
        if (states.length === 0) {
            var errorMessage = "対応する if がありません。";
            throw new ParseError(errorMessage, lineObj);
        }

        var state = _.last(states);
        state.cond = !state.condDisabled;
        state.elseApplied = true;
    }
    function parseEnd(option, lineObj) {
        // いきなり end 出現エラー
        if (states.length === 0) {
            var errorMessage = "対応する if がありません。";
            throw new ParseError(errorMessage, lineObj);
        }
        states.pop();
    }
    function parseCommand(lineObj) {
        var commandMatch = lineObj.line.match(/^@([a-zA-Z]+)(.+)?$/);

        if (!commandMatch) {
            return null;
        }

        var command = commandMatch[1];
        var option = commandMatch[2];

        switch (command) {
            case 'define':
                parseDefine(option, lineObj);
                break;
            case 'undef':
                parseUndef(option, lineObj);
                break;
            case 'set':
                parseSet(option, lineObj);
                break;
            case 'if':
                parseIf(option, lineObj);
                break;
            case 'elif':
                parseElif(option, lineObj);
                break;
            case 'else':
                parseElse(option, lineObj);
                break;
            case 'end':
                parseEnd(option, lineObj);
                break;
            case 'error':
                parseSharpError(option, lineObj);
                break;
            default: {
                var errorMessage = "不明の@コマンドです。";
                throw new ParseError(errorMessage, lineObj);
                break;
            }
        }

        // bool で良いけどなんとなく lineObj を返しておく
        return lineObj;
    }
    function parseInclude(lineObj) {
        var includeMatch = lineObj.line.match(/^<<\[\s*(.+)\s*\]\s*(\((.+)?\))?$/);

        if (!includeMatch) {
            return null;
        }

        var includeFileString = includeMatch[1];
        var includeOptionString = includeMatch[3];

        try {
            var includeFileInfo = parseIncludeFilePath(includeFileString, currentProjectDirectoryFromRoot, filePathAbs, templateVariables);
        }
        catch (e) {
            throw new ParseError(e.errorMessage, lineObj);
        }

        try {
            var includeParam = parseIncludeParameters(includeOptionString, templateVariables);
        }
        catch(e) {
            var errorMessage = "include パラメータが不正です。";
            throw new ParseError(errorMessage, lineObj);
        }

        // include ファイルに渡す用
        // 上書きする（階層が深い方を優先）
        var localTemplateVariables = _.assign(templateVariables, includeParam);

        localTemplateVariables.$currentProjectDirectory = currentProjectDirectoryFromRoot;

        var includeProjectDirectoryFromRoot = includeFileInfo.projectDirectory;

        var path = includeFileInfo.filePath;
        var pathAbs = directoryLocalPathToAbsolutePath(path, includeProjectDirectoryFromRoot, sourceDirectoryName);

        if (File.ResolveAndEnsureLocalPath) {
            pathAbs = File.ResolveAndEnsureLocalPath(pathAbs, filePathAbs);
        }

        if (!FileSystem.FileExists(pathAbs)) {
            var sourceDirectory = FileSystem.BuildPath(includeFileInfo.projectDirectory, sourceDirectoryName);

            var errorMessage = "フォルダ\n" + sourceDirectory + "\nには\nファイル\n" + path + "\nが存在しません";
            throw new ParseError(errorMessage, lineObj);
        }

        return preProcess_Recurse(path, includeProjectDirectoryFromRoot, defines, localTemplateVariables);
    }

    try {

    while (!srcLines.atEnd) {
        var lineObj = srcLines.read();

        if (parseCommand(lineObj)) {
            continue;
        }

        if (!currentCondtion()) {
            continue;
        }

        var includeLines = parseInclude(lineObj);

        if (includeLines) {
            dstLines = dstLines.concat(includeLines);
            continue;
        }

        dstLines.push(lineObj);
    }

    if (states.length !== 0) {
        var state = _.last(states);
        var errorMessage = "@if が完結していません。@end が必要です。";
        throw new ParseError(errorMessage, state.lineObj);
    }

    }
    catch (e) {
        if (_.isUndefined(e.lineObj) || _.isUndefined(e.errorMessage)){
            throw e;
        }
        parseError(e);
    }

    return dstLines;
}

function replaceText(s, data) {
    // XXX: 一旦は正規表現で置換できる程度の仕様にしておく
    var rep_fn = function(m, k) {
        if (!(k in data)) {
            throw k;
        }
        return data[k];
    }

    return s.replace(/\{\{\=\s*([\w\$]+)\s*\}\}/g, rep_fn);
}

// filename.txt とだけ指定した場合は現在のプロジェクトの source 直下
// projectname:filename.txt と指定すると外部プロジェクトの source 直下
// 外部プロジェクトは root を最優先で検索。次に include path から検索（未対応）
// プロジェクト指定なしの場合 ./filename.txt と指定するとそのファイルからの相対
function parseIncludeFilePath(s, currentProjectPathFromRoot, currentFilePathAbs, variables) {

    var IncludeFilePathError = function(errorMessage) {
        this.errorMessage = errorMessage;
    };

    try {
        s = replaceText(s, variables);
    }
    catch (e) {
        var message = "'" + e + "' を置換できません";
        throw new IncludeFilePathError(message);
    }

    var includeMatch = s.match(/^((\/)?([^:]+):)?(\.\/)?(.+)$/);

    // 無効なパス指定
    if (!includeMatch) {
        var message = "無効なパスです";
        throw new IncludeFilePathError(message);
    }

    var localPath = includeMatch[5];
    var projectDirectoryFromRoot = includeMatch[3];
    // 現在のファイル(include元)からの相対指定
    var relativeFromCurrent = (includeMatch[4] !== undefined);

    if (relativeFromCurrent) {
        if (projectDirectoryFromRoot !== undefined) {
            var message = "外部参照では現在のファイルからの相対指定はできません";
            throw new IncludeFilePathError(message);
        }
        var currentFileDirectoryAbs = FileSystem.GetParentFolderName(currentFilePathAbs);
        var pathAbs = FileSystem.BuildPath(currentFileDirectoryAbs, localPath);
        var filePath = absolutePathToSourceLocalPath(pathAbs, currentProjectPathFromRoot);
        var result = {
            projectDirectory: currentProjectPathFromRoot,
            filePath: filePath
        };

        return result;
    }

    // XXX: 当面は root 以下専用
    if (projectDirectoryFromRoot === undefined) {
        projectDirectoryFromRoot = currentProjectPathFromRoot;
    }
    else {
        // root 指定の有無に関係なく root を優先して読む
        var fromRoot = (includeMatch[2] !== undefined);
        if (!fromRoot) {
            // TODO: root 指定ナシの場合は include path も検索する
        }
    }

    // source 直下からの相対
    var result = {
        projectDirectory: projectDirectoryFromRoot,
        filePath: localPath
    };

    return result;
}

// パラメータは文字列のみの想定
// 文字列以外を渡された場合の動作は不定
// variables は include 元で定義済みの変数
function parseIncludeParameters(s, variables) {
    if (s === undefined) {
        return {};
    }

    // object を返すには丸括弧が必要らしい
    var params = eval("({" + s + "})");

    // 各パラメータを template 処理
    _.forEach(params, function(value, name) {
        params[name] = replaceText(value, variables);
        // TODO: システム変数（$currentProject）の処理
        // TODO: $currentProject は root から現在の stack top への相対
    });

    return params;
}

function initPreProcessPathHelpers(rootDirectory) {
    absolutePathToSourceLocalPath = function(filePath, projectPathFromRoot) {
        var sourceDirectoryAbs = FileSystem.BuildPath(
            FileSystem.BuildPath(rootDirectory, projectPathFromRoot),
            sourceDirectoryName
        );
        return CL.getRelativePath(sourceDirectoryAbs, filePath);
    };

    directoryLocalPathToAbsolutePath = function(filePathProjectLocal, projectPathFromRoot, directoryName) {
        var projectPathAbs = FileSystem.BuildPath(rootDirectory, projectPathFromRoot);
        var directoryAbs = _.isUndefined(directoryName)
            ? projectPathAbs
            : FileSystem.BuildPath(projectPathAbs, directoryName);
        return FileSystem.BuildPath(directoryAbs, filePathProjectLocal);
    };
}

function getEntryConfigFileName(entryFilePath) {
    var confFileName = "conf.yml";
    var baseName = Path.GetFileNameWithoutExtension(entryFilePath);
    baseName = baseName.replace(/_index$/, "");
    if (baseName != "index") {
        confFileName = baseName + "_" + confFileName;
    }
    return confFileName;
}

function readConfigRootDirectoryOnly(confFilePath) {
    if (!FileSystem.FileExists(confFilePath)) {
        return null;
    }

    function readConfigPart(filePath) {
        return CL.withActiveReadFile(filePath, function() {
            var data = CL.readYAMLFile(filePath, filePath) || {};
            var baseDirectory = FileSystem.GetParentFolderName(filePath);
            var rootDirectory = null;

            if (!_.isUndefined(data.$rootDirectory)) {
                rootDirectory = FileSystem.BuildPath(baseDirectory, data.$rootDirectory);
            }

            if (!_.isUndefined(data.$include)) {
                _.forEach(data.$include, function(value) {
                    if (rootDirectory) {
                        return;
                    }

                    var includeFilePath = FileSystem.BuildPath(baseDirectory, value);
                    rootDirectory = readConfigPart(includeFilePath);
                });
            }

            return rootDirectory;
        });
    }

    return readConfigPart(confFilePath);
}

function getScaffoldRootDirectory(entryFilePath) {
    var entryProject = FileSystem.GetParentFolderName(entryFilePath);
    var confFilePath = FileSystem.BuildPath(entryProject, getEntryConfigFileName(entryFilePath));
    var rootDirectory = readConfigRootDirectoryOnly(confFilePath) || entryProject;
    return FileSystem.GetAbsolutePathName(rootDirectory);
}

function normalizePathKey(path) {
    return FileSystem.GetAbsolutePathName(path).replace(/\//g, "\\").toLowerCase();
}

function isPathUnderOrEqual(path, rootDirectory) {
    var normalizedPath = normalizePathKey(path).replace(/\\+$/, "");
    var normalizedRoot = normalizePathKey(rootDirectory).replace(/\\+$/, "");

    return normalizedPath === normalizedRoot
        || normalizedPath.indexOf(normalizedRoot + "\\") === 0;
}

function matchIncludeLine(line) {
    return line.match(/^<<\[\s*(.+)\s*\]\s*(\((.+)?\))?$/);
}

function matchTitleCommentLine(line) {
    return line.match(/^\s*\/\/(.*)$/);
}

function findTitleForInclude(lines, includeLineIndex) {
    for (var i = includeLineIndex - 1; i >= 0; i--) {
        var line = lines[i];

        if (/^\s*$/.test(line)) {
            continue;
        }

        var commentMatch = matchTitleCommentLine(line);
        if (commentMatch) {
            return {
                title: commentMatch[1].trim(),
                lineNum: i + 1
            };
        }

        return null;
    }

    return null;
}

function getInvalidExcelSheetNameReason(sheetName) {
    if (sheetName === "") {
        return "空のシート名は使用できません。";
    }
    if (sheetName.length > 31) {
        return "シート名は31文字以内にしてください。";
    }
    if (/[:\\\/\?\*\[\]]/.test(sheetName)) {
        return "シート名に : \\ / ? * [ ] は使用できません。";
    }
    if (/[\x00-\x1F]/.test(sheetName)) {
        return "シート名に制御文字は使用できません。";
    }
    if (sheetName.charAt(0) === "'" || sheetName.charAt(sheetName.length - 1) === "'") {
        return "シート名の先頭または末尾に ' は使用できません。";
    }

    return null;
}

function findScaffoldTemplateFile(targetDirectory, rootDirectoryAbs) {
    var directory = FileSystem.GetAbsolutePathName(targetDirectory);
    var rootDirectory = FileSystem.GetAbsolutePathName(rootDirectoryAbs);

    for (;;) {
        var templatePath = FileSystem.BuildPath(directory, "_template.txt");
        if (FileSystem.FileExists(templatePath)) {
            return templatePath;
        }

        if (normalizePathKey(directory) === normalizePathKey(rootDirectory)) {
            break;
        }
        if (!isPathUnderOrEqual(directory, rootDirectory)) {
            break;
        }

        var parent = FileSystem.GetParentFolderName(directory);
        if (!parent || normalizePathKey(parent) === normalizePathKey(directory)) {
            break;
        }
        directory = parent;
    }

    return null;
}

function renderScaffoldTemplate(templateText, values) {
    return templateText.replace(/\[\[\s*(title|name|path)\s*\]\]/gi, function(match, key) {
        key = key.toLowerCase();
        return values[key];
    });
}

function formatScaffoldLocation(entryFilePath, lineNum) {
    return "\nファイル:\t" + entryFilePath + "\n行:\t" + lineNum;
}

function collectMissingIncludeDraftPlans(entryFilePath) {
    var entryFilePathAbs = FileSystem.GetAbsolutePathName(entryFilePath);
    var entryProject = FileSystem.GetParentFolderName(entryFilePathAbs);
    var rootDirectory = getScaffoldRootDirectory(entryFilePathAbs);
    var projectPathFromRoot = CL.getRelativePath(rootDirectory, entryProject);
    var entrySourceLocalPath;
    var entryFileText;
    var lines;
    var errors = [];
    var plans = [];
    var plannedPathKeys = {};
    var plannedSheetNameKeys = {};

    initPreProcessPathHelpers(rootDirectory);
    entrySourceLocalPath = absolutePathToSourceLocalPath(entryFilePathAbs, projectPathFromRoot);
    entryFileText = File.ReadAllText(entryFilePathAbs);
    lines = entryFileText.replace(/\r\n|\r/g, "\n").split("\n");

    _.forEach(lines, function(line, lineIndex) {
        var includeMatch = matchIncludeLine(line);
        if (!includeMatch) {
            return;
        }

        var titleInfo = findTitleForInclude(lines, lineIndex);
        if (!titleInfo) {
            return;
        }

        var includeFileString = includeMatch[1];
        var includeFileInfo;
        try {
            includeFileInfo = parseIncludeFilePath(includeFileString, projectPathFromRoot, entryFilePathAbs, {});
        }
        catch (e) {
            errors.push(
                "include パスが不正です: " + includeFileString
                + "\n" + e.errorMessage
                + formatScaffoldLocation(entrySourceLocalPath, lineIndex + 1)
            );
            return;
        }

        var targetPath = directoryLocalPathToAbsolutePath(
            includeFileInfo.filePath,
            includeFileInfo.projectDirectory,
            sourceDirectoryName
        );
        targetPath = FileSystem.GetAbsolutePathName(targetPath);

        var projectDirectoryAbs = FileSystem.GetAbsolutePathName(
            FileSystem.BuildPath(rootDirectory, includeFileInfo.projectDirectory)
        );

        if (!isPathUnderOrEqual(targetPath, projectDirectoryAbs)) {
            errors.push(
                "下書きの作成先がプロジェクトフォルダの外です: " + includeFileInfo.filePath
                + formatScaffoldLocation(entrySourceLocalPath, lineIndex + 1)
            );
            return;
        }

        if (FileSystem.FileExists(targetPath)) {
            return;
        }

        var title = titleInfo.title;
        var invalidSheetNameReason = getInvalidExcelSheetNameReason(title);
        if (invalidSheetNameReason) {
            errors.push(
                "シート名として使用できないコメントです: " + title
                + "\n" + invalidSheetNameReason
                + formatScaffoldLocation(entrySourceLocalPath, titleInfo.lineNum)
            );
            return;
        }

        var sheetNameKey = title.toLowerCase();
        if (plannedSheetNameKeys[sheetNameKey]) {
            errors.push(
                "同じシート名の下書きを複数作成しようとしています: " + title
                + "\n既存候補: " + plannedSheetNameKeys[sheetNameKey]
                + "\n重複候補: " + includeFileInfo.filePath
                + formatScaffoldLocation(entrySourceLocalPath, titleInfo.lineNum)
            );
            return;
        }

        var targetPathKey = normalizePathKey(targetPath);
        if (plannedPathKeys[targetPathKey]) {
            errors.push(
                "同じ下書きファイルを複数回作成しようとしています: " + includeFileInfo.filePath
                + formatScaffoldLocation(entrySourceLocalPath, lineIndex + 1)
            );
            return;
        }

        var targetDirectory = FileSystem.GetParentFolderName(targetPath);
        var templatePath = findScaffoldTemplateFile(targetDirectory, rootDirectory);
        if (!templatePath) {
            errors.push(
                "_template.txt が見つかりません。"
                + "\n探索開始フォルダ: " + targetDirectory
                + "\n探索終了フォルダ: " + rootDirectory
                + formatScaffoldLocation(entrySourceLocalPath, lineIndex + 1)
            );
            return;
        }

        var templateText = File.ReadAllText(templatePath);
        var values = {
            title: title,
            name: Path.GetFileNameWithoutExtension(targetPath),
            path: includeFileInfo.filePath.replace(/\\/g, "/")
        };
        var outputText = renderScaffoldTemplate(templateText, values);

        plannedPathKeys[targetPathKey] = true;
        plannedSheetNameKeys[sheetNameKey] = includeFileInfo.filePath;
        plans.push({
            title: title,
            fileName: Path.GetFileName(targetPath),
            targetPath: targetPath,
            targetPathKey: targetPathKey,
            targetDirectory: targetDirectory,
            templatePath: templatePath,
            outputText: outputText
        });
    });

    if (errors.length > 0) {
        throw new Error(
            "下書きは作成されませんでした。"
            + "\n\n" + errors.join("\n\n")
        );
    }

    return {
        rootDirectory: rootDirectory,
        sourceDirectoryPath: FileSystem.BuildPath(rootDirectory, sourceDirectoryName),
        plans: plans
    };
}

function createMissingIncludeDraftPlans(plans) {
    if (plans.length === 0) {
        return "作成対象の下書きはありませんでした。";
    }

    var createdPaths = [];
    var createdFileNames = [];
    try {
        _.forEach(plans, function(plan) {
            if (FileSystem.FileExists(plan.targetPath)) {
                throw new Error("作成直前にファイルが存在しました: " + plan.targetPath);
            }

            CL.createFolder(plan.targetDirectory);
            if (File.WriteAllTextIfNotExists) {
                File.WriteAllTextIfNotExists(plan.targetPath, plan.outputText);
            }
            else {
                File.WriteAllText(plan.targetPath, plan.outputText);
            }
            createdPaths.push(plan.targetPath);
            createdFileNames.push(plan.fileName);
        });
    }
    catch (e) {
        _.forEach(createdPaths, function(path) {
            try {
                if (FileSystem.FileExists(path)) {
                    FileSystem.DeleteFile(path);
                }
            }
            catch (rollbackError) {
            }
        });
        throw new Error(
            "下書き作成に失敗したため、作成済みファイルをロールバックしました。"
            + "\n" + e.message
        );
    }

    return "下書きを " + createdPaths.length + " 件作成しました。"
        + "\n\n" + createdFileNames.join("\n");
}

function buildMissingIncludeDraftPlan(entryFilePath) {
    var result = collectMissingIncludeDraftPlans(entryFilePath);
    return JSON.stringify({
        sourceDirectoryPath: result.sourceDirectoryPath,
        plans: _.map(result.plans, function(plan) {
            return {
                title: plan.title,
                fileName: plan.fileName,
                targetPathKey: plan.targetPathKey
            };
        })
    });
}

function createMissingIncludeDraftsFromSelection(entryFilePath, selectedTargetPathKeysJson) {
    var selectedTargetPathKeys = JSON.parse(selectedTargetPathKeysJson || "[]");
    var selected = {};
    _.forEach(selectedTargetPathKeys, function(key) {
        selected[key] = true;
    });

    var result = collectMissingIncludeDraftPlans(entryFilePath);
    var plans = _.filter(result.plans, function(plan) {
        return selected[plan.targetPathKey];
    });

    return createMissingIncludeDraftPlans(plans);
}

function createMissingIncludeFilesFromEntry(entryFilePath) {
    var result = collectMissingIncludeDraftPlans(entryFilePath);
    return createMissingIncludeDraftPlans(result.plans);
}

// filePaths: 含まれるすべてのファイルのパス
function preProcess_Recurse(filePath, currentProjectDirectoryFromRoot, defines, templateVariables) {
    // 上書きする（階層が深い方を優先）
    templateVariables = _.assign(templateVariables, { $currentProjectDirectory: currentProjectDirectoryFromRoot });

    var filePathAbs = directoryLocalPathToAbsolutePath(filePath, currentProjectDirectoryFromRoot, sourceDirectoryName);

    return CL.withActiveReadFile(filePathAbs, function() {
        var allLines = CL.readTextFile(filePathAbs, filePathAbs);

        // 空要素も結果に含めたいのでsplitには正規表現を使わないように
        var lineArray = allLines.replace(/\r\n|\r/g, "\n").split("\n");

        // 最初に lineObj にしてしまう
        var lines = [];
        _.forEach(lineArray, function(line, lineNum) {
            if (line === "") {
                return;
            }

            var lineObj = {
                line: line,
                lineNum: 1 + lineNum,   // 1 origin
                filePath: filePath,
                projectDirectory: currentProjectDirectoryFromRoot
            };
            lines.push(lineObj);
        });

        lines = parseOneLineComment(lines);
        lines = parseMultilineComment(lines);
        lines = preProcessConditionalCompile(lines, defines, currentProjectDirectoryFromRoot, filePathAbs, templateVariables);

        return lines;
    });
}

// preprocess
// include とかコメント削除とか
// 入れ子の include にも対応
function preProcess(filePathAbs, rootDirectory) {
    var defines = {};

    // メインソースファイルのフォルダを現在のプロジェクトフォルダとする
    var entryProject = FileSystem.GetParentFolderName(filePathAbs);
    var projectPathFromRoot = CL.getRelativePath(rootDirectory, entryProject);
    var filePath = absolutePathToSourceLocalPath(filePathAbs, projectPathFromRoot);

    // TODO: conf.yaml とかで global な変数を指定できるように
    var templateVariables = { };

    try {
        return preProcess_Recurse(filePath, projectPathFromRoot, defines, templateVariables);
    }
    catch (e) {
        if (_.isUndefined(e.lineObj) || _.isUndefined(e.errorMessage)){
            throw e;
        }
        parseError(e);
    }
}
