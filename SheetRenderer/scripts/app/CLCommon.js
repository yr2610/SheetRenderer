CL = {};

CL.kind = {
  H: "H",
  UL: "UL"
};

// json をテキストファイルに書き出すのを作ったけど、 stringify は別にやれば済む話なので単に文字列をテキストファイル化するだけな感じで
CL.writeTextFileUTF8 = function (s, outFilePath) {
  File.WriteAllText(outFilePath, s);
};

// 圧縮されてれば展開
CL.decompressJSON = function (json) {
  var o = JSON.parse(json);

  var count = 0;
  while (o.compress) {
    if (o.compress === "LZString") {
      var decompressor;
      if (o.option === "UTF16") {
        decompressor = LZString.decompressFromUTF16;
      }
      else if (o.option === "EncodedURIComponent") {
        decompressor = LZString.decompressFromEncodedURIComponent;
      }
      else {
        decompressor = LZString.decompress;
      }
      json = decompressor(o.data);
      o = JSON.parse(json);
    }
    else {
      Error("invalid compressor.");
      return;
    }
    if (count++ > 10) {
      Error("invalid JSON data.");
      return;
    }
  }

  return {
    json: json,
    object: o
  };
};

CL.readYAMLFile = function (yamlFilePath) {
  //var s = File.ReadAllText(yamlFilePath);
  //return jsyaml.load(s);
  return Yaml.LoadFile(yamlFilePath);
};

// fun は true を返せばそれ以降の traverse を打ち切る
CL.forAllNodes = function (node, parent, fun) {
  if (node === null) {
    return false;
  }
  if (fun(node, parent)) {
    return true;
  }

  for (var i = 0; i < node.children.length; i++) {
    if (CL.forAllNodes(node.children[i], node, fun)) {
      return true;
    }
  }

  return false;
};
CL.ForAllNodes = CL.forAllNodes;

// ULの各グループの幅を配列で返す
CL.getMaxItemWidth = function (node)
{
    var max = [];   // group 毎

    CL.forAllNodes(node, null, function(node) {
        if (node.kind !== CL.kind.UL) {
            return;
        }

        if (typeof max[node.group] === "undefined") {
            max[node.group] = node.depthInGroup + 1;
        }
        else {
            max[node.group] = Math.max(max[node.group], node.depthInGroup + 1);
        }
    });

    return max;
}

CL.getCheckHeaders = function (nodeH1, checkSheetTableData) {
  if (!nodeH1.tableHeaders) {
    return [ checkSheetTableData.input.header ];
  }
  return nodeH1.tableHeaders.map(function(x) {
    return x.name
  });
}


CL.getLeafNodes = function (node) {
  var leaves = [];
  CL.ForAllNodes(node, null, function (node, parent) {
    if (node.children.length === 0) {
      leaves.push(node);
    }
  });
  return leaves;
}
CL.GetLeafNodes = CL.getLeafNodes;

CL.nodeGetNumLeaves = function (node) {
  return CL.GetLeafNodes(node).length;
};
CL.NodeGetNumLeaves = CL.nodeGetNumLeaves;

CL.deletePropertyForAllNodes = function (node, propertyName) {
  CL.ForAllNodes(node, null, function (node, parent) {
    if (propertyName in node) {
      delete node[propertyName];
    }
  });
};
CL.DeletePropertyForAllNodes = CL.deletePropertyForAllNodes;

CL.addParentPropertyForAllNodes = function (node) {
  CL.ForAllNodes(node, null, function (node, parent) {
    node.parent = parent;
  });
};
CL.AddParentPropertyForAllNodes = CL.addParentPropertyForAllNodes;


// ID を基に node を取得
// leaf にしか ID はふられてないので、返る node は leaf になるはずだけど、 leaf 以外が返ったとしても特に問題ない作りのはず
// シート名を変更したい場合もあるはずなので、毎回すべてを検索するべき
// TODO: indexValues 用に level1 の H node 調べる用に maxDepth を渡せるようにしても良いか
CL.FindNodeById = function (node, id) {
  var resultNode = null;
  CL.ForAllNodes(node, null, function (node, parent) {
    if (node.id === id) {
      resultNode = node;
      // id はユニークという前提なので、１つ見つかった時点で終了して良い
      return true;
    }
  });
  return resultNode;
};

// 階層を考慮して id で検索
// idPath 通りの id の並び（idPath の末尾まで一致）の node を返す
// idPath には親から順に格納された配列を渡す
CL.FindNodeByIdPath = function (node, idPath) {
  if (idPath.length === 0) {
    return null;
  }

  var currentIdPath = [];

  function recurse(node) {
    if (node.id) {
      var i = currentIdPath.length;

      if (idPath[i] !== node.id) {
        return null;
      }
      // idPath の末尾まで一致してた
      // id path はユニークという前提なので、１つ見つかった時点で終了して良い
      if (i === idPath.length) {
        return node;
      }

      // push で idPath と同じ長さになる
      if (i + 1 >= idPath.length) {
        return null;
      }

      currentIdPath.push(node.id);
    }

    if (currentIdPath.length < idPath.length)

      for (var i = 0; i < node.children.length; i++) {
        var result = recurse(node.children[i]);

        if (result) {
          return result;
        }
      }

    if (node.id) {
      currentIdPath.pop();
    }

    return null;
  }

  return recurse(node);
};


CL.yyyymmddhhmmss = function (date) {
  // 1桁の数字を0埋めして2桁に
  function zeroPadding(value) {
    return ('0' + value).slice(-2);
    //return (value < 10) ? "0" + value : value;
  }
  var sa =
    [
      date.getFullYear(),
      zeroPadding(date.getMonth() + 1),
      zeroPadding(date.getDate()),
      zeroPadding(date.getHours()),
      zeroPadding(date.getMinutes()),
      zeroPadding(date.getSeconds())
    ];
  return sa.join("");
};

// フォルダが存在しなければ作成
// フォルダ名として作れないパスを渡された場合は無視
CL.createFolder = function (folderPath) {
  var fso = FileSystem;

  function recurse(folderPath) {
    var parentFolderPath = fso.GetParentFolderName(folderPath);
    // 少なくともここで対象としているフォルダはファイルが置かれている場所より下の階層なので、rootまで遡ってしまうことは考慮しなくていいけど、一応
    if (parentFolderPath !== "" && !fso.FolderExists(parentFolderPath)) {
      recurse(parentFolderPath);
    }

    if (!fso.FolderExists(folderPath)) {
      try {
        fso.CreateFolder(folderPath);
      } catch (e) {
      }
    }
  }

  recurse(folderPath);
}

// 指定したフォルダ（相対パス。なければ作る）にファイルを移動
CL.moveFile = function (filePath, relativeFolderPath) {
  var fso = FileSystem;
  var parentFolderPath = fso.GetParentFolderName(filePath);
  var dstFolderPath = fso.BuildPath(parentFolderPath, relativeFolderPath);
  var fileName = fso.GetFileName(filePath);
  var dstFilePath = fso.BuildPath(dstFolderPath, fileName);

  // なければ作る
  CL.createFolder(dstFolderPath);
  fso.MoveFile(filePath, dstFilePath);
};

// DateLastModified をつけたファイル名を生成
CL.makeBackupFileName = function (filePath) {
  var fso = FileSystem;
  let lastWriteTime = FileSystem.GetLastWriteTimeString(filePath);
  var lastModifiedDate = CL.yyyymmddhhmmss(new Date(lastWriteTime)).slice(2);
  var backupFileName = Path.GetFileNameWithoutExtension(filePath) + "-bak" + lastModifiedDate + "." + FileSystem.GetExtensionName(filePath);

  return backupFileName;
};

// ファイルのバックアップ作成
// 更新日時をファイル名に追加したような名前でコピーする
// filePath にディレクトリー付きのパスを渡してもファイル名だけ返す
CL.makeBackupFile = function (filePath, relativeFolderPath) {
  var fso = FileSystem;
  var backupFolderPath = fso.GetParentFolderName(filePath);
  if (typeof relativeFolderPath !== "undefined") {
    backupFolderPath = fso.BuildPath(backupFolderPath, relativeFolderPath);
    CL.createFolder(backupFolderPath);
  }
  var backupFileName = CL.makeBackupFileName(filePath, fso);
  var backupFilePath = fso.BuildPath(backupFolderPath, backupFileName);

  fso.CopyFile(filePath, backupFilePath);
};
CL.MakeBackupFile = CL.makeBackupFile;

// https://dobon.net/vb/dotnet/file/getabsolutepath.html#section4 をそのまま移植
CL.getRelativePath = function (basePath, absolutePath) {
  if (basePath == null || basePath.length == 0) {
      return absolutePath;
  }
  if (absolutePath == null || absolutePath.length == 0) {
      return "";
  }

  var fso = FileSystem;
  var directorySeparatorChar = "\\";
  var parentDirectoryString = ".." + directorySeparatorChar;

  basePath = _.trimEnd(basePath, directorySeparatorChar);

  basePath = fso.GetAbsolutePathName(basePath);
  absolutePath = fso.GetAbsolutePathName(absolutePath);

  //パスを"\"で分割する
  var basePathDirs = basePath.split(directorySeparatorChar);
  var absolutePathDirs = absolutePath.split(directorySeparatorChar);

  //基準パスと絶対パスで、先頭から共通する部分を探す
  var commonCount = 0;
  for (var i = 0;
      i < basePathDirs.length &&
      i < absolutePathDirs.length &&
      basePathDirs[i].toUpperCase() === absolutePathDirs[i].toUpperCase();
      i++) {
      //共通部分の数を覚えておく
      commonCount++;
  }

  //共通部分がない時
  if (commonCount == 0) {
      return absolutePath;
  }

  //共通部分以降の基準パスのフォルダの深さを取得する
  var baseOnlyCount = basePathDirs.length - commonCount;
  //その数だけ"..\"を付ける
  var buf = _.repeat(parentDirectoryString, baseOnlyCount);

  //共通部分以降の絶対パス部分を追加する
  buf += absolutePathDirs.slice(commonCount).join(directorySeparatorChar);

  return buf;
}


CL.createRandomId = function (len) {
  var c = "abcdefghijklmnopqrstuvwxyz";
  var s = c.charAt(Math.floor(Math.random() * c.length));
  c += "0123456789";
  var cl = c.length;

  for (var i = 1; i < len; i++) {
    s += c.charAt(Math.floor(Math.random() * cl));
  }

  return s;
};

CL.convertUt2Sn = function(unixTimeMillis){ // UNIX時間(ミリ秒)→シリアル値
  var COEFFICIENT = 24 * 60 * 60 * 1000; //日数とミリ秒を変換する係数

  var DATES_OFFSET = 70 * 365 + 17 + 1 + 1; //「1900/1/0」～「1970/1/1」 (日数)
  var MILLIS_DIFFERENCE = 9 * 60 * 60 * 1000; //UTCとJSTの時差 (ミリ秒)

  return (unixTimeMillis + MILLIS_DIFFERENCE) / COEFFICIENT + DATES_OFFSET;
}

CL.yyyymmddhhmmssExcelFormat = function (date) {
  // 1桁の数字を0埋めして2桁に
  function zeroPadding(value) {
    return ('0' + value).slice(-2);
    //return (value < 10) ? "0" + value : value;
  }

  var s = "{0}/{1}/{2} {3}:{4}".format(
    date.getFullYear(),
    zeroPadding(date.getMonth() + 1),
    zeroPadding(date.getDate()),
    zeroPadding(date.getHours()),
    zeroPadding(date.getMinutes())
  );
  
  return s;
};
