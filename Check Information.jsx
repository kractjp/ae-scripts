// ================================================
// KRACT Check Information AutoLayer
// Version 2.35 (Settings File Extensionless)
// ================================================

(function (thisObj) {
  // --- グローバル設定 ---
  var SCRIPT_NAME = "Check Information AutoLayer";
  var SCRIPT_VERSION = "2.35_ext_less"; // バージョン更新
  var FONT_POSTSCRIPT_NAME = "GeistMono-Medium";
  var FONT_SIZE = 20;
  var PRECOMP_PREFIX = "_check-data_";
  var PRECOMP_FOLDER_NAME = "_check-data"; // プリコンポーズとロゴの親フォルダ
  var FOOTAGE_FOLDER_NAME = "footage"; // AEプロジェクト内のロゴフッテージを格納するサブフォルダ名
  var SETTINGS_DIR_NAME = ".settings"; // ★追加: 設定ディレクトリ名
  // var PROJECT_NAME_FILE_EXTENSION = ".txt"; // ★削除: プロジェクト名ファイルの拡張子
  var LOGO_MANAGE_BASE_FOLDER_NAME = "Check Information AutoLayer"; // 書類フォルダ内に作成するベースフォルダ名
  var LOGO_MANAGE_SUBFOLDER_NAME = "footage"; // ロゴを格納するサブフォルダ名
  var LOGO_TARGET_FILENAME = "checkinfo-logo.ai"; // 管理ディレクトリにコピーされるロゴファイル名
  var LOGO_FOOTAGE_NAME = "checkinfo-logo.ai"; // AEプロジェクト内のフッテージ名
  var LOGO_LAYER_NAME = "logo";
  var LOGO_POSITION = [960, 1131.75];
  var LOGO_SCALE = [16, 16];

  // ★★★ デバッグログファイル設定 ★★★
  var DEBUG_LOG_FILENAME = "CheckInfo_DebugLog.txt";
  var DEBUG_LOG_FILE_PATH = Folder.desktop.fsName + "/" + DEBUG_LOG_FILENAME; // デスクトップに保存

  // --- ログ書き込み関数 ---
  function logToFile(message) {
    try {
      var logFile = File(DEBUG_LOG_FILE_PATH);
      logFile.encoding = "UTF-8";
      // 追記モードでファイルを開く (なければ新規作成)
      if (logFile.open("a")) {
        var currentTime = new Date().toLocaleTimeString([], { hour12: false });
        logFile.writeln("[" + currentTime + "] " + message);
        logFile.close();
      } else {
        // alert が使えるなら通知、使えない場合は仕方ない
        try {
          alert(
            "デバッグログファイルを開けません(a): " +
              logFile.error +
              "\nPath: " +
              DEBUG_LOG_FILE_PATH
          );
        } catch (e) {}
      }
    } catch (e) {
      // alert が使えるなら通知、使えない場合は仕方ない
      try {
        alert(
          "デバッグログ書き込み中にエラー: " +
            e.toString() +
            "\nPath: " +
            DEBUG_LOG_FILE_PATH
        );
      } catch (err) {}
    }
  }

  // --- writeLn 関数を修正 ---
  function writeLn(message) {
    // ファイルに書き込む
    logToFile(message);
  }

  // スクリプト開始時にログファイルに区切りを入れる
  function initializeLogFile() {
    var separator =
      "\n====================\nScript Execution Started: " +
      SCRIPT_NAME +
      " " +
      SCRIPT_VERSION +
      " at " +
      new Date().toString() +
      "\n====================\n";
    try {
      var logFile = File(DEBUG_LOG_FILE_PATH);
      logFile.encoding = "UTF-8";
      // 追記モードでファイルを開き、区切りを書き込む
      if (logFile.open("a")) {
        logFile.writeln(separator);
        logFile.close();
      } else {
        try {
          alert(
            "デバッグログファイル初期化失敗(a): " +
              logFile.error +
              "\nPath: " +
              DEBUG_LOG_FILE_PATH
          );
        } catch (e) {}
      }
    } catch (e) {
      try {
        alert(
          "デバッグログファイル初期化エラー: " +
            e.toString() +
            "\nPath: " +
            DEBUG_LOG_FILE_PATH
        );
      } catch (err) {}
    }
  }

  // --- 書類フォルダ内のロゴ管理パスを定義 ---
  var managedLogoDir = Folder(
    Folder.myDocuments.fsName +
      "/" +
      LOGO_MANAGE_BASE_FOLDER_NAME +
      "/" +
      LOGO_MANAGE_SUBFOLDER_NAME
  );
  var managedLogoFile = File(
    managedLogoDir.fsName + "/" + LOGO_TARGET_FILENAME
  );

  var LAYER_SPECS = {
    timecode: {
      pos: [14, 18.5275],
      ap: [0.4633, -7.1503],
      expression: '"REC TC/Fm: " + timeToCurrentFormat(time)',
      justification: ParagraphJustification.LEFT_JUSTIFY,
    },
    frame_num: {
      pos: [283, 18.5275],
      ap: [0.4633, -7.1503],
      expression: '"/" + ("00000" + Math.max(0, timeToFrames(time))).slice(-5)',
      justification: ParagraphJustification.LEFT_JUSTIFY,
    },
    comp_name: {
      pos: [960.1207, 18.5275],
      ap: [0.2238, -7.1503],
      justification: ParagraphJustification.CENTER_JUSTIFY,
    },
    comp_info: {
      pos: [1905.3997, 18.5275],
      ap: [0.8504, -7.1503],
      justification: ParagraphJustification.RIGHT_JUSTIFY,
    },
    current_info: {
      pos: [14, 1132.83],
      ap: [0.2238, -7.1503],
      justification: ParagraphJustification.LEFT_JUSTIFY,
    },
    project_name: {
      pos: [1905.3997, 1132.83],
      ap: [0.7238, -7.1503],
      justification: ParagraphJustification.RIGHT_JUSTIFY,
    },
  };

  // --- UI構築 ---
  function buildUI(thisObj) {
    var myPanel =
      thisObj instanceof Panel
        ? thisObj
        : new Window("palette", SCRIPT_NAME, undefined, { resizeable: true });
    myPanel.text = SCRIPT_NAME + " " + SCRIPT_VERSION;
    myPanel.orientation = "column";
    myPanel.alignChildren = ["fill", "top"];
    myPanel.spacing = 10;
    myPanel.margins = 15;

    var panelSource = myPanel.add("panel", undefined, "ソース選択");
    panelSource.orientation = "column";
    panelSource.alignChildren = "fill";
    panelSource.spacing = 5;
    panelSource.margins = 10;
    var ddSource = panelSource.add("dropdownlist", undefined, []);
    ddSource.helpTip = "情報を表示する基準となるソースレイヤーを選択";
    ddSource.minimumSize.width = 280;

    var panelProject = myPanel.add("panel", undefined, "プロジェクト情報");
    panelProject.orientation = "column";
    panelProject.alignChildren = "left";
    panelProject.spacing = 5;
    panelProject.margins = 10;
    var stProjectName = panelProject.add(
      "statictext",
      undefined,
      "プロジェクト表示名 (未入力時はファイル名):"
    );
    var editProjectName = panelProject.add("edittext", [0, 0, 280, 20], "");
    editProjectName.helpTip =
      "表示するプロジェクト名。設定があれば自動入力。変更は実行時に保存されます。";

    var panelLogo = myPanel.add("panel", undefined, "ロゴ設定");
    panelLogo.orientation = "column";
    panelLogo.alignChildren = "fill";
    panelLogo.spacing = 5;
    panelLogo.margins = 10;
    panelLogo.add(
      "statictext",
      undefined,
      "使用するロゴ (書類フォルダ内で管理):"
    );
    var grpLogoPath = panelLogo.add("group");
    grpLogoPath.orientation = "row";
    grpLogoPath.alignChildren = ["fill", "center"];
    var stLogoPath = grpLogoPath.add("statictext", [0, 0, 200, 20], "未設定", {
      truncate: "middle",
    });
    stLogoPath.helpTip =
      "管理下のロゴファイル (" + managedLogoFile.fsName + ") の状態";
    var btnBrowseLogo = grpLogoPath.add("button", undefined, "参照して設定...");
    btnBrowseLogo.helpTip =
      "ロゴファイル (AI形式) を選択し、管理フォルダにコピーします";

    // ロゴパス表示を更新する関数
    function updateLogoPathDisplay() {
      if (managedLogoFile.exists) {
        stLogoPath.text = managedLogoFile.name + " (設定済)";
        stLogoPath.helpTip = "管理下のロゴ: " + managedLogoFile.fsName;
      } else {
        stLogoPath.text = "未設定 (管理フォルダにロゴ無し)";
        stLogoPath.helpTip =
          "管理フォルダにロゴがありません: " + managedLogoFile.fsName;
      }
    }

    btnBrowseLogo.onClick = function () {
      var selectedFile = File.openDialog(
        "ロゴファイルを選択 (.ai)",
        "Adobe Illustrator:*.ai",
        false
      );
      if (selectedFile) {
        try {
          var copiedPath = setAndCopyLogo(selectedFile);
          if (copiedPath) {
            updateLogoPathDisplay(); // 表示を更新
            try {
              alert(
                "ロゴを管理フォルダにコピーしました:\n" + copiedPath,
                SCRIPT_NAME
              );
            } catch (e) {
              writeLn("ALERT FAILED: ロゴコピー完了");
            }
          }
        } catch (e) {
          try {
            alert(
              "ロゴの設定またはコピー中にエラーが発生しました:\n" +
                e.toString(),
              SCRIPT_NAME
            );
          } catch (err) {
            writeLn("ALERT FAILED: ロゴコピーエラー: " + e.toString());
          }
        }
      }
    };

    var grpBtns = myPanel.add("group");
    grpBtns.orientation = "row";
    grpBtns.alignment = "center";
    grpBtns.spacing = 10;
    var btnRefresh = grpBtns.add("button", undefined, "更新");
    btnRefresh.helpTip = "ソース一覧を更新します";
    var btnApply = grpBtns.add("button", undefined, "実行");
    btnApply.helpTip =
      "テキスト/ロゴ生成・設定、設定保存、プリコンポーズ、ソース再配置";

    var avSources = [];

    btnRefresh.onClick = function () {
      avSources = refreshSourceList(ddSource);
    };

    btnApply.onClick = function () {
      writeLn("--- btnApply onClick START ---"); // ★Debug Log
      if (!hasSavedProject()) {
        try {
          alert("プロジェクトが保存されていません。", SCRIPT_NAME);
        } catch (e) {
          writeLn("ALERT FAILED: プロジェクト未保存");
        }
        writeLn("btnApply onClick: Project not saved."); // ★Debug Log
        return;
      }
      var currentAvSources = getAvailableSources(getActiveComp()); // 最新のソースリストを取得
      writeLn(
        "btnApply onClick: Got currentAvSources: " +
          (currentAvSources ? currentAvSources.length : 0) +
          " items"
      ); // ★Debug Log

      // ★追加: 実行前にロゴ表示のみ更新
      updateLogoPathDisplay();

      var ac = getActiveComp();
      if (!ac) {
        try {
          alert("アクティブなコンポジションがありません。", SCRIPT_NAME);
        } catch (e) {
          writeLn("ALERT FAILED: アクティブコンポなし");
        }
        writeLn("btnApply onClick: No active comp."); // ★Debug Log
        return;
      }
      // ロゴファイルの存在確認 (管理パス)
      if (!managedLogoFile.exists) {
        writeLn(
          "btnApply onClick: Managed logo file not found: " +
            managedLogoFile.fsName
        ); // ★Debug Log
        try {
          alert(
            "管理フォルダにロゴファイルが存在しません。\nロゴ設定パネルでロゴファイルを設定してください。\nパス: " +
              managedLogoFile.fsName,
            SCRIPT_NAME
          );
        } catch (e) {
          writeLn("ALERT FAILED: 管理ロゴなし: " + managedLogoFile.fsName);
        }
        return;
      }
      if (
        ddSource.items.length === 0 ||
        ddSource.selection == null ||
        ddSource.selection.text === "ソースなし" ||
        ddSource.selection.text === "プロジェクト未保存 または コンポなし" // 初期化時のメッセージもチェック
      ) {
        writeLn("btnApply onClick: Invalid source selection in dropdown."); // ★Debug Log
        try {
          alert("有効なソースが選択されていません。", SCRIPT_NAME);
        } catch (e) {
          writeLn("ALERT FAILED: 有効ソース未選択");
        }
        return;
      }

      // ソース選択の有効性チェック
      var selIndex = ddSource.selection.index;
      if (
        !currentAvSources ||
        selIndex >= currentAvSources.length ||
        currentAvSources[selIndex].name !== ddSource.selection.text
      ) {
        writeLn("btnApply onClick: Source selection mismatch or outdated."); // ★Debug Log
        try {
          alert(
            "ソースリストが古いか、選択が無効です。\n更新ボタンを押して再試行してください。",
            SCRIPT_NAME
          );
        } catch (e) {
          writeLn("ALERT FAILED: ソースリスト古い/無効");
        }
        return;
      }
      var pickedObj = currentAvSources[selIndex];
      writeLn(
        "btnApply onClick: Picked source: " +
          (pickedObj ? pickedObj.name : "null")
      ); // ★Debug Log

      if (!pickedObj || !pickedObj.source) {
        writeLn("btnApply onClick: Failed to get source info from pickedObj."); // ★Debug Log
        try {
          alert("選択されたソース情報の取得に失敗しました。", SCRIPT_NAME);
        } catch (e) {
          writeLn("ALERT FAILED: ソース情報取得失敗");
        }
        return;
      }

      // ★修正: 保存するプロジェクト名 (UIの現在の値を取得)
      var projectNameToSave = editProjectName.text; // UIのedittextから直接取得
      writeLn(
        "btnApply onClick: Project name from UI to save: " + projectNameToSave
      ); // ★Debug Log

      // 表示用プロジェクト名決定 (UIの入力値を優先)
      var projectFileNameForLayer =
        projectNameToSave !== "" ? projectNameToSave : getAEProjectIdentifier(); // 未入力時はファイル名
      writeLn(
        "btnApply onClick: projectFileNameForLayer for display: " +
          projectFileNameForLayer
      ); // ★Debug Log

      // ★修正: 現在のプロジェクトの設定を保存 (新しい方式)
      try {
        writeLn("btnApply onClick: Calling saveProjectNameToSettingsFile..."); // ★Debug Log
        // 表示用に使った名前ではなく、UIに入力された名前（空文字列の可能性もある）を保存する
        saveProjectNameToSettingsFile(projectNameToSave);
        writeLn("プロジェクト設定 (.settings/[projectName]) を保存しました。");
      } catch (e) {
        writeLn("btnApply onClick: Error saving project name: " + e.toString()); // ★Debug Log
        try {
          alert("設定の保存に失敗しました:\n" + e.toString(), SCRIPT_NAME);
        } catch (err) {
          writeLn("ALERT FAILED: 設定保存失敗: " + e.toString());
        }
        // 保存失敗しても処理を続ける
      }

      writeLn(
        "btnApply onClick: Starting Undo Group and processComposition..."
      ); // ★Debug Log
      app.beginUndoGroup(SCRIPT_NAME + ": 実行");
      try {
        // managedLogoPath は引数から削除済
        processComposition(ac, pickedObj, projectFileNameForLayer);
        writeLn("btnApply onClick: processComposition completed."); // ★Debug Log
        try {
          alert("処理が完了しました。", SCRIPT_NAME);
        } catch (e) {
          writeLn("ALERT FAILED: 処理完了");
        }
      } catch (e) {
        writeLn(
          "btnApply onClick: Error during processComposition: " +
            e.toString() +
            (e.line ? " Line: " + e.line : "")
        ); // ★Debug Log
        try {
          alert(
            "エラーが発生しました:\n" +
              e.toString() +
              (e.line ? "\nLine: " + e.line : ""),
            SCRIPT_NAME
          );
        } catch (err) {
          writeLn(
            "ALERT FAILED: 処理中エラー: " +
              e.toString() +
              (e.line ? " Line: " + e.line : "")
          );
        }
      } finally {
        app.endUndoGroup();
        writeLn("--- btnApply onClick END ---"); // ★Debug Log
      }
    };

    function initializePanel() {
      writeLn("--- initializePanel START ---"); // ★Debug Log
      if (hasSavedProject()) {
        writeLn("initializePanel: Project is saved."); // ★Debug Log
        try {
          // ★修正: 新しい方式でプロジェクト名を読み込み
          var loadedProjectName = loadProjectNameFromSettingsFile();
          writeLn("initializePanel: Loaded project name: " + loadedProjectName); // ★Debug Log
          editProjectName.text = loadedProjectName || ""; // UIに直接ロード
        } catch (e) {
          writeLn(
            "initializePanel: Error loading project name: " + e.toString()
          ); // ★Debug Log
          try {
            alert(
              "プロジェクト名の読み込み中にエラーが発生しました:\n" +
                e.toString(),
              SCRIPT_NAME
            );
          } catch (err) {
            writeLn(
              "ALERT FAILED: プロジェクト名読み込みエラー: " + e.toString()
            );
          }
        }
        avSources = refreshSourceList(ddSource);
        btnBrowseLogo.enabled = true; // プロジェクトが保存されていればロゴ参照は可能
      } else {
        writeLn(
          "initializePanel: Project is NOT saved. Disabling UI elements."
        ); // ★Debug Log
        // プロジェクト未保存時の初期状態
        editProjectName.text = "";
        editProjectName.enabled = false;
        btnBrowseLogo.enabled = false; // ロゴ参照も不可とする（設定保存場所がないため）
        ddSource.removeAll();
        ddSource.add("item", "プロジェクト未保存 または コンポなし");
        ddSource.selection = 0;
        ddSource.enabled = false;
        btnApply.enabled = false;
        btnRefresh.enabled = false;
      }
      updateLogoPathDisplay(); // ロゴの状態を表示
      writeLn("--- initializePanel END ---"); // ★Debug Log
    }

    function getAvailableSources(comp) {
      var sources = [];
      if (!comp) return sources;
      for (var i = 1; i <= comp.numLayers; i++) {
        var lyr = comp.layer(i);
        // AVLayer かつ 有効な source を持つレイヤーを対象とする
        // かつ、それがプリコンポーズ(_check-data_*) や ロゴフッテージでないこと
        if (
          lyr instanceof AVLayer &&
          lyr.source &&
          lyr.source !== null &&
          lyr.source.name !== LOGO_FOOTAGE_NAME &&
          !(
            lyr.source instanceof CompItem &&
            lyr.source.name.indexOf(PRECOMP_PREFIX) === 0
          )
        ) {
          var s = lyr.source;
          // FootageItem の場合は isStill でないこと（動画・シーケンスのみ）
          // CompItem の場合はそのまま追加
          if (s instanceof CompItem) {
            sources.push({ source: s, type: "comp", name: s.name });
          } else if (s instanceof FootageItem) {
            // FootageItemの場合、mainSourceが存在し、かつisStillでないことを確認
            if (s.mainSource && !s.mainSource.isStill) {
              sources.push({ source: s, type: "footage", name: s.name });
            }
          }
        }
      }
      return sources;
    }

    function refreshSourceList(dropdown) {
      writeLn("--- refreshSourceList START ---"); // ★Debug Log
      dropdown.removeAll();
      var currentSources = [];
      var ac = getActiveComp();
      var projectSaved = hasSavedProject();
      writeLn(
        "refreshSourceList: projectSaved=" +
          projectSaved +
          ", activeComp=" +
          (ac ? ac.name : "null")
      ); // ★Debug Log

      // UI要素の有効/無効状態設定
      var enableUI = projectSaved && ac;
      ddSource.enabled = enableUI;
      btnApply.enabled = enableUI;
      editProjectName.enabled = projectSaved; // プロジェクト名入力はプロジェクト保存が条件
      btnBrowseLogo.enabled = true; // ロゴ参照はプロジェクト保存状態に依存しない
      btnRefresh.enabled = projectSaved; // 更新ボタンもプロジェクト保存が条件

      if (!projectSaved) {
        dropdown.add("item", "プロジェクト未保存");
        dropdown.selection = 0;
        btnApply.enabled = false; // 実行不可
        writeLn("refreshSourceList: Project not saved state."); // ★Debug Log
      } else if (!ac) {
        dropdown.add("item", "アクティブなコンポなし");
        dropdown.selection = 0;
        btnApply.enabled = false; // 実行不可
        writeLn("refreshSourceList: No active comp state."); // ★Debug Log
      } else {
        currentSources = getAvailableSources(ac);
        writeLn(
          "refreshSourceList: Found " +
            currentSources.length +
            " available sources."
        ); // ★Debug Log
        if (currentSources.length === 0) {
          dropdown.add("item", "有効なソースレイヤーなし");
          dropdown.selection = 0;
          btnApply.enabled = false; // 実行不可
          writeLn("refreshSourceList: No valid source layers found state."); // ★Debug Log
        } else {
          for (var k = 0; k < currentSources.length; k++) {
            dropdown.add("item", currentSources[k].name);
          }
          dropdown.selection = 0;
          btnApply.enabled = true; // 実行可能
          writeLn("refreshSourceList: Populated dropdown. Selection index 0."); // ★Debug Log
        }
      }
      writeLn("--- refreshSourceList END ---"); // ★Debug Log
      return currentSources;
    }

    initializePanel();
    myPanel.layout.layout(true);
    myPanel.onResizing = myPanel.onResize = function () {
      this.layout.resize();
    };
    if (myPanel instanceof Window) {
      myPanel.center();
      myPanel.show();
    }
    return myPanel;
  }

  // ロゴファイルを指定された管理フォルダ（書類内）にコピーする関数
  function setAndCopyLogo(sourceFile) {
    if (!sourceFile || !sourceFile.exists) {
      throw new Error("指定されたロゴファイルが見つかりません。");
    }
    if (!/\.ai$/i.test(sourceFile.name)) {
      throw new Error(
        "ロゴファイルは Adobe Illustrator (.ai) 形式である必要があります。"
      );
    }

    // 管理用ベースフォルダ、サブフォルダの存在確認と作成
    var baseDir = Folder(
      Folder.myDocuments.fsName + "/" + LOGO_MANAGE_BASE_FOLDER_NAME
    );
    if (!baseDir.exists) {
      if (!baseDir.create()) {
        throw new Error(
          "書類フォルダ内にベースディレクトリを作成できませんでした: " +
            baseDir.fsName
        );
      }
      writeLn("ロゴ管理用ベースディレクトリを作成しました: " + baseDir.fsName);
    }

    var targetDir = Folder(baseDir.fsName + "/" + LOGO_MANAGE_SUBFOLDER_NAME);
    if (!targetDir.exists) {
      if (!targetDir.create()) {
        throw new Error(
          "ロゴ管理用サブディレクトリを作成できませんでした: " +
            targetDir.fsName
        );
      }
      writeLn("ロゴ管理用サブディレクトリを作成しました: " + targetDir.fsName);
    }

    var targetFile = File(targetDir.fsName + "/" + LOGO_TARGET_FILENAME);

    // 既存ファイルがあっても上書きコピーする
    if (sourceFile.copy(targetFile.fsName)) {
      writeLn(
        "ロゴファイルを管理フォルダにコピーしました: " +
          sourceFile.fsName +
          " -> " +
          targetFile.fsName
      );
      return targetFile.fsName; // コピー後のパスを返す
    } else {
      throw new Error(
        "ロゴファイルのコピーに失敗しました。\nコピー元: " +
          sourceFile.fsName +
          "\nコピー先: " +
          targetFile.fsName +
          "\nError: " +
          sourceFile.error
      );
    }
  }

  // --- メイン処理関数 ---
  // managedLogoPath 引数を削除し、グローバルの managedLogoFile を使用
  function processComposition(activeComp, pickedObj, projectFileName) {
    writeLn("--- processComposition START ---"); // ★Debug Log
    if (!activeComp) {
      writeLn("processComposition: No activeComp. Returning."); // ★Debug Log
      return;
    }
    if (!pickedObj || !pickedObj.source) {
      writeLn("processComposition: Invalid pickedObj or source."); // ★Debug Log
      throw new Error("処理に必要なソース情報がありません。");
    }
    // 管理下のロゴファイルの存在を再確認
    if (!managedLogoFile.exists) {
      writeLn(
        "processComposition: Managed logo file does not exist: " +
          managedLogoFile.fsName
      ); // ★Debug Log
      throw new Error(
        "管理下のロゴファイルが存在しません:\n" + managedLogoFile.fsName
      );
    }
    writeLn(
      "processComposition: Processing comp '" +
        activeComp.name +
        "' with source '" +
        pickedObj.name +
        "' and projectFileName '" +
        projectFileName +
        "'"
    ); // ★Debug Log

    var sourceItem = pickedObj.source;
    var sourceLayerInfo = { index: -1, name: "" };

    // --- 0. 既存レイヤーの情報を保持し、削除 ---
    writeLn("--- 既存レイヤー検索・削除開始 ---");
    var foundAndRemovedSource = false;
    for (var i = activeComp.numLayers; i >= 1; i--) {
      var currentLayer = activeComp.layer(i);
      var isManagedLayer = false;
      // 既存のテキストレイヤー、ロゴレイヤー、またはプリコンプレイヤーかチェック
      if (
        LAYER_SPECS[currentLayer.name] ||
        currentLayer.name === LOGO_LAYER_NAME ||
        (currentLayer.source &&
          currentLayer.source instanceof CompItem &&
          currentLayer.source.name.indexOf(PRECOMP_PREFIX) === 0)
      ) {
        isManagedLayer = true;
      }

      if (isManagedLayer) {
        try {
          writeLn(
            "  削除対象レイヤー発見: " +
              currentLayer.name +
              " (Index: " +
              currentLayer.index +
              ")"
          );
          if (currentLayer.locked) currentLayer.locked = false;
          currentLayer.remove();
          writeLn("  レイヤー削除完了。");
        } catch (e) {
          writeLn(
            "警告: 既存レイヤー '" +
              currentLayer.name +
              "' の削除失敗。\n" +
              e.toString()
          );
        }
      } else if (
        // 管理対象外 かつ 今回選択されたソースと同じAVLayerかチェック
        currentLayer instanceof AVLayer &&
        currentLayer.source &&
        currentLayer.source === sourceItem
      ) {
        try {
          writeLn(
            "  削除対象ソースレイヤー発見: " +
              currentLayer.name +
              " (Index: " +
              currentLayer.index +
              ")"
          );
          // 最初に見つかったソースレイヤーの情報を保持
          if (!foundAndRemovedSource) {
            sourceLayerInfo.index = currentLayer.index; // 位置復元は行わないが、名前は保持
            sourceLayerInfo.name = currentLayer.name;
          }
          if (currentLayer.locked) currentLayer.locked = false;
          currentLayer.remove();
          writeLn("  ソースレイヤー削除完了。");
          foundAndRemovedSource = true;
        } catch (e) {
          writeLn(
            "警告: 既存ソースレイヤー '" +
              currentLayer.name +
              "' の削除失敗。\n" +
              e.toString()
          );
        }
      }
    }
    if (!foundAndRemovedSource)
      writeLn("今回選択されたソースと同じ既存レイヤーは見つかりませんでした。");
    writeLn("--- 既存レイヤー検索・削除終了 ---");

    // --- 0.5 コンポジションサイズ変更 ---
    if (activeComp.width === 1920 && activeComp.height === 1080) {
      activeComp.height = 1150;
      writeLn("コンポジションの高さを 1150px に変更しました。");
    } else if (activeComp.width !== 1920 || activeComp.height !== 1150) {
      writeLn(
        "警告: コンポジションサイズが予期される値 (1920x1080 または 1920x1150) ではありません (" +
          activeComp.width +
          "x" +
          activeComp.height +
          "). サイズ変更はスキップします。"
      );
    }

    // --- 1. ロゴフッテージの準備 (インポートまたは置換) とフォルダ移動 ---
    writeLn("--- ロゴフッテージ準備開始 ---");
    var logoFootage = null;

    try {
      var existingLogoFootage = findItemByName(LOGO_FOOTAGE_NAME);

      if (existingLogoFootage && existingLogoFootage instanceof FootageItem) {
        writeLn(
          "既存のロゴフッテージ '" +
            LOGO_FOOTAGE_NAME +
            "' を発見。ソースを置換します。"
        );
        existingLogoFootage.replaceSource(managedLogoFile, false); // managedLogoFile を使用
        logoFootage = existingLogoFootage;
        writeLn(
          "  フッテージソースを管理下のロゴで置換しました: " +
            managedLogoFile.fsName
        );
      } else {
        writeLn(
          "ロゴフッテージ '" +
            LOGO_FOOTAGE_NAME +
            "' が見つからないため、新規にインポートします。"
        );
        var importOptions = new ImportOptions(managedLogoFile); // managedLogoFile を使用
        logoFootage = app.project.importFile(importOptions);
        if (!logoFootage) {
          throw new Error(
            "管理下のロゴファイル (" +
              managedLogoFile.fsName +
              ") のインポートに失敗しました。"
          );
        }
        logoFootage.name = LOGO_FOOTAGE_NAME;
        writeLn(
          "  ロゴをインポートし、名前を '" +
            LOGO_FOOTAGE_NAME +
            "' に設定しました。"
        );
      }
    } catch (e) {
      writeLn(
        "processComposition: Error during logo footage prep: " + e.toString()
      ); // ★Debug Log
      throw new Error(
        "ロゴフッテージの準備（インポート/置換）中にエラーが発生しました:\n" +
          e.toString()
      );
    }

    if (!logoFootage || !(logoFootage instanceof FootageItem)) {
      writeLn("processComposition: Failed to get valid logoFootage item."); // ★Debug Log
      throw new Error(
        "ロゴフッテージの準備に失敗しました。有効なフッテージアイテムを取得できませんでした。"
      );
    }

    // ロゴフッテージを AE プロジェクト内の _check-data/footage フォルダに移動
    try {
      var checkDataFolder = findOrCreateFolder(PRECOMP_FOLDER_NAME);
      if (!checkDataFolder) {
        writeLn(
          "processComposition: Failed to find/create checkDataFolder: " +
            PRECOMP_FOLDER_NAME
        ); // ★Debug Log
        throw new Error(
          "AEプロジェクトフォルダ '" +
            PRECOMP_FOLDER_NAME +
            "' の取得/作成に失敗。"
        );
      }

      var footageFolderInAE = findOrCreateFolder(
        FOOTAGE_FOLDER_NAME,
        checkDataFolder
      );
      if (!footageFolderInAE) {
        writeLn(
          "processComposition: Failed to find/create footageFolderInAE: " +
            FOOTAGE_FOLDER_NAME
        ); // ★Debug Log
        throw new Error(
          "AEプロジェクトフォルダ '" +
            FOOTAGE_FOLDER_NAME +
            "' の取得/作成に失敗。"
        );
      }

      if (logoFootage.parentFolder !== footageFolderInAE) {
        logoFootage.parentFolder = footageFolderInAE;
        writeLn(
          "  ロゴフッテージをAEプロジェクト内フォルダ '" +
            PRECOMP_FOLDER_NAME +
            "/" +
            FOOTAGE_FOLDER_NAME +
            "' に移動しました。"
        );
      }
    } catch (e) {
      writeLn(
        "警告: ロゴフッテージのAEプロジェクト内フォルダ移動中にエラー: " +
          e.toString()
      );
    }
    writeLn("--- ロゴフッテージ準備完了 ---");

    // --- 2. テキストレイヤーの生成・更新 ---
    writeLn("--- テキストレイヤー生成開始 ---");
    var textLayers = {};
    var textLayerNames = [];
    var layerName;
    for (layerName in LAYER_SPECS) {
      if (!LAYER_SPECS.hasOwnProperty(layerName)) continue;
      writeLn("  Creating text layer: " + layerName); // ★Debug Log
      var spec = LAYER_SPECS[layerName];
      textLayers[layerName] = createTextLayer(activeComp, layerName, spec);
      if (!textLayers[layerName]) {
        writeLn(
          "processComposition: Failed to create text layer: " + layerName
        ); // ★Debug Log
        throw new Error("テキストレイヤー '" + layerName + "' の作成失敗。");
      }
      textLayerNames.push(layerName);
      var targetLayer = textLayers[layerName];
      // トランスフォーム設定
      if (targetLayer.property("Position")) {
        targetLayer.property("Position").setValue(spec.pos);
      }
      if (targetLayer.property("Anchor Point")) {
        targetLayer.property("Anchor Point").setValue(spec.ap);
      }
      // エクスプレッション設定
      if (spec.expression && targetLayer.property("Source Text")) {
        try {
          targetLayer.property("Source Text").expression = spec.expression;
          writeLn("    Set expression for: " + layerName); // ★Debug Log
        } catch (e) {
          writeLn(
            "processComposition: Failed to set expression for layer '" +
              layerName +
              "': " +
              e.toString()
          ); // ★Debug Log
          throw new Error(
            "レイヤー '" +
              layerName +
              "' のエクスプレッション設定失敗。\n" +
              e.toString()
          );
        }
      }
    }
    // エクスプレッションが設定されていないテキストレイヤーの内容を設定
    writeLn("  Setting text layer contents..."); // ★Debug Log
    setTextLayerContents(activeComp, textLayers, pickedObj, projectFileName);
    writeLn("--- テキストレイヤー生成完了 ---");

    // --- 3. ロゴレイヤーをアクティブコンポに追加 ---
    writeLn("--- ロゴレイヤー追加開始 ---");
    var logoLayer = null;
    try {
      logoLayer = activeComp.layers.add(logoFootage);
      if (!logoLayer) throw new Error("ロゴレイヤーの追加に失敗しました。");
      logoLayer.name = LOGO_LAYER_NAME;
      writeLn("ロゴレイヤーを追加しました: " + logoLayer.name);
    } catch (e) {
      writeLn("processComposition: Error adding logo layer: " + e.toString()); // ★Debug Log
      throw new Error("ロゴレイヤー追加中にエラー発生:\n" + e.toString());
    }
    writeLn("--- ロゴレイヤー追加終了 ---");

    // --- 4. 追加したロゴレイヤーのトランスフォーム設定 ---
    writeLn("--- ロゴレイヤートランスフォーム設定開始 ---");
    try {
      if (logoLayer.property("Position")) {
        logoLayer.property("Position").setValue(LOGO_POSITION);
        logoLayer.property("Scale").setValue(LOGO_SCALE);
        writeLn("  ロゴ位置・スケールを設定しました。");
      } else {
        writeLn(
          "警告: ロゴレイヤーのトランスフォームプロパティが見つかりません。"
        );
      }
    } catch (e) {
      writeLn(
        "警告: ロゴレイヤートランスフォーム設定中にエラー発生:\n" + e.toString()
      );
    }
    writeLn("--- ロゴレイヤートランスフォーム設定終了 ---");

    // --- 5. プリコンポーズ処理 (テキストレイヤー + ロゴレイヤー) とフォルダ移動 ---
    writeLn("--- プリコンポーズ処理開始 ---");
    var precompName = PRECOMP_PREFIX + activeComp.name;
    var layersToPrecomposeIndices = [];
    // アクティブコンポのレイヤーを再度チェックして対象インデックスを取得
    for (var i = 1; i <= activeComp.numLayers; i++) {
      var layer = activeComp.layer(i);
      // textLayers オブジェクトにキーとして存在するか（=今回作成したテキストレイヤー）、
      // または今回作成したロゴレイヤーか
      if (textLayers[layer.name] || layer === logoLayer) {
        layersToPrecomposeIndices.push(layer.index);
      }
    }

    writeLn(
      "プリコンポーズ対象レイヤーインデックス: " +
        layersToPrecomposeIndices.join(", ")
    );
    var newComp = null;
    var precompLayer = null;
    if (layersToPrecomposeIndices.length > 0) {
      try {
        writeLn("  Precomposing layers..."); // ★Debug Log
        // 既存の同名プリコンポーズを検索して削除
        var existingPreComp = findItemByName(precompName);
        if (existingPreComp && existingPreComp instanceof CompItem) {
          writeLn("既存のプリコンポーズ '" + precompName + "' を削除します。");
          try {
            existingPreComp.remove();
          } catch (removeError) {
            writeLn(
              "警告: 既存プリコンポーズの削除に失敗: " + removeError.toString()
            );
          }
        }

        // プリコンポーズ実行
        newComp = activeComp.layers.precompose(
          layersToPrecomposeIndices,
          precompName,
          true // move all attributes
        );
        if (newComp) {
          writeLn("プリコンポーズ成功 (テキスト+ロゴ): " + newComp.name);
          // プリコンポーズによって生成された新しいレイヤーを取得
          precompLayer = findLayerByName(activeComp, precompName);
          if (!precompLayer) {
            writeLn("警告: プリコンポーズ後のレイヤーが見つかりません。");
          }

          // プリコンポーズアイテムを AE プロジェクト内の _check-data フォルダに移動
          try {
            var precompTargetFolder = findOrCreateFolder(PRECOMP_FOLDER_NAME);
            if (!precompTargetFolder) {
              writeLn(
                "processComposition: Failed to find/create precompTargetFolder: " +
                  PRECOMP_FOLDER_NAME
              ); // ★Debug Log
              throw new Error(
                "AEプロジェクトフォルダ '" +
                  PRECOMP_FOLDER_NAME +
                  "' の取得/作成に失敗。"
              );
            }

            if (newComp.parentFolder !== precompTargetFolder) {
              newComp.parentFolder = precompTargetFolder;
              writeLn(
                "  プリコンポーズアイテム '" +
                  newComp.name +
                  "' をAEプロジェクト内フォルダ '" +
                  PRECOMP_FOLDER_NAME +
                  "' に移動しました。"
              );
            }
          } catch (folderError) {
            writeLn(
              "警告: プリコンポーズアイテムのフォルダ移動中にエラー: " +
                folderError.toString()
            );
          }
        } else {
          writeLn("processComposition: Precompose returned null."); // ★Debug Log
          throw new Error("プリコンポーズ作成失敗(null)。");
        }
      } catch (e) {
        writeLn("processComposition: Error during precompose: " + e.toString()); // ★Debug Log
        throw new Error(
          "プリコンポーズ中にエラー発生:\n対象インデックス: " +
            layersToPrecomposeIndices.join(",") +
            "\n" +
            e.toString()
        );
      }
    } else {
      writeLn("プリコンポーズ対象レイヤーが見つかりませんでした。");
    }
    writeLn("--- プリコンポーズ処理終了 ---");

    // --- 6. ソースレイヤーを再度追加し、設定 ---
    writeLn("--- ソースレイヤー再追加処理開始 ---");
    var reAddedSourceLayer = null;
    try {
      reAddedSourceLayer = activeComp.layers.add(sourceItem);
      if (!reAddedSourceLayer || !reAddedSourceLayer.isValid) {
        writeLn("processComposition: Failed to re-add source layer."); // ★Debug Log
        throw new Error("ソースレイヤーの再追加失敗。");
      }
      writeLn("ソースレイヤーを再追加しました: " + reAddedSourceLayer.name);

      // 保持していた元のレイヤー名があれば再設定
      if (sourceLayerInfo.name !== "") {
        reAddedSourceLayer.name = sourceLayerInfo.name;
        writeLn("  名前を '" + sourceLayerInfo.name + "' に再設定。");
      }

      // Y座標を中央に配置 (高さ 1150 を想定)
      var posProp = reAddedSourceLayer.property("Position");
      if (posProp && posProp.canSetValue) {
        var currentPos = posProp.value;
        posProp.setValue([currentPos[0], 575]); // 1150 / 2 = 575
        writeLn("  Y座標を 575 に設定。");
      } else {
        writeLn("警告: 再追加したソースレイヤーの位置を設定できません。");
      }

      // レイヤー順序をプリコンポーズレイヤーの直下に移動
      if (precompLayer && precompLayer.isValid) {
        reAddedSourceLayer.moveAfter(precompLayer);
        writeLn("  レイヤー順序をプリコンポーズレイヤーの下に移動。");
      } else {
        // プリコンポーズレイヤーが見つからない場合は、とりあえず最下層へ
        reAddedSourceLayer.moveToEnd();
        writeLn(
          "警告: プリコンポーズレイヤーが見つからないため、最下層に移動。"
        );
      }
    } catch (e) {
      writeLn(
        "警告: ソースレイヤーの再追加または設定中にエラー発生:\n" + e.toString()
      );
    }
    writeLn("--- ソースレイヤー再追加処理終了 ---");
    writeLn("--- processComposition END ---"); // ★Debug Log
  } // processComposition 終了

  // --- テキストレイヤー関連ヘルパー ---
  function createTextLayer(comp, layerName, spec) {
    var textLayer;
    try {
      textLayer = comp.layers.addText("");
      textLayer.name = layerName;
    } catch (e) {
      writeLn(
        "createTextLayer: Failed to add text layer '" +
          layerName +
          "': " +
          e.toString()
      ); // ★Debug Log
      throw new Error(
        "テキストレイヤー追加失敗 ('" + layerName + "')。\n" + e.toString()
      );
    }
    var textProp = textLayer.property("Source Text");
    if (!textProp) {
      writeLn(
        "createTextLayer: Failed to get Source Text property for '" +
          layerName +
          "'"
      ); // ★Debug Log
      throw new Error("Source Textプロパティ取得失敗 ('" + layerName + "')。");
    }
    var textDocument = textProp.value;
    try {
      textDocument.font = FONT_POSTSCRIPT_NAME;
      textDocument.fontSize = FONT_SIZE;
      if (spec.justification != null) {
        textDocument.justification = spec.justification;
      } else {
        textDocument.justification = ParagraphJustification.LEFT_JUSTIFY; // デフォルト
      }
      textProp.setValue(textDocument);
    } catch (e) {
      writeLn(
        "createTextLayer: Failed to set font/justification for '" +
          layerName +
          "': " +
          e.toString()
      ); // ★Debug Log
      try {
        alert(
          "フォント/揃え設定失敗 ('" +
            layerName +
            "')。\nFont: '" +
            FONT_POSTSCRIPT_NAME +
            "'\nError: " +
            e.toString() +
            "\nフォントがインストールされているか確認してください。",
          SCRIPT_NAME
        );
      } catch (err) {
        writeLn("ALERT FAILED: フォント設定失敗 " + layerName);
      }
      try {
        textLayer.remove();
      } catch (removeErr) {
        writeLn(
          "createTextLayer: Failed to remove layer after font error: " +
            removeErr.toString()
        ); // ★Debug Log
      } // 削除失敗は無視
      return null; // 作成失敗を示す
    }
    return textLayer;
  }

  function setTextLayerContents(
    activeComp,
    textLayers,
    pickedObj,
    projectFileName
  ) {
    writeLn("--- setTextLayerContents START ---"); // ★Debug Log
    var layerName;
    var sourceTextProp;

    // 現在日時
    layerName = "current_info";
    sourceTextProp = textLayers[layerName]
      ? textLayers[layerName].property("Source Text")
      : null;
    if (sourceTextProp && !sourceTextProp.expressionEnabled) {
      var dateTimeStr = createDateTimeString();
      writeLn("  Setting " + layerName + " to: " + dateTimeStr); // ★Debug Log
      sourceTextProp.setValue(dateTimeStr);
    }

    // コンポジション名
    layerName = "comp_name";
    sourceTextProp = textLayers[layerName]
      ? textLayers[layerName].property("Source Text")
      : null;
    if (sourceTextProp && !sourceTextProp.expressionEnabled) {
      writeLn("  Setting " + layerName + " to: " + activeComp.name); // ★Debug Log
      sourceTextProp.setValue(activeComp.name);
    }

    // プロジェクト表示名
    layerName = "project_name";
    sourceTextProp = textLayers[layerName]
      ? textLayers[layerName].property("Source Text")
      : null;
    if (sourceTextProp && !sourceTextProp.expressionEnabled) {
      var projNameToSet = projectFileName || getAEProjectIdentifier();
      writeLn("  Setting " + layerName + " to: " + projNameToSet); // ★Debug Log
      sourceTextProp.setValue(projNameToSet); // 設定値がなければファイル名
    }

    // コンポジション情報（ソース情報）
    layerName = "comp_info";
    sourceTextProp = textLayers[layerName]
      ? textLayers[layerName].property("Source Text")
      : null;
    if (sourceTextProp && !sourceTextProp.expressionEnabled) {
      var compInfoString = "N/A";
      try {
        if (pickedObj && pickedObj.source) {
          var sourceItem = pickedObj.source;
          writeLn(
            "  Getting info for source: " +
              sourceItem.name +
              ", Type: " +
              (sourceItem instanceof CompItem ? "Comp" : "Footage")
          ); // ★Debug Log
          var fps = sourceItem.frameRate || 0;
          var dur = sourceItem.duration || 0;
          var totalFrames =
            isNaN(fps) || isNaN(dur) || fps <= 0 || dur <= 0
              ? 0
              : Math.round(fps * dur);
          var tc =
            isNaN(fps) || fps <= 0
              ? "00:00:00:00"
              : convertSecondsToTC(dur, fps);
          var framesPadded = padDigits(totalFrames, 5);
          var sourceName = pickedObj.name || "";
          writeLn(
            "    Source Info: fps=" +
              fps +
              ", dur=" +
              dur +
              ", totalFrames=" +
              totalFrames +
              ", tc=" +
              tc
          ); // ★Debug Log

          // コンポかフッテージかで表示内容を少し変える
          if (pickedObj.type === "comp") {
            // コンポの場合は尺情報のみ
            compInfoString = "TC/Fm: " + tc + "/" + framesPadded;
          } else {
            // フッテージの場合は名前も表示（長すぎる場合は省略）
            if (sourceName.length > 30) {
              sourceName = sourceName.substring(0, 27) + "...";
            }
            compInfoString = sourceName + " TC/Fm: " + tc + "/" + framesPadded;
          }
        } else {
          writeLn("  Source info unavailable (pickedObj or source is null)."); // ★Debug Log
        }
      } catch (e) {
        compInfoString = "Error getting info";
        writeLn("Comp Info 取得エラー: " + e.toString());
      }
      writeLn("  Setting " + layerName + " to: " + compInfoString); // ★Debug Log
      sourceTextProp.setValue(compInfoString);
    }
    writeLn("--- setTextLayerContents END ---"); // ★Debug Log
  }

  // --- After Effects 一般ヘルパー ---
  function findLayerByName(comp, name) {
    try {
      if (!comp || typeof name !== "string" || name === "") return null;
      return comp.layer(name);
    } catch (e) {
      return null;
    }
  }
  function hasSavedProject() {
    var saved = app.project && app.project.file !== null;
    // writeLn("hasSavedProject: Returning " + saved); // ★Debug Log - 必要なら有効化
    return saved;
  }
  function getActiveComp() {
    var ac = app.project.activeItem;
    var comp = ac && ac instanceof CompItem ? ac : null;
    // writeLn("getActiveComp: Returning " + (comp ? comp.name : 'null')); // ★Debug Log - 必要なら有効化
    return comp;
  }
  function findItemByName(name, parentFolder) {
    var proj = app.project;
    if (!proj || typeof name !== "string" || name === "") return null;
    var searchFolder = parentFolder || proj.rootFolder;
    writeLn(
      "findItemByName: Searching for '" +
        name +
        "' in folder '" +
        searchFolder.name +
        "'"
    ); // ★Debug Log

    for (var i = 1; i <= searchFolder.numItems; i++) {
      var item = searchFolder.item(i);
      // writeLn("  Checking item: " + item.name + " (Type: " + item.constructor.name + ")"); // ★Debug Log - 詳細すぎるかも
      // 名前が一致し、かつフォルダでない場合に返す（フォルダとアイテムの同名避け）
      if (item.name === name && !(item instanceof FolderItem)) {
        writeLn("  Found item: " + item.name); // ★Debug Log
        return item;
      }
    }
    writeLn("  Item not found in this folder."); // ★Debug Log
    return null; // 見つからなかった場合
  }

  // 指定したフォルダ内に、指定した名前のフォルダを探し、なければ作成する関数
  function findOrCreateFolder(folderName, parentFolder) {
    var proj = app.project;
    if (!proj) return null;
    var targetParent = parentFolder || proj.rootFolder;
    writeLn(
      "findOrCreateFolder: Looking for folder '" +
        folderName +
        "' in '" +
        targetParent.name +
        "'"
    ); // ★Debug Log

    // 既存フォルダを検索
    for (var i = 1; i <= targetParent.numItems; i++) {
      var item = targetParent.item(i);
      if (item instanceof FolderItem && item.name === folderName) {
        writeLn("  Found existing folder."); // ★Debug Log
        return item;
      }
    }
    // 見つからなければ作成
    writeLn("  Folder not found. Creating new folder..."); // ★Debug Log
    try {
      var newFolder = targetParent.items.addFolder(folderName);
      writeLn("  Successfully created folder: " + newFolder.name); // ★Debug Log
      return newFolder;
    } catch (e) {
      writeLn(
        "フォルダ '" +
          folderName +
          "' の作成に失敗しました in " +
          targetParent.name +
          ": " +
          e.toString()
      );
      return null;
    }
  }

  // アイテムのフォルダパスを文字列で取得（デバッグ用）
  function getItemFolderPath(item) {
    if (!item || !item.parentFolder) return "";
    var path = [];
    var currentFolder = item.parentFolder;
    while (currentFolder && currentFolder !== app.project.rootFolder) {
      path.unshift(currentFolder.name);
      currentFolder = currentFolder.parentFolder;
    }
    return path.join("/");
  }

  // --- ファイルI/Oヘルパー関数 (設定ファイル関連) ---

  // プロジェクトファイル名（拡張子なし）を取得する関数
  // これをプロジェクト名ファイルのファイル名として使用
  function getAEProjectIdentifier() {
    if (!hasSavedProject()) {
      writeLn("getAEProjectIdentifier: Project not saved, returning 'default'"); // ★Debug Log
      return "default"; // 保存されていない場合の仮ID
    }
    var file = app.project.file;
    if (file) {
      try {
        var nm = decodeURI(file.name);
        var id = nm.replace(/\.aep$/i, "");
        writeLn("getAEProjectIdentifier: Returning ID: " + id); // ★Debug Log
        return id;
      } catch (e) {
        // decodeURI失敗時など
        var fallbackId = file.name.replace(/\.aep$/i, "");
        writeLn(
          "getAEProjectIdentifier: decodeURI failed, returning fallback ID: " +
            fallbackId +
            ", Error: " +
            e.toString()
        ); // ★Debug Log
        return fallbackId;
      }
    }
    writeLn(
      "getAEProjectIdentifier: Could not get file object, returning 'default'"
    ); // ★Debug Log
    return "default"; // 何らかの理由で取得できない場合
  }

  // ★追加: プロジェクトファイルのあるディレクトリのパスを取得する関数
  function getProjectDirectoryPath() {
    if (!app.project || !app.project.file) {
      writeLn("getProjectDirectoryPath: Project not saved, returning null");
      return null; // プロジェクトが保存されていない場合は null
    }
    var projPath = "";
    try {
      projPath = app.project.file.path;
      if (projPath === null || projPath === undefined) {
        if (app.project.file.parent) {
          projPath = app.project.file.parent.fsName;
          writeLn(
            "getProjectDirectoryPath: Project path was null, using parent folder fsName: " +
              projPath
          );
        } else {
          writeLn(
            "getProjectDirectoryPath: Project path is null and no parent folder. Cannot determine path."
          );
          return null;
        }
      }
    } catch (e) {
      writeLn(
        "getProjectDirectoryPath: Error getting project path: " + e.toString()
      );
      return null;
    }

    if (projPath === "") {
      writeLn(
        "getProjectDirectoryPath: Project path is empty string. Trying parent folder fsName again."
      );
      if (app.project.file.parent) {
        projPath = app.project.file.parent.fsName;
      } else {
        writeLn(
          "getProjectDirectoryPath: Project path is empty and no parent folder. Cannot determine path."
        );
        return null;
      }
    }

    if (typeof projPath !== "string" || projPath.length === 0) {
      writeLn(
        "getProjectDirectoryPath: Could not determine valid project path. Path: " +
          projPath
      );
      return null;
    }
    writeLn("getProjectDirectoryPath: Returning path: " + projPath);
    return projPath;
  }

  // ★追加: .settings ディレクトリのフルパスを取得する関数
  function getSettingsDirPath() {
    var projDir = getProjectDirectoryPath();
    if (!projDir) {
      writeLn("getSettingsDirPath: Could not get project directory path.");
      return null;
    }
    var settingsPath = projDir + "/" + SETTINGS_DIR_NAME;
    writeLn("getSettingsDirPath: Returning settings dir path: " + settingsPath);
    return settingsPath;
  }

  // ★追加: プロジェクト名設定ファイルのフルパスを取得する関数 (拡張子なし)
  function getProjectNameFilePath() {
    var settingsDir = getSettingsDirPath();
    if (!settingsDir) {
      writeLn("getProjectNameFilePath: Could not get settings directory path.");
      return null;
    }
    var projectId = getAEProjectIdentifier();
    if (projectId === "default") {
      writeLn(
        "getProjectNameFilePath: Invalid project ID 'default'. Cannot determine file path."
      );
      return null;
    }
    // ★修正: 拡張子 (.txt) を連結しない
    var filePath = settingsDir + "/" + projectId;
    writeLn(
      "getProjectNameFilePath: Returning project name file path (extensionless): " +
        filePath
    );
    return filePath;
  }

  // ★追加: プロジェクト名を .settings/[projectName] に保存する関数 (拡張子なし)
  function saveProjectNameToSettingsFile(projectName) {
    writeLn("--- saveProjectNameToSettingsFile START ---"); // ★Debug Log
    var settingsDirPath = getSettingsDirPath();
    var filePath = getProjectNameFilePath(); // 拡張子なしのパスを取得

    if (!settingsDirPath || !filePath) {
      writeLn(
        "saveProjectNameToSettingsFile: Could not get settings/file path. Aborting."
      );
      throw new Error(
        "プロジェクト未保存またはパス取得エラーのため設定保存不可。"
      );
    }

    // .settings ディレクトリが存在するか確認し、なければ作成
    var settingsDir = Folder(settingsDirPath);
    if (!settingsDir.exists) {
      writeLn(
        "saveProjectNameToSettingsFile: Settings directory does not exist. Creating: " +
          settingsDirPath
      );
      if (!settingsDir.create()) {
        writeLn(
          "saveProjectNameToSettingsFile: Failed to create settings directory: " +
            settingsDir.error
        );
        throw new Error(
          "設定ディレクトリの作成に失敗しました: " + settingsDir.error
        );
      }
      writeLn(
        "saveProjectNameToSettingsFile: Settings directory created successfully."
      );
    } else {
      writeLn(
        "saveProjectNameToSettingsFile: Settings directory already exists."
      );
    }

    var projectFile = File(filePath); // File オブジェクトは拡張子なしでも扱える
    writeLn(
      "saveProjectNameToSettingsFile: Writing project name '" +
        projectName +
        "' to: " +
        filePath
    );

    projectFile.encoding = "UTF-8";
    if (!projectFile.open("w")) {
      writeLn(
        "saveProjectNameToSettingsFile: Error opening file for write: " +
          projectFile.error
      );
      throw new Error(
        "プロジェクト名ファイルを開けませんでした(w): " + projectFile.error
      );
    }
    try {
      // 引数で受け取った projectName (UIの入力値) をそのまま書き込む
      var writeSuccess = projectFile.write(projectName);
      if (!writeSuccess) {
        var writeErrorMsg = projectFile.error || "Unknown write error";
        writeLn(
          "saveProjectNameToSettingsFile: File write failed: " + writeErrorMsg
        );
        throw new Error("プロジェクト名ファイル書き込み失敗: " + writeErrorMsg);
      }
      writeLn("saveProjectNameToSettingsFile: Write successful.");
    } catch (writeError) {
      writeLn(
        "saveProjectNameToSettingsFile: Exception during file write: " +
          writeError.toString()
      );
      projectFile.close(); // 書き込みエラー時も閉じる
      throw writeError;
    } finally {
      if (projectFile.isOpen) projectFile.close();
    }
    writeLn("--- saveProjectNameToSettingsFile END ---"); // ★Debug Log
  }

  // ★追加: .settings/[projectName] からプロジェクト名を読み込む関数 (拡張子なし)
  function loadProjectNameFromSettingsFile() {
    writeLn("--- loadProjectNameFromSettingsFile START ---");
    var filePath = getProjectNameFilePath(); // 拡張子なしのパスを取得
    if (!filePath) {
      writeLn(
        "loadProjectNameFromSettingsFile: Could not get file path. Returning empty string."
      );
      writeLn("--- loadProjectNameFromSettingsFile END (No Path) ---");
      return ""; // パスが取得できない場合は空を返す
    }

    var projectFile = File(filePath); // File オブジェクトは拡張子なしでも扱える
    if (!projectFile.exists) {
      writeLn(
        "loadProjectNameFromSettingsFile: File not found: " +
          filePath +
          ". Returning empty string."
      );
      writeLn("--- loadProjectNameFromSettingsFile END (Not Found) ---");
      return ""; // ファイルが存在しない場合は空を返す
    }

    writeLn("loadProjectNameFromSettingsFile: Reading file: " + filePath);
    projectFile.encoding = "UTF-8";
    if (!projectFile.open("r")) {
      writeLn(
        "loadProjectNameFromSettingsFile: Error opening file for read: " +
          projectFile.error +
          ". Returning empty string."
      );
      writeLn("--- loadProjectNameFromSettingsFile END (Error Open) ---");
      return ""; // ファイルを開けない場合は空を返す
    }

    var projectName = "";
    try {
      projectName = projectFile.read();
      writeLn(
        "loadProjectNameFromSettingsFile: Read project name: " + projectName
      );
    } catch (readError) {
      writeLn(
        "loadProjectNameFromSettingsFile: Error reading file: " +
          readError.toString() +
          ". Returning empty string."
      );
      projectName = ""; // 読み込みエラーの場合も空にする
      writeLn("--- loadProjectNameFromSettingsFile END (Error Read) ---");
    } finally {
      if (projectFile.isOpen) projectFile.close();
    }

    writeLn("--- loadProjectNameFromSettingsFile END (Success/Fail) ---");
    return projectName; // 読み込んだ内容 (空文字列の場合もある) を返す
  }

  // --- ユーティリティ関数 ---
  function createDateTimeString() {
    var d = new Date();
    function pad(num) {
      return ("0" + num).slice(-2);
    }
    var y = d.getFullYear();
    var m = pad(d.getMonth() + 1);
    var day = pad(d.getDate());
    var h = pad(d.getHours());
    var min = pad(d.getMinutes());
    var s = pad(d.getSeconds());
    // タイムゾーンオフセットの書式を修正 (例: +0900)
    var offsetMinutes = -d.getTimezoneOffset();
    var sign = offsetMinutes >= 0 ? "+" : "-";
    offsetMinutes = Math.abs(offsetMinutes);
    var offsetH = Math.floor(offsetMinutes / 60);
    var offsetM = offsetMinutes % 60;
    function padOffset(v) {
      return (v < 10 ? "0" : "") + v;
    }
    var offsetStr = sign + padOffset(offsetH) + padOffset(offsetM); // 区切り文字なし
    return (
      y +
      "/" +
      m +
      "/" +
      day +
      " " +
      h +
      ":" +
      min +
      ":" +
      s +
      " (" +
      offsetStr +
      ")"
    ); // より一般的な書式に変更
  }

  function convertSecondsToTC(durationSec, fps) {
    if (isNaN(durationSec) || isNaN(fps) || fps <= 0 || durationSec < 0) {
      return "00:00:00:00";
    }
    try {
      var totalFrames = Math.floor(durationSec * fps + 0.00001); // 丸め誤差考慮
      var rate = Math.round(fps); // 計算用のフレームレート整数

      // 簡易計算
      var isDrop = Math.abs(fps - 29.97) < 0.01 || Math.abs(fps - 59.94) < 0.01;
      function pad2(n) {
        return n < 10 ? "0" + n : "" + n;
      }
      var hours = Math.floor(totalFrames / (rate * 3600));
      var minutes = Math.floor((totalFrames % (rate * 3600)) / (rate * 60));
      var seconds = Math.floor((totalFrames % (rate * 60)) / rate);
      var frames = totalFrames % rate;

      if (isDrop) {
        return (
          pad2(hours) +
          ":" +
          pad2(minutes) +
          ":" +
          pad2(seconds) +
          ";" +
          pad2(frames)
        );
      } else {
        return (
          pad2(hours) +
          ":" +
          pad2(minutes) +
          ":" +
          pad2(seconds) +
          ":" +
          pad2(frames)
        );
      }
    } catch (e) {
      writeLn("Timecode calculation error: " + e.toString());
      return "TC Error";
    }
  }

  function padDigits(num, digits) {
    var s = String(Math.max(0, Math.round(num)));
    while (s.length < digits) {
      s = "0" + s;
    }
    return s;
  }

  // --- スクリプト実行開始 ---
  initializeLogFile(); // ログファイルに開始マーカーを書き込む
  writeLn("Starting script: " + SCRIPT_NAME + " " + SCRIPT_VERSION);
  var myPanel = buildUI(thisObj);
})(this);
