// ================================================
// KRACT Check Information AutoLayer
// Version 2.40 (Remove Top Margin/Border)
// ================================================

(function (thisObj) {
  // --- グローバル設定 ---
  var SCRIPT_NAME = "Check Information AutoLayer";
  var SCRIPT_VERSION = "2.40_no_top_margin"; // バージョン更新
  var FONT_POSTSCRIPT_NAME = "GeistMono-Medium";
  var FONT_SIZE = 20;
  var PRECOMP_PREFIX = "_check-data_";
  var PRECOMP_FOLDER_NAME = "_check-data";
  var FOOTAGE_FOLDER_NAME = "footage";
  var SETTINGS_DIR_NAME = ".settings";
  var LOGO_MANAGE_BASE_FOLDER_NAME = "Check Information AutoLayer";
  var LOGO_MANAGE_SUBFOLDER_NAME = "footage";
  var LOGO_TARGET_FILENAME = "checkinfo-logo.ai";
  var LOGO_FOOTAGE_NAME = "checkinfo-logo.ai";
  var LOGO_LAYER_NAME = "logo";
  var LOGO_POSITION = [960, 1131.75];
  var LOGO_SCALE = [16, 16];

  // ★★★ デバッグログファイル設定 ★★★
  var ENABLE_DEBUG_LOGGING = false;
  var DEBUG_LOG_FILENAME = "CheckInfo_DebugLog.txt";
  var DEBUG_LOG_FILE_PATH = Folder.desktop.fsName + "/" + DEBUG_LOG_FILENAME;

  // --- ログ書き込み関数 ---
  function logToFile(message) {
    if (!ENABLE_DEBUG_LOGGING) {
      return;
    }
    try {
      var logFile = File(DEBUG_LOG_FILE_PATH);
      logFile.encoding = "UTF-8";
      if (logFile.open("a")) {
        var currentTime = new Date().toLocaleTimeString([], { hour12: false });
        logFile.writeln("[" + currentTime + "] " + message);
        logFile.close();
      } else {
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
  function writeLn(message) {
    logToFile(message);
  }
  function initializeLogFile() {
    if (!ENABLE_DEBUG_LOGGING) {
      return;
    }
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
        : new Window("palette", "", undefined, { resizeable: true });
    myPanel.orientation = "column";
    myPanel.alignChildren = ["fill", "top"];
    myPanel.spacing = 10;
    // ★修正: 上マージンを0に設定 [left, top, right, bottom]
    myPanel.margins = [15, 0, 15, 15];

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
      "プロジェクト表示名。「.settings/[projectName]」ファイルがあれば自動入力。変更は実行時に保存されます。";
    var panelLogo = myPanel.add("panel", undefined, "ロゴ設定");
    panelLogo.orientation = "column";
    panelLogo.alignChildren = "fill";
    panelLogo.spacing = 5;
    panelLogo.margins = 10;
    panelLogo.add("statictext", undefined, "使用するロゴ:");
    var grpLogoPath = panelLogo.add("group");
    grpLogoPath.orientation = "row";
    grpLogoPath.alignChildren = ["fill", "center"];
    var stLogoPath = grpLogoPath.add("statictext", [0, 0, 200, 20], "未設定", {
      truncate: "middle",
    });
    stLogoPath.helpTip =
      "管理下のロゴファイル (" + managedLogoFile.fsName + ") の状態";
    var btnBrowseLogo = grpLogoPath.add("button", undefined, "参照");
    btnBrowseLogo.helpTip =
      "ロゴファイル (AI形式) を選択し、管理フォルダにコピーします";

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
            updateLogoPathDisplay();
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
      writeLn("--- btnApply onClick START ---");
      if (!hasSavedProject()) {
        try {
          alert("プロジェクトが保存されていません。", SCRIPT_NAME);
        } catch (e) {
          writeLn("ALERT FAILED: プロジェクト未保存");
        }
        writeLn("btnApply onClick: Project not saved.");
        return;
      }
      var currentAvSources = getAvailableSources(getActiveComp());
      writeLn(
        "btnApply onClick: Got currentAvSources: " +
          (currentAvSources ? currentAvSources.length : 0) +
          " items"
      );
      updateLogoPathDisplay();
      var ac = getActiveComp();
      if (!ac) {
        try {
          alert("アクティブなコンポジションがありません。", SCRIPT_NAME);
        } catch (e) {
          writeLn("ALERT FAILED: アクティブコンポなし");
        }
        writeLn("btnApply onClick: No active comp.");
        return;
      }
      if (!managedLogoFile.exists) {
        writeLn(
          "btnApply onClick: Managed logo file not found: " +
            managedLogoFile.fsName
        );
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
        ddSource.selection.text === "プロジェクト未保存 または コンポなし"
      ) {
        writeLn("btnApply onClick: Invalid source selection in dropdown.");
        try {
          alert("有効なソースが選択されていません。", SCRIPT_NAME);
        } catch (e) {
          writeLn("ALERT FAILED: 有効ソース未選択");
        }
        return;
      }
      var selIndex = ddSource.selection.index;
      if (
        !currentAvSources ||
        selIndex >= currentAvSources.length ||
        currentAvSources[selIndex].name !== ddSource.selection.text
      ) {
        writeLn("btnApply onClick: Source selection mismatch or outdated.");
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
      );
      if (!pickedObj || !pickedObj.source) {
        writeLn("btnApply onClick: Failed to get source info from pickedObj.");
        try {
          alert("選択されたソース情報の取得に失敗しました。", SCRIPT_NAME);
        } catch (e) {
          writeLn("ALERT FAILED: ソース情報取得失敗");
        }
        return;
      }
      var projectNameToSave = editProjectName.text;
      writeLn(
        "btnApply onClick: Project name from UI to save: " + projectNameToSave
      );
      var projectFileNameForLayer =
        projectNameToSave !== "" ? projectNameToSave : getAEProjectIdentifier();
      writeLn(
        "btnApply onClick: projectFileNameForLayer for display: " +
          projectFileNameForLayer
      );
      try {
        writeLn("btnApply onClick: Calling saveProjectNameToSettingsFile...");
        saveProjectNameToSettingsFile(projectNameToSave);
        writeLn("プロジェクト設定 (.settings/[projectName]) を保存しました。");
      } catch (e) {
        writeLn("btnApply onClick: Error saving project name: " + e.toString());
        try {
          alert("設定の保存に失敗しました:\n" + e.toString(), SCRIPT_NAME);
        } catch (err) {
          writeLn("ALERT FAILED: 設定保存失敗: " + e.toString());
        }
      }
      writeLn(
        "btnApply onClick: Starting Undo Group and processComposition..."
      );
      app.beginUndoGroup(SCRIPT_NAME + ": 実行");
      try {
        processComposition(ac, pickedObj, projectFileNameForLayer);
        writeLn("btnApply onClick: processComposition completed.");
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
        );
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
        writeLn("--- btnApply onClick END ---");
      }
    };
    function initializePanel() {
      writeLn("--- initializePanel START ---");
      if (hasSavedProject()) {
        writeLn("initializePanel: Project is saved.");
        try {
          var loadedProjectName = loadProjectNameFromSettingsFile();
          writeLn("initializePanel: Loaded project name: " + loadedProjectName);
          editProjectName.text = loadedProjectName || "";
        } catch (e) {
          writeLn(
            "initializePanel: Error loading project name: " + e.toString()
          );
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
        btnBrowseLogo.enabled = true;
      } else {
        writeLn(
          "initializePanel: Project is NOT saved. Disabling UI elements."
        );
        editProjectName.text = "";
        editProjectName.enabled = false;
        btnBrowseLogo.enabled = false;
        ddSource.removeAll();
        ddSource.add("item", "プロジェクト未保存 または コンポなし");
        ddSource.selection = 0;
        ddSource.enabled = false;
        btnApply.enabled = false;
        btnRefresh.enabled = false;
      }
      updateLogoPathDisplay();
      writeLn("--- initializePanel END ---");
    }
    function getAvailableSources(comp) {
      var sources = [];
      if (!comp) return sources;
      for (var i = 1; i <= comp.numLayers; i++) {
        var lyr = comp.layer(i);
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
          if (s instanceof CompItem) {
            sources.push({ source: s, type: "comp", name: s.name });
          } else if (s instanceof FootageItem) {
            if (s.mainSource && !s.mainSource.isStill) {
              sources.push({ source: s, type: "footage", name: s.name });
            }
          }
        }
      }
      return sources;
    }
    function refreshSourceList(dropdown) {
      writeLn("--- refreshSourceList START ---");
      dropdown.removeAll();
      var currentSources = [];
      var ac = getActiveComp();
      var projectSaved = hasSavedProject();
      writeLn(
        "refreshSourceList: projectSaved=" +
          projectSaved +
          ", activeComp=" +
          (ac ? ac.name : "null")
      );
      var enableUI = projectSaved && ac;
      ddSource.enabled = enableUI;
      btnApply.enabled = enableUI;
      editProjectName.enabled = projectSaved;
      btnBrowseLogo.enabled = true;
      btnRefresh.enabled = projectSaved;
      if (!projectSaved) {
        dropdown.add("item", "プロジェクト未保存");
        dropdown.selection = 0;
        btnApply.enabled = false;
        writeLn("refreshSourceList: Project not saved state.");
      } else if (!ac) {
        dropdown.add("item", "アクティブなコンポなし");
        dropdown.selection = 0;
        btnApply.enabled = false;
        writeLn("refreshSourceList: No active comp state.");
      } else {
        currentSources = getAvailableSources(ac);
        writeLn(
          "refreshSourceList: Found " +
            currentSources.length +
            " available sources."
        );
        if (currentSources.length === 0) {
          dropdown.add("item", "有効なソースレイヤーなし");
          dropdown.selection = 0;
          btnApply.enabled = false;
          writeLn("refreshSourceList: No valid source layers found state.");
        } else {
          for (var k = 0; k < currentSources.length; k++) {
            dropdown.add("item", currentSources[k].name);
          }
          dropdown.selection = 0;
          btnApply.enabled = true;
          writeLn("refreshSourceList: Populated dropdown. Selection index 0.");
        }
      }
      writeLn("--- refreshSourceList END ---");
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

  // ロゴファイルを指定された管理フォルダにコピー
  function setAndCopyLogo(sourceFile) {
    if (!sourceFile || !sourceFile.exists)
      throw new Error("指定されたロゴファイルが見つかりません。");
    if (!/\.ai$/i.test(sourceFile.name))
      throw new Error(
        "ロゴファイルは Adobe Illustrator (.ai) 形式である必要があります。"
      );
    var baseDir = Folder(
      Folder.myDocuments.fsName + "/" + LOGO_MANAGE_BASE_FOLDER_NAME
    );
    if (!baseDir.exists) {
      if (!baseDir.create())
        throw new Error(
          "書類フォルダ内にベースディレクトリを作成できませんでした: " +
            baseDir.fsName
        );
      writeLn("ロゴ管理用ベースディレクトリを作成しました: " + baseDir.fsName);
    }
    var targetDir = Folder(baseDir.fsName + "/" + LOGO_MANAGE_SUBFOLDER_NAME);
    if (!targetDir.exists) {
      if (!targetDir.create())
        throw new Error(
          "ロゴ管理用サブディレクトリを作成できませんでした: " +
            targetDir.fsName
        );
      writeLn("ロゴ管理用サブディレクトリを作成しました: " + targetDir.fsName);
    }
    var targetFile = File(targetDir.fsName + "/" + LOGO_TARGET_FILENAME);
    if (sourceFile.copy(targetFile.fsName)) {
      writeLn(
        "ロゴファイルを管理フォルダにコピーしました: " +
          sourceFile.fsName +
          " -> " +
          targetFile.fsName
      );
      return targetFile.fsName;
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
  function processComposition(activeComp, pickedObj, projectFileName) {
    writeLn("--- processComposition START ---");
    if (!activeComp) {
      writeLn("processComposition: No activeComp. Returning.");
      return;
    }
    if (!pickedObj || !pickedObj.source) {
      writeLn("processComposition: Invalid pickedObj or source.");
      throw new Error("処理に必要なソース情報がありません。");
    }
    if (!managedLogoFile.exists) {
      writeLn(
        "processComposition: Managed logo file does not exist: " +
          managedLogoFile.fsName
      );
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
    );
    var sourceItem = pickedObj.source;
    var sourceLayerInfo = { index: -1, name: "" };
    writeLn("--- 既存レイヤー検索・削除開始 ---");
    var foundAndRemovedSource = false;
    for (var i = activeComp.numLayers; i >= 1; i--) {
      var currentLayer = activeComp.layer(i);
      var isManagedLayer = false;
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
          if (!foundAndRemovedSource) {
            sourceLayerInfo.index = currentLayer.index;
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
        existingLogoFootage.replaceSource(managedLogoFile, false);
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
        var importOptions = new ImportOptions(managedLogoFile);
        logoFootage = app.project.importFile(importOptions);
        if (!logoFootage)
          throw new Error(
            "管理下のロゴファイル (" +
              managedLogoFile.fsName +
              ") のインポートに失敗しました。"
          );
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
      );
      throw new Error(
        "ロゴフッテージの準備（インポート/置換）中にエラーが発生しました:\n" +
          e.toString()
      );
    }
    if (!logoFootage || !(logoFootage instanceof FootageItem)) {
      writeLn("processComposition: Failed to get valid logoFootage item.");
      throw new Error(
        "ロゴフッテージの準備に失敗しました。有効なフッテージアイテムを取得できませんでした。"
      );
    }
    try {
      var checkDataFolderForLogo = findOrCreateFolder(PRECOMP_FOLDER_NAME);
      if (!checkDataFolderForLogo)
        throw new Error(
          "AEプロジェクトフォルダ '" +
            PRECOMP_FOLDER_NAME +
            "' の取得/作成に失敗。"
        );
      var footageFolderInAE = findOrCreateFolder(
        FOOTAGE_FOLDER_NAME,
        checkDataFolderForLogo
      );
      if (!footageFolderInAE)
        throw new Error(
          "AEプロジェクトフォルダ '" +
            FOOTAGE_FOLDER_NAME +
            "' の取得/作成に失敗。"
        );
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
    writeLn("--- テキストレイヤー生成開始 ---");
    var textLayers = {};
    var textLayerNames = [];
    var layerName;
    for (layerName in LAYER_SPECS) {
      if (!LAYER_SPECS.hasOwnProperty(layerName)) continue;
      writeLn("  Creating text layer: " + layerName);
      var spec = LAYER_SPECS[layerName];
      textLayers[layerName] = createTextLayer(activeComp, layerName, spec);
      if (!textLayers[layerName]) {
        writeLn(
          "processComposition: Failed to create text layer: " + layerName
        );
        throw new Error("テキストレイヤー '" + layerName + "' の作成失敗。");
      }
      textLayerNames.push(layerName);
      var targetLayer = textLayers[layerName];
      if (targetLayer.property("Position"))
        targetLayer.property("Position").setValue(spec.pos);
      if (targetLayer.property("Anchor Point"))
        targetLayer.property("Anchor Point").setValue(spec.ap);
      if (spec.expression && targetLayer.property("Source Text")) {
        try {
          targetLayer.property("Source Text").expression = spec.expression;
          writeLn("    Set expression for: " + layerName);
        } catch (e) {
          writeLn(
            "processComposition: Failed to set expression for layer '" +
              layerName +
              "': " +
              e.toString()
          );
          throw new Error(
            "レイヤー '" +
              layerName +
              "' のエクスプレッション設定失敗。\n" +
              e.toString()
          );
        }
      }
    }
    writeLn("  Setting text layer contents...");
    setTextLayerContents(activeComp, textLayers, pickedObj, projectFileName);
    writeLn("--- テキストレイヤー生成完了 ---");
    writeLn("--- ロゴレイヤー追加開始 ---");
    var logoLayer = null;
    try {
      logoLayer = activeComp.layers.add(logoFootage);
      if (!logoLayer) throw new Error("ロゴレイヤーの追加に失敗しました。");
      logoLayer.name = LOGO_LAYER_NAME;
      writeLn("ロゴレイヤーを追加しました: " + logoLayer.name);
    } catch (e) {
      writeLn("processComposition: Error adding logo layer: " + e.toString());
      throw new Error("ロゴレイヤー追加中にエラー発生:\n" + e.toString());
    }
    writeLn("--- ロゴレイヤー追加終了 ---");
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
    writeLn("--- プリコンポーズ処理開始 ---");
    var precompName = PRECOMP_PREFIX + activeComp.name;
    var layersToPrecomposeIndices = [];
    for (var i = 1; i <= activeComp.numLayers; i++) {
      var layer = activeComp.layer(i);
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
        writeLn("  Precomposing layers...");
        var checkDataFolder = findOrCreateFolder(PRECOMP_FOLDER_NAME);
        if (!checkDataFolder) {
          writeLn(
            "警告: 親フォルダ '" +
              PRECOMP_FOLDER_NAME +
              "' の取得/作成に失敗。既存プリコンポーズ削除をスキップする可能性があります。"
          );
        } else {
          var existingPreCompInFolder = findItemByName(
            precompName,
            checkDataFolder
          );
          if (
            existingPreCompInFolder &&
            existingPreCompInFolder instanceof CompItem
          ) {
            writeLn(
              "既存のプリコンポーズ '" +
                precompName +
                "' をフォルダ '" +
                checkDataFolder.name +
                "' 内で発見。削除します。"
            );
            try {
              existingPreCompInFolder.remove();
              writeLn(
                "  フォルダ内のプリコンポーズ削除を試行しました: " + precompName
              );
            } catch (removeError) {
              writeLn(
                "警告: フォルダ内の既存プリコンポーズの削除に失敗: " +
                  removeError.toString()
              );
            }
          } else {
            writeLn(
              "フォルダ '" +
                checkDataFolder.name +
                "' 内に既存プリコンポーズ '" +
                precompName +
                "' は見つかりませんでした。"
            );
          }
          var nestedCheckDataFolder = null;
          for (var j = 1; j <= checkDataFolder.numItems; j++) {
            var itemInside = checkDataFolder.item(j);
            if (
              itemInside instanceof FolderItem &&
              itemInside.name === PRECOMP_FOLDER_NAME
            ) {
              nestedCheckDataFolder = itemInside;
              break;
            }
          }
          if (nestedCheckDataFolder) {
            writeLn(
              "警告: フォルダ '" +
                checkDataFolder.name +
                "' 内にネストしたフォルダ '" +
                nestedCheckDataFolder.name +
                "' を発見。フォルダごと削除します。"
            );
            try {
              nestedCheckDataFolder.remove();
              writeLn(
                "  ネストしたフォルダの削除を試行しました: " +
                  nestedCheckDataFolder.name
              );
            } catch (removeFolderError) {
              writeLn(
                "警告: ネストしたフォルダの削除に失敗: " +
                  removeFolderError.toString()
              );
            }
          } else {
            writeLn(
              "フォルダ '" +
                checkDataFolder.name +
                "' 内にネストしたフォルダ '" +
                PRECOMP_FOLDER_NAME +
                "' は見つかりませんでした。"
            );
          }
        }
        var existingPreCompRoot = findItemByName(precompName);
        if (existingPreCompRoot && existingPreCompRoot instanceof CompItem) {
          writeLn(
            "既存のプリコンポーズ '" +
              precompName +
              "' をルートフォルダ直下で発見。削除します。"
          );
          try {
            existingPreCompRoot.remove();
            writeLn(
              "  ルートのプリコンポーズ削除を試行しました: " + precompName
            );
          } catch (removeErrorRoot) {
            writeLn(
              "警告: ルートの既存プリコンポーズの削除に失敗: " +
                removeErrorRoot.toString()
            );
          }
        } else {
          writeLn(
            "ルートフォルダ直下に既存プリコンポーズ '" +
              precompName +
              "' は見つかりませんでした。"
          );
        }
        newComp = activeComp.layers.precompose(
          layersToPrecomposeIndices,
          precompName,
          true
        );
        if (newComp) {
          writeLn("プリコンポーズ成功 (テキスト+ロゴ): " + newComp.name);
          precompLayer = findLayerByName(activeComp, precompName);
          if (!precompLayer)
            writeLn("警告: プリコンポーズ後のレイヤーが見つかりません。");
          try {
            var precompTargetFolder = findOrCreateFolder(PRECOMP_FOLDER_NAME);
            if (!precompTargetFolder)
              throw new Error(
                "AEプロジェクトフォルダ '" +
                  PRECOMP_FOLDER_NAME +
                  "' の取得/作成に失敗。"
              );
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
          writeLn("processComposition: Precompose returned null.");
          throw new Error("プリコンポーズ作成失敗(null)。");
        }
      } catch (e) {
        writeLn("processComposition: Error during precompose: " + e.toString());
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
    writeLn("--- ソースレイヤー再追加処理開始 ---");
    var reAddedSourceLayer = null;
    try {
      reAddedSourceLayer = activeComp.layers.add(sourceItem);
      if (!reAddedSourceLayer || !reAddedSourceLayer.isValid) {
        writeLn("processComposition: Failed to re-add source layer.");
        throw new Error("ソースレイヤーの再追加失敗。");
      }
      writeLn("ソースレイヤーを再追加しました: " + reAddedSourceLayer.name);
      if (sourceLayerInfo.name !== "") {
        reAddedSourceLayer.name = sourceLayerInfo.name;
        writeLn("  名前を '" + sourceLayerInfo.name + "' に再設定。");
      }
      var posProp = reAddedSourceLayer.property("Position");
      if (posProp && posProp.canSetValue) {
        var currentPos = posProp.value;
        posProp.setValue([currentPos[0], 575]);
        writeLn("  Y座標を 575 に設定。");
      } else {
        writeLn("警告: 再追加したソースレイヤーの位置を設定できません。");
      }
      if (precompLayer && precompLayer.isValid) {
        reAddedSourceLayer.moveAfter(precompLayer);
        writeLn("  レイヤー順序をプリコンポーズレイヤーの下に移動。");
      } else {
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
    writeLn("--- processComposition END ---");
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
      );
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
      );
      throw new Error("Source Textプロパティ取得失敗 ('" + layerName + "')。");
    }
    var textDocument = textProp.value;
    try {
      textDocument.font = FONT_POSTSCRIPT_NAME;
      textDocument.fontSize = FONT_SIZE;
      if (spec.justification != null)
        textDocument.justification = spec.justification;
      else textDocument.justification = ParagraphJustification.LEFT_JUSTIFY;
      textProp.setValue(textDocument);
    } catch (e) {
      writeLn(
        "createTextLayer: Failed to set font/justification for '" +
          layerName +
          "': " +
          e.toString()
      );
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
        );
      }
      return null;
    }
    return textLayer;
  }
  function setTextLayerContents(
    activeComp,
    textLayers,
    pickedObj,
    projectFileName
  ) {
    writeLn("--- setTextLayerContents START ---");
    var layerName;
    var sourceTextProp;
    layerName = "current_info";
    sourceTextProp = textLayers[layerName]
      ? textLayers[layerName].property("Source Text")
      : null;
    if (sourceTextProp && !sourceTextProp.expressionEnabled) {
      var dateTimeStr = createDateTimeString();
      writeLn("  Setting " + layerName + " to: " + dateTimeStr);
      sourceTextProp.setValue(dateTimeStr);
    }
    layerName = "comp_name";
    sourceTextProp = textLayers[layerName]
      ? textLayers[layerName].property("Source Text")
      : null;
    if (sourceTextProp && !sourceTextProp.expressionEnabled) {
      writeLn("  Setting " + layerName + " to: " + activeComp.name);
      sourceTextProp.setValue(activeComp.name);
    }
    layerName = "project_name";
    sourceTextProp = textLayers[layerName]
      ? textLayers[layerName].property("Source Text")
      : null;
    if (sourceTextProp && !sourceTextProp.expressionEnabled) {
      var projNameToSet = projectFileName || getAEProjectIdentifier();
      writeLn("  Setting " + layerName + " to: " + projNameToSet);
      sourceTextProp.setValue(projNameToSet);
    }
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
          );
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
          );
          if (pickedObj.type === "comp") {
            compInfoString = "TC/Fm: " + tc + "/" + framesPadded;
          } else {
            if (sourceName.length > 30)
              sourceName = sourceName.substring(0, 27) + "...";
            compInfoString = sourceName + " TC/Fm: " + tc + "/" + framesPadded;
          }
        } else {
          writeLn("  Source info unavailable (pickedObj or source is null).");
        }
      } catch (e) {
        compInfoString = "Error getting info";
        writeLn("Comp Info 取得エラー: " + e.toString());
      }
      writeLn("  Setting " + layerName + " to: " + compInfoString);
      sourceTextProp.setValue(compInfoString);
    }
    writeLn("--- setTextLayerContents END ---");
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
    return app.project && app.project.file !== null;
  }
  function getActiveComp() {
    var ac = app.project.activeItem;
    return ac && ac instanceof CompItem ? ac : null;
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
    );
    for (var i = 1; i <= searchFolder.numItems; i++) {
      var item = searchFolder.item(i);
      if (item.name === name && !(item instanceof FolderItem)) {
        writeLn("  Found item: " + item.name);
        return item;
      }
    }
    writeLn("  Item not found in this folder.");
    return null;
  }
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
    );
    for (var i = 1; i <= targetParent.numItems; i++) {
      var item = targetParent.item(i);
      if (item instanceof FolderItem && item.name === folderName) {
        writeLn("  Found existing folder.");
        return item;
      }
    }
    writeLn("  Folder not found. Creating new folder...");
    try {
      var newFolder = targetParent.items.addFolder(folderName);
      writeLn("  Successfully created folder: " + newFolder.name);
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

  // --- ファイルI/Oヘルパー関数 ---
  function getAEProjectIdentifier() {
    if (!hasSavedProject()) {
      writeLn("getAEProjectIdentifier: Project not saved, returning 'default'");
      return "default";
    }
    var file = app.project.file;
    if (file) {
      try {
        var nm = decodeURI(file.name);
        var id = nm.replace(/\.aep$/i, "");
        writeLn("getAEProjectIdentifier: Returning ID: " + id);
        return id;
      } catch (e) {
        var fallbackId = file.name.replace(/\.aep$/i, "");
        writeLn(
          "getAEProjectIdentifier: decodeURI failed, returning fallback ID: " +
            fallbackId +
            ", Error: " +
            e.toString()
        );
        return fallbackId;
      }
    }
    writeLn(
      "getAEProjectIdentifier: Could not get file object, returning 'default'"
    );
    return "default";
  }
  function getProjectDirectoryPath() {
    if (!app.project || !app.project.file) {
      writeLn("getProjectDirectoryPath: Project not saved, returning null");
      return null;
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
    var filePath = settingsDir + "/" + projectId;
    writeLn(
      "getProjectNameFilePath: Returning project name file path (extensionless): " +
        filePath
    );
    return filePath;
  }
  function saveProjectNameToSettingsFile(projectName) {
    writeLn("--- saveProjectNameToSettingsFile START ---");
    var settingsDirPath = getSettingsDirPath();
    var filePath = getProjectNameFilePath();
    if (!settingsDirPath || !filePath) {
      writeLn(
        "saveProjectNameToSettingsFile: Could not get settings/file path. Aborting."
      );
      throw new Error(
        "プロジェクト未保存またはパス取得エラーのため設定保存不可。"
      );
    }
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
    var projectFile = File(filePath);
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
      projectFile.close();
      throw writeError;
    } finally {
      if (projectFile.isOpen) projectFile.close();
    }
    writeLn("--- saveProjectNameToSettingsFile END ---");
  }
  function loadProjectNameFromSettingsFile() {
    writeLn("--- loadProjectNameFromSettingsFile START ---");
    var filePath = getProjectNameFilePath();
    if (!filePath) {
      writeLn(
        "loadProjectNameFromSettingsFile: Could not get file path. Returning empty string."
      );
      writeLn("--- loadProjectNameFromSettingsFile END (No Path) ---");
      return "";
    }
    var projectFile = File(filePath);
    if (!projectFile.exists) {
      writeLn(
        "loadProjectNameFromSettingsFile: File not found: " +
          filePath +
          ". Returning empty string."
      );
      writeLn("--- loadProjectNameFromSettingsFile END (Not Found) ---");
      return "";
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
      return "";
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
      projectName = "";
      writeLn("--- loadProjectNameFromSettingsFile END (Error Read) ---");
    } finally {
      if (projectFile.isOpen) projectFile.close();
    }
    writeLn("--- loadProjectNameFromSettingsFile END (Success/Fail) ---");
    return projectName;
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
    var offsetMinutes = -d.getTimezoneOffset();
    var sign = offsetMinutes >= 0 ? "+" : "-";
    offsetMinutes = Math.abs(offsetMinutes);
    var offsetH = Math.floor(offsetMinutes / 60);
    var offsetM = offsetMinutes % 60;
    var offsetStr = sign + pad(offsetH) + ":" + pad(offsetM);
    return (
      y +
      "-" +
      m +
      "-" +
      day +
      " T " +
      h +
      ":" +
      min +
      ":" +
      s +
      " " +
      offsetStr
    );
  }
  function convertSecondsToTC(durationSec, fps) {
    if (isNaN(durationSec) || isNaN(fps) || fps <= 0 || durationSec < 0)
      return "00:00:00:00";
    try {
      var totalFrames = Math.floor(durationSec * fps + 0.00001);
      var rate = Math.round(fps);
      var isDrop = Math.abs(fps - 29.97) < 0.01 || Math.abs(fps - 59.94) < 0.01;
      function pad2(n) {
        return n < 10 ? "0" + n : "" + n;
      }
      var hours = Math.floor(totalFrames / (rate * 3600));
      var minutes = Math.floor((totalFrames % (rate * 3600)) / (rate * 60));
      var seconds = Math.floor((totalFrames % (rate * 60)) / rate);
      var frames = totalFrames % rate;
      if (isDrop)
        return (
          pad2(hours) +
          ":" +
          pad2(minutes) +
          ":" +
          pad2(seconds) +
          ";" +
          pad2(frames)
        );
      else
        return (
          pad2(hours) +
          ":" +
          pad2(minutes) +
          ":" +
          pad2(seconds) +
          ":" +
          pad2(frames)
        );
    } catch (e) {
      writeLn("Timecode calculation error: " + e.toString());
      return "TC Error";
    }
  }
  function padDigits(num, digits) {
    var s = String(Math.max(0, Math.round(num)));
    while (s.length < digits) s = "0" + s;
    return s;
  }

  // --- スクリプト実行開始 ---
  initializeLogFile();
  writeLn("Starting script: " + SCRIPT_NAME + " " + SCRIPT_VERSION);
  var myPanel = buildUI(thisObj);
})(this);
