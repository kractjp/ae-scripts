// ================================================
// KRACT Check Information AutoLayer
// Version 2.31 (Folder Management Update)
// ================================================

(function (thisObj) {
  // --- グローバル設定 ---
  var SCRIPT_NAME = "Check Information AutoLayer";
  var SCRIPT_VERSION = "2.31_folderMgmt"; // バージョン
  var FONT_POSTSCRIPT_NAME = "GeistMono-Medium";
  var FONT_SIZE = 20;
  var PRECOMP_PREFIX = "_check-data_";
  var PRECOMP_FOLDER_NAME = "_check-data"; // プリコンポーズとロゴの親フォルダ
  var FOOTAGE_FOLDER_NAME = "footage"; // ロゴフッテージを格納するサブフォルダ名
  var SETTINGS_SUBDIR = "settings";
  var SETTINGS_SUFFIX = ".config.json";
  var LOGO_TARGET_FILENAME = "checkinfo-logo.ai"; // 管理ディレクトリ内のロゴファイル名
  var LOGO_FOOTAGE_NAME = "checkinfo-logo.ai"; // AEプロジェクト内のフッテージ名
  var LOGO_LAYER_NAME = "logo";
  var LOGO_POSITION = [960, 1131.75];
  var LOGO_SCALE = [16, 16];

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

  var currentConfig = {
    project_name: "",
    logo_path: "",
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
      "プロジェクト名 (未入力時はファイル名):"
    );
    var editProjectName = panelProject.add("edittext", [0, 0, 280, 20], "");
    editProjectName.helpTip = "表示するプロジェクト名。設定があれば自動入力。";
    editProjectName.onChange = function () {
      currentConfig.project_name = this.text;
    };

    var panelLogo = myPanel.add("panel", undefined, "ロゴ設定");
    panelLogo.orientation = "column";
    panelLogo.alignChildren = "fill";
    panelLogo.spacing = 5;
    panelLogo.margins = 10;
    panelLogo.add("statictext", undefined, "使用するロゴファイル:");
    var grpLogoPath = panelLogo.add("group");
    grpLogoPath.orientation = "row";
    grpLogoPath.alignChildren = ["fill", "center"];
    var stLogoPath = grpLogoPath.add("statictext", [0, 0, 200, 20], "未設定", {
      truncate: "middle",
    });
    stLogoPath.helpTip = "現在設定されているロゴファイル (管理下のパス)";
    var btnBrowseLogo = grpLogoPath.add("button", undefined, "参照...");
    btnBrowseLogo.helpTip = "ロゴファイル (AI形式) を選択してください";

    btnBrowseLogo.onClick = function () {
      if (!hasSavedProject()) {
        alert(
          "プロジェクトが保存されていません。先にプロジェクトを保存してください。",
          SCRIPT_NAME
        );
        return;
      }
      var selectedFile = File.openDialog(
        "ロゴファイルを選択 (.ai)",
        "Adobe Illustrator:*.ai",
        false
      );
      if (selectedFile) {
        try {
          var managedPath = setAndCopyLogo(selectedFile);
          if (managedPath) {
            currentConfig.logo_path = managedPath;
            stLogoPath.text = File(managedPath).name + " (設定済)";
            saveConfig(currentConfig);
            alert(
              "ロゴを設定し、管理フォルダにコピーしました:\n" + managedPath
            );
          }
        } catch (e) {
          alert(
            "ロゴの設定またはコピー中にエラーが発生しました:\n" + e.toString(),
            SCRIPT_NAME
          );
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
      var currentAvSources = getAvailableSources(getActiveComp());
      if (!hasSavedProject()) {
        alert("プロジェクトが保存されていません。", SCRIPT_NAME);
        return;
      }
      try {
        var loaded = loadConfig();
        if (loaded) currentConfig = loaded;
        editProjectName.text = currentConfig.project_name || "";
        stLogoPath.text = currentConfig.logo_path
          ? File(currentConfig.logo_path).name + " (設定済)"
          : "未設定";
      } catch (e) {
        alert("実行前に設定を再読み込みできませんでした:\n" + e.toString());
      }

      var ac = getActiveComp();
      if (!ac) {
        alert("アクティブなコンポジションがありません。", SCRIPT_NAME);
        return;
      }
      if (!currentConfig.logo_path || !File(currentConfig.logo_path).exists) {
        alert(
          "ロゴファイルが設定されていないか、指定されたパスにファイルが存在しません。\nロゴ設定パネルでロゴファイルを設定してください。",
          SCRIPT_NAME
        );
        return;
      }
      if (
        ddSource.items.length === 0 ||
        ddSource.selection == null ||
        ddSource.selection.text === "ソースなし"
      ) {
        alert("有効なソースが選択されていません。", SCRIPT_NAME);
        return;
      }

      var selIndex = ddSource.selection.index;
      if (
        !currentAvSources ||
        selIndex >= currentAvSources.length ||
        currentAvSources[selIndex].name !== ddSource.selection.text
      ) {
        alert(
          "ソースリストが古いか、選択が無効です。更新ボタンを押して再試行してください。",
          SCRIPT_NAME
        );
        return;
      }
      var pickedObj = currentAvSources[selIndex];

      if (!pickedObj || !pickedObj.source) {
        alert("選択されたソース情報の取得に失敗しました。", SCRIPT_NAME);
        return;
      }

      var projectFileNameForLayer =
        currentConfig.project_name !== ""
          ? currentConfig.project_name
          : getAEProjectName();

      try {
        saveConfig(currentConfig);
      } catch (e) {
        alert("設定の保存に失敗しました:\n" + e.toString(), SCRIPT_NAME);
      }

      app.beginUndoGroup(SCRIPT_NAME + ": 実行");
      try {
        processComposition(
          ac,
          pickedObj,
          projectFileNameForLayer,
          currentConfig.logo_path
        );
        alert("処理が完了しました。", SCRIPT_NAME);
      } catch (e) {
        alert(
          "エラーが発生しました:\n" +
            e.toString() +
            (e.line ? "\nLine: " + e.line : ""),
          SCRIPT_NAME
        );
      } finally {
        app.endUndoGroup();
      }
    };

    function initializePanel() {
      avSources = refreshSourceList(ddSource);
      try {
        var loadedSettings = loadConfig();
        if (loadedSettings) {
          currentConfig = loadedSettings;
          editProjectName.text = currentConfig.project_name || "";
          if (currentConfig.logo_path && File(currentConfig.logo_path).exists) {
            stLogoPath.text = File(currentConfig.logo_path).name + " (設定済)";
          } else {
            stLogoPath.text = "未設定";
            currentConfig.logo_path = "";
          }
        } else {
          stLogoPath.text = "未設定";
        }
      } catch (e) {
        alert(
          "設定の読み込み中にエラーが発生しました:\n" + e.toString(),
          SCRIPT_NAME
        );
        stLogoPath.text = "読込エラー";
      }
      btnBrowseLogo.enabled = hasSavedProject();
    }

    function getAvailableSources(comp) {
      var sources = [];
      if (!comp) return sources;
      for (var i = 1; i <= comp.numLayers; i++) {
        var lyr = comp.layer(i);
        if (lyr instanceof AVLayer && lyr.source && lyr.source !== null) {
          var s = lyr.source;
          if (s instanceof CompItem && s.name.indexOf(PRECOMP_PREFIX) !== 0) {
            sources.push({ source: s, type: "comp", name: s.name });
          } else if (s instanceof FootageItem && s.name !== LOGO_FOOTAGE_NAME) {
            if (s.mainSource && !s.mainSource.isStill) {
              sources.push({ source: s, type: "footage", name: s.name });
            }
          }
        }
      }
      return sources;
    }

    function refreshSourceList(dropdown) {
      dropdown.removeAll();
      var currentSources = [];
      var ac = getActiveComp();
      var enableUI = hasSavedProject() && ac;
      btnBrowseLogo.enabled = enableUI;
      ddSource.enabled = enableUI;
      btnApply.enabled = enableUI;

      if (!enableUI) {
        dropdown.add("item", "プロジェクト未保存 または コンポなし");
        dropdown.selection = 0;
      } else {
        currentSources = getAvailableSources(ac);
        if (currentSources.length === 0) {
          dropdown.add("item", "ソースなし");
          dropdown.selection = 0;
          dropdown.enabled = false;
          btnApply.enabled = false;
        } else {
          for (var k = 0; k < currentSources.length; k++) {
            dropdown.add("item", currentSources[k].name);
          }
          dropdown.selection = 0;
          dropdown.enabled = true;
          btnApply.enabled = true;
        }
      }
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

  function setAndCopyLogo(sourceFile) {
    if (!sourceFile || !sourceFile.exists) {
      throw new Error("指定されたロゴファイルが見つかりません。");
    }
    if (!/\.ai$/i.test(sourceFile.name)) {
      throw new Error(
        "ロゴファイルは Adobe Illustrator (.ai) 形式である必要があります。"
      );
    }
    var proj = app.project;
    if (!proj || !proj.file) {
      throw new Error("プロジェクトが保存されていません。");
    }
    var projDir = proj.file.parent;
    var settingsDir = Folder(projDir.fsName + "/" + SETTINGS_SUBDIR);

    if (!settingsDir.exists) {
      if (!settingsDir.create()) {
        throw new Error(
          "設定ディレクトリの作成に失敗しました: " + settingsDir.fsName
        );
      }
      writeLn("設定ディレクトリを作成しました: " + settingsDir.fsName);
    }

    var targetFile = File(settingsDir.fsName + "/" + LOGO_TARGET_FILENAME);

    if (sourceFile.copy(targetFile.fsName)) {
      writeLn(
        "ロゴファイルをコピーしました: " +
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
  function processComposition(
    activeComp,
    pickedObj,
    projectFileName,
    managedLogoPath
  ) {
    if (!activeComp) return;
    if (!pickedObj || !pickedObj.source) {
      throw new Error("処理に必要なソース情報がありません。");
    }
    if (!managedLogoPath || !File(managedLogoPath).exists) {
      throw new Error(
        "管理下のロゴファイルパスが無効か、ファイルが存在しません:\n" +
          managedLogoPath
      );
    }

    var sourceItem = pickedObj.source;
    var sourceLayerInfo = { index: -1, name: "" };

    // --- 0. 既存レイヤーの情報を保持し、削除 ---
    writeLn("--- 既存レイヤー検索・削除開始 ---");
    var foundAndRemovedSource = false;
    for (var i = activeComp.numLayers; i >= 1; i--) {
      var currentLayer = activeComp.layer(i);
      if (
        LAYER_SPECS[currentLayer.name] ||
        currentLayer.name === LOGO_LAYER_NAME ||
        (currentLayer.source &&
          currentLayer.source instanceof CompItem &&
          currentLayer.source.name.indexOf(PRECOMP_PREFIX) === 0) // プリコンプレイヤーの判定を修正
      ) {
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
      writeLn("既存のソースレイヤーは見つかりませんでした。");
    writeLn("--- 既存レイヤー検索・削除終了 ---");

    // --- 0.5 コンポジションサイズ変更 ---
    if (activeComp.width === 1920 && activeComp.height === 1080) {
      activeComp.height = 1150;
      writeLn("コンポジションの高さを 1150px に変更しました。");
    }

    // --- 1. ロゴフッテージの準備 (インポートまたは置換) とフォルダ移動 ---
    writeLn("--- ロゴフッテージ準備開始 ---");
    var logoFootage = null;
    var managedLogoFile = File(managedLogoPath);

    try {
      var existingLogoFootage = findItemByName(LOGO_FOOTAGE_NAME);

      if (existingLogoFootage && existingLogoFootage instanceof FootageItem) {
        writeLn(
          "既存のロゴフッテージ '" +
            LOGO_FOOTAGE_NAME +
            "' を発見。ソースを置換します。"
        );
        // 既存のフッテージが正しいフォルダにあるか確認
        var expectedParentPath =
          PRECOMP_FOLDER_NAME + "/" + FOOTAGE_FOLDER_NAME;
        var currentParentFolder = getItemFolderPath(existingLogoFootage);
        if (currentParentFolder !== expectedParentPath) {
          writeLn(
            "  フッテージが期待されるフォルダ (" +
              expectedParentPath +
              ") にありません。移動します。"
          );
          // この後の処理で移動させるのでここではログのみ
        } else {
          writeLn(
            "  フッテージは既に正しいフォルダ (" +
              expectedParentPath +
              ") にあります。"
          );
        }
        existingLogoFootage.replaceSource(managedLogoFile, false);
        logoFootage = existingLogoFootage;
        writeLn("  フッテージソースを置換しました: " + managedLogoFile.fsName);
      } else {
        // ★★★★★★★ 要件3: 既存フッテージがない場合のみインポート ★★★★★★★
        writeLn(
          "ロゴフッテージ '" +
            LOGO_FOOTAGE_NAME +
            "' が見つからないため、新規にインポートします。"
        );
        var importOptions = new ImportOptions(managedLogoFile);
        logoFootage = app.project.importFile(importOptions);
        if (!logoFootage) {
          throw new Error("ロゴファイルのインポートに失敗しました。");
        }
        logoFootage.name = LOGO_FOOTAGE_NAME;
        writeLn(
          "  ロゴをインポートし、名前を '" +
            LOGO_FOOTAGE_NAME +
            "' に設定しました。"
        );
      }
    } catch (e) {
      throw new Error(
        "ロゴフッテージの準備（インポート/置換）中にエラーが発生しました:\n" +
          e.toString()
      );
    }

    if (!logoFootage || !(logoFootage instanceof FootageItem)) {
      throw new Error(
        "ロゴフッテージの準備に失敗しました。有効なフッテージアイテムを取得できませんでした。"
      );
    }

    // ★★★★★★★ 要件2: ロゴフッテージを _check-data/footage フォルダに移動 ★★★★★★★
    try {
      var checkDataFolder = findOrCreateFolder(PRECOMP_FOLDER_NAME);
      if (!checkDataFolder)
        throw new Error(
          "'" + PRECOMP_FOLDER_NAME + "' フォルダの取得/作成に失敗。"
        );

      var footageFolder = findOrCreateFolder(
        FOOTAGE_FOLDER_NAME,
        checkDataFolder
      ); // 親フォルダを指定
      if (!footageFolder)
        throw new Error(
          "'" + FOOTAGE_FOLDER_NAME + "' フォルダの取得/作成に失敗。"
        );

      // 既に正しいフォルダにあるか再確認（replaceSourceの場合も考慮）
      if (logoFootage.parentFolder !== footageFolder) {
        logoFootage.parentFolder = footageFolder;
        writeLn(
          "  ロゴフッテージを '" +
            PRECOMP_FOLDER_NAME +
            "/" +
            FOOTAGE_FOLDER_NAME +
            "' フォルダに移動しました。"
        );
      }
    } catch (e) {
      writeLn("警告: ロゴフッテージのフォルダ移動中にエラー: " + e.toString());
      // エラーが発生しても処理は続行させる（ルートに残るだけ）
    }
    writeLn("--- ロゴフッテージ準備完了 ---");

    // --- 2. テキストレイヤーの生成・更新 ---
    var textLayers = {};
    var textLayerNames = [];
    var layerName;
    for (layerName in LAYER_SPECS) {
      if (!LAYER_SPECS.hasOwnProperty(layerName)) continue;
      var spec = LAYER_SPECS[layerName];
      textLayers[layerName] = createTextLayer(activeComp, layerName, spec);
      if (!textLayers[layerName]) {
        throw new Error("テキストレイヤー '" + layerName + "' の作成失敗。");
      }
      textLayerNames.push(layerName);
      var targetLayer = textLayers[layerName];
      if (targetLayer.property("Position")) {
        targetLayer.property("Position").setValue(spec.pos);
      }
      if (targetLayer.property("Anchor Point")) {
        targetLayer.property("Anchor Point").setValue(spec.ap);
      }
      if (spec.expression && targetLayer.property("Source Text")) {
        try {
          targetLayer.property("Source Text").expression = spec.expression;
        } catch (e) {
          throw new Error(
            "レイヤー '" +
              layerName +
              "' のエクスプレッション設定失敗。\n" +
              e.toString()
          );
        }
      }
    }
    setTextLayerContents(activeComp, textLayers, pickedObj, projectFileName);

    // --- 3. ロゴレイヤーをアクティブコンポに追加 ---
    writeLn("--- ロゴレイヤー追加開始 ---");
    var logoLayer = null;
    try {
      logoLayer = activeComp.layers.add(logoFootage);
      if (!logoLayer) throw new Error("ロゴレイヤーの追加に失敗しました。");
      logoLayer.name = LOGO_LAYER_NAME;
      writeLn("ロゴレイヤーを追加しました: " + logoLayer.name);
    } catch (e) {
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
    var precompName = PRECOMP_PREFIX + activeComp.name;
    var layersToPrecomposeIndices = [];
    for (var i = 1; i <= activeComp.numLayers; i++) {
      var layer = activeComp.layer(i);
      // textLayersオブジェクトにキーとして存在するか、またはロゴレイヤーか
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
        // 既存の同名プリコンポーズを検索して削除（必要に応じて）
        var existingPreComp = findItemByName(precompName);
        if (existingPreComp && existingPreComp instanceof CompItem) {
          writeLn("既存のプリコンポーズ '" + precompName + "' を削除します。");
          try {
            existingPreComp.remove();
          } catch (removeError) {
            writeLn(
              "警告: 既存プリコンポーズの削除に失敗: " + removeError.toString()
            );
            // 続行を試みる
          }
        }

        newComp = activeComp.layers.precompose(
          layersToPrecomposeIndices,
          precompName,
          true // move all attributes
        );
        if (newComp) {
          writeLn("プリコンポーズ成功 (テキスト+ロゴ): " + newComp.name);
          precompLayer = findLayerByName(activeComp, precompName); // プリコンポーズされたレイヤーを取得

          // ★★★★★★★ 要件1: プリコンポーズを _check-data フォルダに移動 ★★★★★★★
          try {
            var precompTargetFolder = findOrCreateFolder(PRECOMP_FOLDER_NAME);
            if (!precompTargetFolder)
              throw new Error(
                "'" + PRECOMP_FOLDER_NAME + "' フォルダの取得/作成に失敗。"
              );

            if (newComp.parentFolder !== precompTargetFolder) {
              newComp.parentFolder = precompTargetFolder;
              writeLn(
                "  プリコンポーズ '" +
                  newComp.name +
                  "' を '" +
                  PRECOMP_FOLDER_NAME +
                  "' フォルダに移動しました。"
              );
            } else {
              writeLn(
                "  プリコンポーズ '" +
                  newComp.name +
                  "' は既に '" +
                  PRECOMP_FOLDER_NAME +
                  "' フォルダにあります。"
              );
            }
          } catch (folderError) {
            writeLn(
              "警告: プリコンポーズのフォルダ移動中にエラー: " +
                folderError.toString()
            );
            // エラーが発生しても処理は続行（ルートに残る）
          }
        } else {
          throw new Error("プリコンポーズ作成失敗(null)。");
        }
      } catch (e) {
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

    // --- 6. ソースレイヤーを再度追加し、設定 ---
    writeLn("--- ソースレイヤー再追加処理開始 ---");
    var reAddedSourceLayer = null;
    try {
      reAddedSourceLayer = activeComp.layers.add(sourceItem);
      if (!reAddedSourceLayer || !reAddedSourceLayer.isValid)
        throw new Error("ソースレイヤーの再追加失敗。");
      writeLn("ソースレイヤーを再追加しました: " + reAddedSourceLayer.name);
      if (sourceLayerInfo.name !== "") {
        reAddedSourceLayer.name = sourceLayerInfo.name;
        writeLn("  名前を '" + sourceLayerInfo.name + "' に再設定。");
      }
      var posProp = reAddedSourceLayer.property("Position");
      if (posProp && posProp.canSetValue) {
        var currentPos = posProp.value;
        posProp.setValue([currentPos[0], 575]); // 1150 / 2 = 575 (Yセンター想定)
        writeLn("  Y座標を 575 に設定。");
      } else {
        writeLn("警告: 再追加したソースレイヤーの位置を設定できません。");
      }

      // ソースレイヤーをプリコンポーズの下に移動
      if (precompLayer && precompLayer.isValid) {
        reAddedSourceLayer.moveAfter(precompLayer);
        writeLn("  レイヤー順序をプリコンポーズの下に移動。");
      } else {
        reAddedSourceLayer.moveToEnd();
        writeLn("  レイヤー順序を最下層に移動。");
      }
    } catch (e) {
      writeLn(
        "警告: ソースレイヤーの再追加または設定中にエラー発生:\n" + e.toString()
      );
    }
    writeLn("--- ソースレイヤー再追加処理終了 ---");
  } // processComposition 終了

  // --- テキストレイヤー関連ヘルパー ---
  function createTextLayer(comp, layerName, spec) {
    var textLayer;
    try {
      textLayer = comp.layers.addText("");
      textLayer.name = layerName;
    } catch (e) {
      throw new Error(
        "テキストレイヤー追加失敗 ('" + layerName + "')。\n" + e.toString()
      );
    }
    var textProp = textLayer.property("Source Text");
    if (!textProp) {
      throw new Error("Source Textプロパティ取得失敗 ('" + layerName + "')。");
    }
    var textDocument = textProp.value;
    try {
      textDocument.font = FONT_POSTSCRIPT_NAME;
      textDocument.fontSize = FONT_SIZE;
      if (spec.justification != null) {
        textDocument.justification = spec.justification;
      } else {
        textDocument.justification = ParagraphJustification.LEFT_JUSTIFY;
      }
      textProp.setValue(textDocument);
    } catch (e) {
      alert(
        "フォント/揃え設定失敗 ('" +
          layerName +
          "')。\nFont: '" +
          FONT_POSTSCRIPT_NAME +
          "'\n" +
          e.toString() +
          "\nフォントがインストールされているか確認してください。",
        SCRIPT_NAME
      );
      try {
        textLayer.remove();
      } catch (err) {}
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
    var layerName;
    layerName = "current_info";
    if (
      textLayers[layerName] &&
      textLayers[layerName].property("Source Text") &&
      !textLayers[layerName].property("Source Text").expressionEnabled
    ) {
      textLayers[layerName]
        .property("Source Text")
        .setValue(createDateTimeString());
    }
    layerName = "comp_name";
    if (
      textLayers[layerName] &&
      textLayers[layerName].property("Source Text") &&
      !textLayers[layerName].property("Source Text").expressionEnabled
    ) {
      textLayers[layerName].property("Source Text").setValue(activeComp.name);
    }
    layerName = "project_name";
    if (
      textLayers[layerName] &&
      textLayers[layerName].property("Source Text") &&
      !textLayers[layerName].property("Source Text").expressionEnabled
    ) {
      textLayers[layerName].property("Source Text").setValue(projectFileName);
    }
    layerName = "comp_info";
    if (
      textLayers[layerName] &&
      textLayers[layerName].property("Source Text") &&
      !textLayers[layerName].property("Source Text").expressionEnabled
    ) {
      var compInfoString = "N/A";
      try {
        if (pickedObj && pickedObj.source) {
          var sourceItem = pickedObj.source;
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
          if (pickedObj.type === "comp") {
            compInfoString = "TC/Fm: " + tc + "/" + framesPadded;
          } else {
            if (sourceName.length > 30) {
              sourceName = sourceName.substring(0, 27) + "...";
            }
            compInfoString = sourceName + " TC/Fm: " + tc + "/" + framesPadded;
          }
        }
      } catch (e) {
        compInfoString = "Error getting info";
      }
      textLayers[layerName].property("Source Text").setValue(compInfoString);
    }
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
    var searchFolder = parentFolder || proj.rootFolder; // 指定がなければルートから検索

    for (var i = 1; i <= searchFolder.numItems; i++) {
      var item = searchFolder.item(i);
      if (item.name === name) {
        return item;
      }
      // サブフォルダ内も再帰的に検索する場合 (今回は不要)
      // if (item instanceof FolderItem) {
      //   var found = findItemByName(name, item);
      //   if (found) return found;
      // }
    }
    // ルート直下も検索（parentFolder指定時） - 必要に応じて
    if (parentFolder && parentFolder !== proj.rootFolder) {
      for (var i = 1; i <= proj.rootFolder.numItems; i++) {
        var item = proj.rootFolder.item(i);
        if (item.name === name && !(item instanceof FolderItem)) {
          // ルートのフォルダは除く
          return item;
        }
      }
    }

    return null;
  }

  // 指定したフォルダ内に、指定した名前のフォルダを探し、なければ作成する関数
  // parentFolder を指定可能に修正
  function findOrCreateFolder(folderName, parentFolder) {
    var proj = app.project;
    if (!proj) return null;
    var targetParent = parentFolder || proj.rootFolder; // 親フォルダ指定がなければルート

    // 既存フォルダを検索
    for (var i = 1; i <= targetParent.numItems; i++) {
      var item = targetParent.item(i);
      if (item instanceof FolderItem && item.name === folderName) {
        return item; // 見つかったら返す
      }
    }
    // 見つからなければ作成
    try {
      return targetParent.items.addFolder(folderName);
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

  // アイテムのフォルダパスを文字列で取得する（デバッグや確認用）
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

  // --- ファイルI/Oヘルパー関数 ---
  function getAEProjectName() {
    if (!hasSavedProject()) return "Untitled";
    var file = app.project.file;
    if (file) {
      try {
        var nm = decodeURI(file.name);
        return nm.replace(/\.aep$/i, "");
      } catch (e) {
        return file.name.replace(/\.aep$/i, "");
      }
    }
    return "Untitled";
  }

  function getConfigFilePath() {
    if (!app.project || !app.project.file) {
      return null;
    }
    var projPath = app.project.file.path;
    var projFileName = getAEProjectName();
    if (projPath === "" || projFileName === "Untitled") {
      return null;
    }
    var settingsFolderPath = projPath + "/" + SETTINGS_SUBDIR;
    var configFilePath =
      settingsFolderPath + "/" + projFileName + SETTINGS_SUFFIX;
    return configFilePath;
  }

  function saveConfig(configData) {
    var filePath = getConfigFilePath();
    if (!filePath) {
      throw new Error("プロジェクト未保存/パス無効のため設定保存不可。");
    }
    var settingsFolder = Folder(Folder(filePath).parent);
    if (!settingsFolder.exists) {
      if (!settingsFolder.create()) {
        throw new Error(
          "settings ディレクトリ作成失敗: " + settingsFolder.fsName
        );
      }
    }
    if (typeof configData !== "object" || configData === null) {
      throw new Error("保存する設定データが無効です。");
    }
    var jsonString = JSON.stringify(configData, null, 2);
    var configFile = File(filePath);
    configFile.encoding = "UTF-8";
    if (!configFile.open("w")) {
      throw new Error("設定ファイルを開けませんでした(w): " + configFile.error);
    }
    try {
      if (!configFile.write(jsonString)) {
        throw new Error("設定ファイル書き込み失敗: " + configFile.error);
      }
      writeLn("設定を保存しました: " + filePath);
    } finally {
      configFile.close();
    }
  }

  function loadConfig() {
    var filePath = getConfigFilePath();
    if (!filePath) {
      return null;
    }
    var configFile = File(filePath);
    if (!configFile.exists) {
      return null;
    }
    configFile.encoding = "UTF-8";
    if (!configFile.open("r")) {
      throw new Error("設定ファイルを開けませんでした(r): " + configFile.error);
    }
    var jsonString = "";
    try {
      jsonString = configFile.read();
    } finally {
      configFile.close();
    }
    if (jsonString === "") {
      return null;
    }
    try {
      var configData = JSON.parse(jsonString);
      var validConfig = {
        project_name:
          configData && typeof configData.project_name === "string"
            ? configData.project_name
            : "",
        logo_path:
          configData && typeof configData.logo_path === "string"
            ? configData.logo_path
            : "",
      };
      if (validConfig.logo_path && !File(validConfig.logo_path).exists) {
        writeLn(
          "警告: 設定ファイルに記載されたロゴパスが見つかりません: " +
            validConfig.logo_path
        );
        // validConfig.logo_path = ""; // パスが存在しない場合はクリアする方が安全かもしれない
      }
      writeLn("設定を読み込みました: " + filePath);
      return validConfig;
    } catch (e) {
      throw new Error(
        "設定ファイル(JSON)解析失敗:\n" +
          e.toString() +
          "\nFile: " +
          configFile.fsName
      );
    }
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
    function padOffset(v) {
      return v < 10 ? "0" + v : v;
    }
    var offsetStr = sign + padOffset(offsetH) + ":" + padOffset(offsetM);
    return y + "." + m + "." + day + "T" + h + ":" + min + ":" + s + offsetStr;
  }
  function convertSecondsToTC(durationSec, fps) {
    if (isNaN(durationSec) || isNaN(fps) || fps <= 0 || durationSec < 0) {
      return "00:00:00:00";
    }
    var totalFrames = Math.round(durationSec * fps);
    var hours = Math.floor(totalFrames / (fps * 3600));
    var minutes = Math.floor((totalFrames % (fps * 3600)) / (fps * 60));
    var seconds = Math.floor((totalFrames % (fps * 60)) / fps);
    var frames =
      totalFrames - hours * fps * 3600 - minutes * fps * 60 - seconds * fps;
    frames = Math.max(0, Math.min(Math.round(frames), Math.round(fps) - 1));
    function pad2(n) {
      return n < 10 ? "0" + n : "" + n;
    }
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
  function padDigits(num, digits) {
    var s = String(Math.max(0, Math.round(num)));
    while (s.length < digits) {
      s = "0" + s;
    }
    return s;
  }
  function writeLn(message) {
    var currentTime = new Date().toLocaleTimeString();
    var output = "[" + currentTime + "] " + message;
    $.writeln(output);
    // UIにログ表示エリアがあれば追加
  }

  // --- スクリプト実行開始 ---
  var myPanel = buildUI(thisObj);
})(this);
