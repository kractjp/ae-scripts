// ================================================
// KRACT Check Information AutoLayer
// Version 2.24
// ================================================

(function (thisObj) {
  // --- グローバル設定 ---
  var SCRIPT_NAME = "Check Information AutoLayer";
  var SCRIPT_VERSION = "2.24_logoInPrecompTarget"; // バージョン（ロゴもプリコンポーズ対象に）

  var FONT_POSTSCRIPT_NAME = "GeistMono-Medium";
  var FONT_SIZE = 20;

  var PRECOMP_PREFIX = "_check-data_";
  var PRECOMP_FOLDER_NAME = "_check-data";

  var SETTINGS_SUBDIR = "settings";
  var SETTINGS_SUFFIX = ".config.json";

  var LOGO_SUBDIR = "Adobe/After Effects/CheckInfoLogoData";
  var LOGO_FILE_NAME = "checkinfo-logo.ai";
  var LOGO_LAYER_NAME = "logo"; // コンポに追加するレイヤー名
  // var LOGO_TARGET_HEIGHT = 16; // スケール計算は不要
  var LOGO_POSITION = [960, 1131.75]; // ロゴの位置 (X, Y)

  // レイヤー設定
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
    ddSource.helpTip =
      "情報を表示する基準となるソースレイヤーを選択してください";
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
    editProjectName.helpTip =
      "テキストレイヤーに表示するプロジェクト名。設定ファイルがあれば自動入力されます。";
    var panelLogo = myPanel.add("panel", undefined, "ロゴ設定");
    panelLogo.orientation = "row";
    panelLogo.alignChildren = ["left", "center"];
    panelLogo.spacing = 10;
    panelLogo.margins = 10;
    var btnSelectLogo = panelLogo.add("button", undefined, "ロゴ選択＆設定...");
    btnSelectLogo.helpTip =
      "ロゴファイルを選択し、ユーザーデータフォルダ内の指定場所に'" +
      LOGO_FILE_NAME +
      "'としてコピー・設定します";
    var stSelectedLogo = panelLogo.add("statictext", undefined, "...", {
      truncate: "middle",
    });
    stSelectedLogo.helpTip =
      "現在のロゴ設定状態 (" +
      Folder.userData.fsName +
      "/" +
      LOGO_SUBDIR +
      "/" +
      LOGO_FILE_NAME +
      ")";
    stSelectedLogo.preferredSize.width = 180;
    var grpBtns = myPanel.add("group");
    grpBtns.orientation = "row";
    grpBtns.alignment = "center";
    grpBtns.spacing = 10;
    var btnRefresh = grpBtns.add("button", undefined, "更新");
    btnRefresh.helpTip = "ソース一覧を更新します";
    var btnApply = grpBtns.add("button", undefined, "実行");
    btnApply.helpTip =
      "テキスト/ロゴ生成・設定、設定保存、プリコンポーズ、ソース再配置(Y575)"; // Help Tip 更新
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
      var ac = getActiveComp();
      if (!ac) {
        alert("アクティブなコンポジションがありません。", SCRIPT_NAME);
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
        alert("ソース情報の不一致、または無効な選択です。", SCRIPT_NAME);
        avSources = refreshSourceList(ddSource);
        return;
      }
      var pickedObj = currentAvSources[selIndex];
      if (!pickedObj || !pickedObj.source) {
        alert("選択されたソース情報の取得に失敗しました。", SCRIPT_NAME);
        return;
      }
      var inputProjectName = editProjectName.text;
      var projectFileNameForLayer =
        inputProjectName !== "" ? inputProjectName : getAEProjectName();
      if (inputProjectName !== "") {
        try {
          saveProjectNameConfig(inputProjectName);
        } catch (e) {
          alert(
            "プロジェクト名設定の保存に失敗しました:\n" + e.toString(),
            SCRIPT_NAME
          );
        }
      }
      app.beginUndoGroup(SCRIPT_NAME + ": 実行 (ロゴもプリコンポーズ)"); // Undo Group名変更
      try {
        processComposition(ac, pickedObj, projectFileNameForLayer);
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
    btnSelectLogo.onClick = function () {
      try {
        var userSelectedFile = File.openDialog(
          "ロゴファイルを選択してください",
          "Illustrator AI:*.ai;All Files:*.*",
          false
        );
        if (userSelectedFile) {
          var targetFolderPath = getLogoFolderPath();
          var targetFilePath = targetFolderPath + "/" + LOGO_FILE_NAME;
          if (!targetFolderPath) {
            throw new Error(
              "スクリプトが保存されていないため、ロゴを格納できません。"
            );
          }
          var targetFolder = Folder(targetFolderPath);
          if (!targetFolder.exists) {
            if (!targetFolder.create()) {
              throw new Error(
                "ロゴ格納フォルダの作成に失敗しました:\n" + targetFolderPath
              );
            }
          }
          if (userSelectedFile.copy(targetFilePath)) {
            stSelectedLogo.text = LOGO_FILE_NAME + " (設定済)";
            alert(
              "ロゴを設定しました。\n指定のユーザーデータフォルダ内に'" +
                LOGO_FILE_NAME +
                "'としてコピーされました。",
              SCRIPT_NAME
            );
          } else {
            throw new Error(
              "ロゴファイルのコピーに失敗しました。\nコピー元: " +
                userSelectedFile.fsName +
                "\nコピー先: " +
                targetFilePath
            );
          }
        } else {
          updateLogoStatusUI(stSelectedLogo);
        }
      } catch (e) {
        alert(
          "ロゴ設定中にエラーが発生しました:\n" + e.toString(),
          SCRIPT_NAME
        );
        updateLogoStatusUI(stSelectedLogo);
      }
    };
    function initializePanel() {
      avSources = refreshSourceList(ddSource);
      try {
        var loadedName = loadProjectNameConfig();
        if (loadedName !== null && loadedName !== "") {
          editProjectName.text = loadedName;
        }
      } catch (e) {
        alert(
          "プロジェクト名設定の読み込み中にエラーが発生しました:\n" +
            e.toString(),
          SCRIPT_NAME
        );
      }
      updateLogoStatusUI(stSelectedLogo);
    }
    function updateLogoStatusUI(statusTextElement) {
      var logoFilePath = getLogoFilePath(true);
      var logoFolder = Folder(getLogoFolderPath(true));
      if (logoFilePath && logoFolder) {
        var logoFile = File(logoFilePath);
        if (logoFile.exists) {
          statusTextElement.text = LOGO_FILE_NAME + " (設定済)";
          statusTextElement.helpTip =
            "ロゴファイルが見つかりました:\n" + logoFile.fsName;
        } else {
          statusTextElement.text = "ロゴ未設定";
          statusTextElement.helpTip =
            "ロゴファイルが見つかりません:\n" + logoFilePath;
        }
      } else {
        statusTextElement.text = "ロゴ未設定 (パス取得不可)";
        statusTextElement.helpTip =
          "ロゴの格納先フォルダパスを取得できませんでした。";
      }
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
          } else if (s instanceof FootageItem) {
            var isLogo = false;
            var logoPath = getLogoFilePath(true);
            if (logoPath) {
              var logoFile = File(logoPath);
              try {
                var itemFilePath =
                  s.mainSource && s.mainSource.file
                    ? decodeURI(s.mainSource.file.fsName)
                    : s.file
                    ? decodeURI(s.file.fsName)
                    : null;
                if (
                  itemFilePath &&
                  logoFile.fsName === File(itemFilePath).fsName
                ) {
                  isLogo = true;
                }
              } catch (e) {}
            }
            if (!isLogo && s.mainSource && !s.mainSource.isStill) {
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
      if (!hasSavedProject() || !ac) {
        dropdown.add("item", "ソースなし");
        dropdown.selection = 0;
        dropdown.enabled = false;
      } else {
        currentSources = getAvailableSources(ac);
        if (currentSources.length === 0) {
          dropdown.add("item", "ソースなし");
          dropdown.selection = 0;
          dropdown.enabled = false;
        } else {
          for (var k = 0; k < currentSources.length; k++) {
            dropdown.add("item", currentSources[k].name);
          }
          dropdown.selection = 0;
          dropdown.enabled = true;
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
  } // buildUI 終了

  // --- メイン処理関数 (処理順序を変更) ---
  function processComposition(activeComp, pickedObj, projectFileName) {
    if (!activeComp) return;
    if (!pickedObj || !pickedObj.source) {
      throw new Error("処理に必要なソース情報がありません。");
    }

    var sourceItem = pickedObj.source; // 操作対象のソースアイテム
    var sourceLayerInfo = { index: -1, name: "" }; // 削除前のソースレイヤー情報
    var logoFootageItem = null; // ロゴフッテージアイテム
    var addedLogoLayer = null; // アクティブコンポに追加されたロゴレイヤー

    // --- 0. 既存ソースレイヤーの情報を保持し、削除 ---
    writeLn("--- 既存ソースレイヤー検索・削除開始 ---");
    var foundAndRemoved = false;
    for (var i = activeComp.numLayers; i >= 1; i--) {
      var currentLayer = activeComp.layer(i);
      // 既存のテキストレイヤーやプリコンポーズ、ロゴレイヤーも削除対象にする
      if (
        LAYER_SPECS[currentLayer.name] ||
        currentLayer.name.indexOf(PRECOMP_PREFIX) === 0 ||
        currentLayer.name === LOGO_LAYER_NAME
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
              "' の削除に失敗。\n" +
              e.toString()
          );
        }
      }
      // ソースレイヤーの情報を保持して削除
      else if (
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
          if (!foundAndRemoved) {
            // 最初のソースレイヤー情報のみ保持
            sourceLayerInfo.index = currentLayer.index;
            sourceLayerInfo.name = currentLayer.name;
          }
          if (currentLayer.locked) currentLayer.locked = false;
          currentLayer.remove();
          writeLn("  ソースレイヤー削除完了。");
          foundAndRemoved = true;
        } catch (e) {
          writeLn(
            "警告: 既存ソースレイヤー '" +
              currentLayer.name +
              "' の削除に失敗。\n" +
              e.toString()
          );
        }
      }
    }
    if (!foundAndRemoved) {
      writeLn("既存のソースレイヤーは見つかりませんでした。");
    }
    writeLn("--- 既存レイヤー検索・削除終了 ---");

    // --- 0.5 コンポジションサイズ変更 ---
    if (activeComp.width === 1920 && activeComp.height === 1080) {
      activeComp.height = 1150;
      writeLn("コンポジションの高さを 1150px に変更しました。");
    }

    // --- 1. ロゴフッテージの準備 ---
    try {
      var logoFilePath = getLogoFilePath();
      if (logoFilePath) {
        var logoFile = File(logoFilePath);
        if (logoFile.exists) {
          logoFootageItem = findFootageByPath(logoFile.fsName);
          if (!logoFootageItem) {
            logoFootageItem = importLogoFile(logoFile);
          }
          if (logoFootageItem && logoFootageItem.isValid) {
            writeLn("ロゴフッテージを準備しました: " + logoFootageItem.name);
          } else {
            writeLn(
              "警告: ロゴファイルのインポート/検索に失敗、または無効なアイテムです。"
            );
            logoFootageItem = null;
          }
        } else {
          writeLn("設定されたロゴファイルが見つかりません: " + logoFile.fsName);
        }
      } else {
        writeLn("ロゴファイルのパスを取得できませんでした。");
      }
    } catch (e) {
      alert("ロゴ処理中にエラー:\n" + e.toString(), SCRIPT_NAME);
    }

    // --- 2. テキストレイヤーの生成・更新 ---
    var textLayers = {}; // 生成したテキストレイヤーを保持
    var textLayerNames = []; // レイヤー名を保持
    var layerName;
    for (layerName in LAYER_SPECS) {
      if (!LAYER_SPECS.hasOwnProperty(layerName)) continue;
      var spec = LAYER_SPECS[layerName];
      // 既存レイヤーはステップ0で削除済みのはずなので、常に新規作成
      textLayers[layerName] = createTextLayer(activeComp, layerName, spec);
      if (!textLayers[layerName]) {
        throw new Error(
          "テキストレイヤー '" + layerName + "' の作成に失敗しました。"
        );
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
    if (logoFootageItem && logoFootageItem.isValid) {
      writeLn(
        "アクティブコンポ '" + activeComp.name + "' にロゴを配置します..."
      );
      try {
        addedLogoLayer = activeComp.layers.add(logoFootageItem); // アクティブコンポに追加
        if (addedLogoLayer && addedLogoLayer.isValid) {
          addedLogoLayer.name = LOGO_LAYER_NAME; // 名前設定
          writeLn(
            "ロゴレイヤー '" +
              addedLogoLayer.name +
              "' をアクティブコンポに追加しました。"
          );
        } else {
          writeLn("警告: アクティブコンポへのロゴレイヤー追加に失敗しました。");
          addedLogoLayer = null;
        }
      } catch (e) {
        writeLn(
          "警告: アクティブコンポへのロゴレイヤー追加中にエラー:\n" +
            e.toString()
        );
        addedLogoLayer = null;
      }
    } else {
      writeLn(
        "ロゴフッテージが見つからないか無効なため、ロゴは配置されません。"
      );
    }

    // --- 4. 追加したロゴレイヤーのトランスフォーム設定 ---
    if (addedLogoLayer) {
      // ロゴレイヤーが正常に追加された場合のみ実行
      try {
        setLogoLayerTransform(addedLogoLayer, logoFootageItem); // トランスフォーム設定
        writeLn(
          "アクティブコンポ内のロゴレイヤーのトランスフォームを設定しました。"
        );
      } catch (e) {
        writeLn(
          "警告: アクティブコンポ内のロゴレイヤーのトランスフォーム設定中にエラー:\n" +
            e.toString()
        );
      }
    }

    // --- 5. プリコンポーズ処理 (テキストレイヤー + ロゴレイヤー) ---
    var precompName = PRECOMP_PREFIX + activeComp.name;
    // 既存プリコンポーズはステップ0で削除済み

    // プリコンポーズ対象レイヤー特定 (生成したテキストレイヤー + 追加したロゴレイヤー)
    var layersToPrecomposeIndices = [];
    for (var i = 1; i <= activeComp.numLayers; i++) {
      var layer = activeComp.layer(i);
      // textLayersオブジェクトに存在するレイヤーか、addedLogoLayerと同じレイヤーを対象とする
      if (
        textLayers[layer.name] ||
        (addedLogoLayer && layer.index === addedLogoLayer.index)
      ) {
        layersToPrecomposeIndices.push(layer.index); // レイヤーのインデックスを追加
      }
    }
    writeLn(
      "プリコンポーズ対象レイヤーインデックス: " +
        layersToPrecomposeIndices.join(", ")
    );

    var newComp = null; // 生成されたプリコンポーズアイテム
    var precompLayer = null; // 生成されたプリコンポーズレイヤー
    // プリコンポーズ実行
    if (layersToPrecomposeIndices.length > 0) {
      try {
        // moveAllAttributes=true でエフェクトなども移動
        newComp = activeComp.layers.precompose(
          layersToPrecomposeIndices,
          precompName,
          true
        );
        if (newComp) {
          writeLn("プリコンポーズ成功 (テキスト+ロゴ): " + newComp.name);
          precompLayer = findLayerByName(activeComp, precompName); // 生成されたレイヤーを取得
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
      reAddedSourceLayer = activeComp.layers.add(sourceItem); // 再度追加
      if (!reAddedSourceLayer || !reAddedSourceLayer.isValid)
        throw new Error("ソースレイヤーの再追加に失敗しました。");
      writeLn("ソースレイヤーを再追加しました: " + reAddedSourceLayer.name);

      if (sourceLayerInfo.name !== "") {
        reAddedSourceLayer.name = sourceLayerInfo.name;
        writeLn("  名前を '" + sourceLayerInfo.name + "' に再設定しました。");
      }

      var posProp = reAddedSourceLayer.property("Position");
      if (posProp && posProp.canSetValue) {
        var currentPos = posProp.value;
        posProp.setValue([currentPos[0], 575]);
        writeLn("  Y座標を 575 に設定しました。");
      } else {
        writeLn("警告: 再追加したソースレイヤーの位置を設定できません。");
      }

      // レイヤー順序調整 (プリコンポーズレイヤーがあればその下、なければ最下層)
      if (precompLayer && precompLayer.isValid) {
        reAddedSourceLayer.moveAfter(precompLayer);
        writeLn("  レイヤー順序をプリコンポーズの下に移動しました。");
      } else {
        reAddedSourceLayer.moveToEnd();
        writeLn("  レイヤー順序を最下層に移動しました。");
      }
    } catch (e) {
      writeLn(
        "警告: ソースレイヤーの再追加または設定中にエラーが発生しました。\n" +
          e.toString()
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
        "フォントまたはテキスト揃え設定失敗 ('" +
          layerName +
          "')。\nFont: '" +
          FONT_POSTSCRIPT_NAME +
          "'\n" +
          e.toString(),
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
  function findOrCreateFolder(folderName) {
    var proj = app.project;
    if (!proj) return null;
    for (var i = 1; i <= proj.numItems; i++) {
      var item = proj.item(i);
      if (item instanceof FolderItem && item.name === folderName) {
        return item;
      }
    }
    try {
      return proj.items.addFolder(folderName);
    } catch (e) {
      return null;
    }
  }
  function findItemByNameAndType(name, itemType) {
    var proj = app.project;
    if (!proj || typeof name !== "string" || name === "" || !itemType)
      return null;
    for (var i = 1; i <= proj.numItems; i++) {
      var item = proj.item(i);
      if (item instanceof itemType && item.name === name) {
        return item;
      }
    }
    return null;
  }

  // --- ロゴ関連ヘルパー関数 ---
  function getLogoFolderPath(suppressError) {
    var userDataFolder = Folder.userData;
    if (!userDataFolder || !userDataFolder.exists) {
      if (!suppressError) writeLn("ユーザーデータフォルダが見つかりません。");
      return null;
    }
    var logoFolder = Folder(userDataFolder.fsName + "/" + LOGO_SUBDIR);
    if (!logoFolder.exists) {
      if (!logoFolder.create()) {
        if (!suppressError)
          writeLn("ロゴフォルダの作成に失敗: " + logoFolder.fsName);
        return null;
      } else {
        if (!suppressError)
          writeLn("ロゴフォルダを作成しました: " + logoFolder.fsName);
      }
    }
    return logoFolder.fsName;
  }
  function getLogoFilePath(suppressError) {
    var logoFolderPath = getLogoFolderPath(suppressError);
    if (!logoFolderPath) return null;
    var logoFilePath = logoFolderPath + "/" + LOGO_FILE_NAME;
    return logoFilePath;
  }
  function findFootageByPath(filePath) {
    var proj = app.project;
    if (!proj || typeof filePath !== "string" || filePath === "") return null;
    var targetFile = File(decodeURI(filePath));
    if (!targetFile.exists) return null;

    for (var i = 1; i <= proj.numItems; i++) {
      var item = proj.item(i);
      if (item instanceof FootageItem) {
        var itemFile = null;
        try {
          if (item.mainSource && item.mainSource.file) {
            itemFile = item.mainSource.file;
          } else if (item.file) {
            itemFile = item.file;
          }
          if (itemFile && decodeURI(itemFile.fsName) === targetFile.fsName) {
            return item;
          }
        } catch (e) {}
      }
    }
    return null;
  }
  function importLogoFile(logoFile) {
    if (!(logoFile instanceof File) || !logoFile.exists) return null;
    try {
      writeLn("ロゴファイルをインポートします: " + logoFile.fsName);
      var importOptions = new ImportOptions(logoFile);
      if (importOptions.canImportAs(ImportAsType.FOOTAGE)) {
        importOptions.importAs = ImportAsType.FOOTAGE;
        writeLn("  ImportAsType.FOOTAGE を設定");
      } else {
        writeLn(
          "警告: フッテージとしてインポートできません。インポートタイプを確認してください。"
        );
        importOptions.sequence = false;
      }
      importOptions.sequence = false;
      var importedItem = app.project.importFile(importOptions);
      if (importedItem) {
        writeLn("インポート成功: " + importedItem.name);
      } else {
        writeLn("インポート失敗: importFile が null を返しました。");
      }
      return importedItem;
    } catch (e) {
      writeLn(
        "ロゴファイルのインポート中にエラー: " +
          logoFile.fsName +
          "\n" +
          e.toString()
      );
      return null;
    }
  }

  // ロゴレイヤーのトランスフォームを設定する関数 (変更なし)
  function setLogoLayerTransform(logoLayer, footageItem) {
    if (
      !logoLayer ||
      !logoLayer.isValid ||
      !footageItem ||
      !footageItem.isValid
    ) {
      writeLn("setLogoLayerTransform: 無効な引数です。");
      return;
    }
    writeLn(
      "ロゴレイヤー '" + logoLayer.name + "' のトランスフォーム設定開始..."
    );
    try {
      var scaleProp = logoLayer.property("Scale");
      var anchorPointProp = logoLayer.property("Anchor Point");
      var positionProp = logoLayer.property("Position");

      if (!scaleProp || !anchorPointProp || !positionProp) {
        throw new Error("トランスフォームプロパティが見つかりません。");
      }

      // 1. アンカーポイントをフッテージの中央に設定
      try {
        var anchorX = footageItem.width / 2;
        var anchorY = footageItem.height / 2;
        anchorPointProp.setValue([anchorX, anchorY]);
        writeLn("  アンカーポイント設定: [" + anchorX + ", " + anchorY + "]");
      } catch (apError) {
        writeLn("警告: アンカーポイント設定エラー: " + apError.toString());
      }

      // 2. 位置を設定 (LOGO_POSITION 変数を参照)
      try {
        positionProp.setValue(LOGO_POSITION);
        writeLn(
          "  位置設定: [" + LOGO_POSITION[0] + ", " + LOGO_POSITION[1] + "]"
        );
      } catch (posError) {
        writeLn("警告: 位置設定エラー: " + posError.toString());
      }

      // 3. サイズ設定 (固定スケール [16, 16]%)
      try {
        var targetScale = [16, 16];
        if (
          scaleProp.canSetValue &&
          scaleProp.value instanceof Array &&
          scaleProp.value.length >= 2
        ) {
          scaleProp.setValue(targetScale);
          writeLn(
            "  固定スケール設定: [" +
              targetScale[0] +
              "%, " +
              targetScale[1] +
              "%]"
          );
        } else {
          writeLn(
            "警告: スケールプロパティに値を設定できないか、形式が不正です。"
          );
        }
      } catch (scaleError) {
        writeLn("警告: スケール設定中にエラー: " + scaleError.toString());
      }
    } catch (e) {
      writeLn("ロゴレイヤーの設定中に予期せぬエラー: " + e.toString());
    }
    writeLn(
      "ロゴレイヤー '" + logoLayer.name + "' のトランスフォーム設定終了。"
    );
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
  function saveProjectNameConfig(projectName) {
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
    var configData = { project_name: projectName };
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
    } finally {
      configFile.close();
    }
  }
  function loadProjectNameConfig() {
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
      if (configData && typeof configData.project_name === "string") {
        return configData.project_name;
      } else {
        return null;
      }
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

  // デバッグ用 writeLn 関数
  function writeLn(message) {
    var currentTime = new Date().toLocaleTimeString();
    var output = "[" + currentTime + "] " + message;
    // $.writeln(output); // ExtendScript Toolkit のコンソールに出力する場合
    try {
      // スクリプトUIパネルにログ表示エリアがあればそこに追加するなどの処理
      // 例: if (myLogEditText) { myLogEditText.text += output + "\n"; }
      // alert(output); // 簡易的にアラートで表示する場合（デバッグ時のみ有効に）
    } catch (e) {
      // alert("writeLn Error: " + e.toString());
    }
  }

  // --- スクリプト実行開始 ---
  var myPanel = buildUI(thisObj);
})(this);
