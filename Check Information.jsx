(function (thisObj) {
  //=====================================
  // 1) UI(パネル)を構築する関数
  //=====================================
  function buildUI(thisObj) {
    var myPanel =
      thisObj instanceof Panel
        ? thisObj
        : new Window("palette", "", undefined, { resizeable: true });

    myPanel.text = "";
    myPanel.orientation = "column";
    myPanel.alignChildren = "fill";

    // --- (A) ソース一覧パネル ---
    var panelComp = myPanel.add("panel", undefined, "ソース一覧");
    panelComp.orientation = "column";
    panelComp.alignChildren = "fill";

    var ddComp = panelComp.add("dropdownlist", undefined, []);
    ddComp.minimumSize.width = 300;

    // --- (B) テキストレイヤー名パネル ---
    var panelLayers = myPanel.add("panel", undefined, "各テキストレイヤー名");
    panelLayers.orientation = "column";
    panelLayers.alignChildren = "left";

    // currentInfo
    var grpCurrent = panelLayers.add("group");
    grpCurrent.add("statictext", undefined, "currentInfo:");
    var editCurrentInfoName = grpCurrent.add(
      "edittext",
      undefined,
      "currentInfo"
    );
    editCurrentInfoName.characters = 20;

    // compositionInfo
    var grpCompInfo = panelLayers.add("group");
    grpCompInfo.add("statictext", undefined, "compositionInfo:");
    var editCompositionInfoName = grpCompInfo.add(
      "edittext",
      undefined,
      "compositionInfo"
    );
    editCompositionInfoName.characters = 20;

    // compositionName
    var grpCompName = panelLayers.add("group");
    grpCompName.add("statictext", undefined, "compositionName:");
    var editCompositionName = grpCompName.add(
      "edittext",
      undefined,
      "compositionName"
    );
    editCompositionName.characters = 20;

    // --- (C) ボタン ---
    var grpBtns = myPanel.add("group");
    grpBtns.alignment = "center";

    var btnRefresh = grpBtns.add("button", undefined, "更新");
    var btnApply = grpBtns.add("button", undefined, "OK");

    // === 内部管理用配列 ===
    var avSources = []; // {source, type, name} を格納

    // === (1) 「更新」ボタン ===
    btnRefresh.onClick = function () {
      ddComp.removeAll();
      avSources = [];

      // (a) プロジェクト保存済み & アクティブコンポあるか？
      if (!hasSavedProject()) {
        // 何も表示せず終了
        return;
      }
      var ac = getActiveComp();
      if (!ac) {
        // 何も表示せず終了
        return;
      }

      // (b) アクティブコンポ内のAVレイヤーを収集
      for (var i = 1; i <= ac.numLayers; i++) {
        var lyr = ac.layer(i);
        if (!(lyr instanceof AVLayer)) continue;
        var s = lyr.source;
        if (!s) continue;

        // -----------------------------------------
        // (1) s が CompItem → コンポレイヤー
        // (2) s が FootageItem → 動画のみ (静止画は除外)
        // -----------------------------------------
        if (s instanceof CompItem) {
          avSources.push({ source: s, type: "comp", name: s.name });
        } else if (s instanceof FootageItem) {
          // フッテージが動画かどうかを判定
          // 静止画かどうかは mainSource.isStill が true/false でわかる
          var f = s.mainSource; // FootageSource
          if (!f.isStill) {
            // 静止画でなければ「動画」とみなして追加
            avSources.push({ source: s, type: "footage", name: s.name });
          }
        }
      }

      if (avSources.length === 0) {
        ddComp.add("item", "ソースなし");
        ddComp.selection = 0;
      } else {
        for (var k = 0; k < avSources.length; k++) {
          ddComp.add("item", avSources[k].name);
        }
        ddComp.selection = 0;
      }
    };

    // === (2) 「OK」ボタン ===
    btnApply.onClick = function () {
      if (!hasSavedProject()) {
        return;
      }
      var ac = getActiveComp();
      if (!ac) {
        return;
      }
      if (ddComp.items.length === 0 || !ddComp.selection) {
        return;
      }

      var selIndex = ddComp.selection.index;
      var pickedObj = avSources[selIndex];
      if (!pickedObj || !pickedObj.source) {
        return;
      }

      // レイヤー名取得
      var currentInfoName = editCurrentInfoName.text;
      var compositionInfoName = editCompositionInfoName.text;
      var compositionNameName = editCompositionName.text;

      // 書き込み
      applyInfoToTextLayers(
        ac,
        pickedObj,
        currentInfoName,
        compositionInfoName,
        compositionNameName
      );
    };

    // パネルの初期状態は「ソースなし」
    ddComp.removeAll();
    ddComp.add("item", "ソースなし");
    ddComp.selection = 0;

    myPanel.layout.layout(true);
    return myPanel;
  }

  //=====================================
  // 2) 書き込み処理の本体
  //=====================================
  function applyInfoToTextLayers(
    activeComp,
    pickedObj,
    layerNameCurrent,
    layerNameCompInfo,
    layerNameCompName
  ) {
    if (!activeComp) return;

    var currentInfoLayer = activeComp.layer(layerNameCurrent);
    var compositionInfoLayer = activeComp.layer(layerNameCompInfo);
    var compositionNameLayer = activeComp.layer(layerNameCompName);

    if (!currentInfoLayer || !currentInfoLayer.property("Source Text")) return;
    if (!compositionInfoLayer || !compositionInfoLayer.property("Source Text"))
      return;
    if (!compositionNameLayer || !compositionNameLayer.property("Source Text"))
      return;

    // (A) currentInfo: 現在時刻 + コンピュータ名
    var currentTimeString = createDateTimeString();
    currentInfoLayer.property("Source Text").setValue(currentTimeString);

    // (B) compositionName / compositionInfo
    var projectFileName = getProjectName();
    var activeCompName = activeComp.name;
    var compNameString = "";
    var compInfoString = "";

    if (pickedObj.type === "comp") {
      // ****** コンポレイヤー ******
      var rawFps = pickedObj.source.frameRate;
      var displayedFps = parseFloat(rawFps.toFixed(2));

      var compW = pickedObj.source.width;
      var compH = pickedObj.source.height;
      var compDur = pickedObj.source.duration;
      var compTotalFrames = Math.round(displayedFps * compDur);

      compNameString = projectFileName + " | " + pickedObj.source.name;
      compInfoString =
        compW +
        "x" +
        compH +
        " / " +
        displayedFps +
        "fps / " +
        compTotalFrames +
        "f";
    } else if (pickedObj.type === "footage") {
      // ****** 動画フッテージ ******
      compNameString = projectFileName + " | " + activeCompName;

      var footRawFps = pickedObj.source.frameRate;
      var footFps = parseFloat(footRawFps.toFixed(2));
      var footDur = pickedObj.source.duration;
      var footTotalFrames = Math.round(footFps * footDur);
      var footTC = convertSecondsToTC(footDur, footFps);

      compInfoString =
        pickedObj.name + " TC/Fm: " + footTC + " / " + footTotalFrames + "f";
    } else {
      // 不明な場合
      compNameString = projectFileName + " | " + activeCompName;
      compInfoString = "(不明なソース)";
    }

    compositionNameLayer.property("Source Text").setValue(compNameString);
    compositionInfoLayer.property("Source Text").setValue(compInfoString);

    // ここでもアラート等は出さない
  }

  //=====================================
  // 3) ユーティリティ関数
  //=====================================
  function hasSavedProject() {
    return app.project && app.project.file;
  }
  function getActiveComp() {
    if (
      !app.project ||
      !app.project.activeItem ||
      !(app.project.activeItem instanceof CompItem)
    ) {
      return null;
    }
    return app.project.activeItem;
  }
  function getProjectName() {
    if (!hasSavedProject()) return "Untitled";
    var nm = app.project.file.name; // "MyProject.aep"
    return nm.replace(/\.aep$/i, "");
  }
  function createDateTimeString() {
    var d = new Date();
    function pad(num) {
      return ("0" + num).slice(-2);
    }
    var year = d.getFullYear();
    var month = pad(d.getMonth() + 1);
    var day = pad(d.getDate());
    var hour = pad(d.getHours());
    var min = pad(d.getMinutes());
    var sec = pad(d.getSeconds());

    var offsetMinutes = -d.getTimezoneOffset();
    var sign = offsetMinutes >= 0 ? "+" : "-";
    offsetMinutes = Math.abs(offsetMinutes);
    var offsetH = Math.floor(offsetMinutes / 60);
    var offsetM = offsetMinutes % 60;
    function padOffset(v) {
      return v < 10 ? "0" + v : v;
    }
    var offsetStr = sign + padOffset(offsetH) + ":" + padOffset(offsetM);

    var osInfo = $.os;
    var machineName = "";
    if (osInfo.indexOf("Mac") >= 0) {
      machineName = system.callSystem("scutil --get ComputerName");
    } else if (osInfo.indexOf("Windows") >= 0) {
      machineName = system.callSystem("hostname");
    } else {
      machineName = "UnknownOS";
    }
    machineName = machineName.replace(/[\r\n]+$/, "");

    return (
      year +
      "." +
      month +
      "." +
      day +
      " T " +
      hour +
      ":" +
      min +
      ":" +
      sec +
      " " +
      offsetStr +
      " " +
      machineName
    );
  }
  function convertSecondsToTC(durationSec, fps) {
    var totalFrames = Math.round(durationSec * fps);
    var hours = Math.floor(totalFrames / (fps * 3600));
    var remainder = totalFrames % (fps * 3600);

    var minutes = Math.floor(remainder / (fps * 60));
    remainder = remainder % (fps * 60);

    var seconds = Math.floor(remainder / fps);
    var frames = remainder % fps;

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

  //=====================================
  // 4) パネル生成・表示
  //=====================================
  var myPanel = buildUI(thisObj);

  if (myPanel instanceof Window) {
    myPanel.center();
    myPanel.show();
  } else {
    myPanel.layout.layout(true);
  }
})(this);
