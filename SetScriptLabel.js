//----------------------------------------------------------

function getTextFramesInPage(targetPage) {
    if(!targetPage.isValid) {
      return {};
    }
  
    var textFrames = targetPage.textFrames;
    var formatTextFrames = {};
    var maxTextFrames = textFrames.length;
    for (var j = 0; j < maxTextFrames; j++) {
      textFrames[j].contents = '';
      formatTextFrames[textFrames[j].label] = textFrames[j];
    }
  
    return formatTextFrames;
  }

//----------------------------------------------------------

// 初期処理
(function () {
      // 今開いているinDesignドキュメントを取得
  var doc = app.activeDocument;

  var maxPages = doc.pages.length;

// 赤色のスウォッチを作成または取得
var redColor;
try {
    redColor = doc.colors.item("Red");
    var colorName = redColor.name;
} catch (myError) {
    redColor = doc.colors.add({name:"Red", model:ColorModel.process, colorValue:[0, 100, 100, 0]});
}

  var x = 25.105;
  var y = 68.209;

  for (var i = 0; i < maxPages; i++) {
    
    var targetPage = doc.pages.item(i);

    if (!targetPage) {
      continue;
    }

    for (var j = 0; j < targetPage.textFrames.length; j++) {
        var textFrame = targetPage.textFrames[j];
        
        // テキストフレームの座標を取得
        var bounds = textFrame.geometricBounds;
        var x1 = bounds[1];
        var y1 = bounds[0];
        var x2 = bounds[3];
        var y2 = bounds[2];
        // alert("テスト" + bounds.join(','))

        // 座標がテキストフレーム内にあるか確認
        if (x >= x1 && x <= x2 && y >= y1 && y <= y2) {
            // 座標が一致するテキストフレームを見つけた場合
            
            textFrame.fillColor = redColor;
            break;
        }
    }

    // マスターデータに応じてアクションを切り替える
    // switch (pageData.masterPageName) {
    //   case MASTER_PAGE_NAME_MAP:
    //     // setMapPage(pageData.data, formatTextFrames);
    //     break;
    // }

  }

})();