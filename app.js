var IS_EDIT = false;

var EDIT_TARGET_START_PAGE = 0;// 編集対象の開始ページ 0は最初から
var EDIT_TARGET_END_PAGE = 0;// 編集対象の終了ページ 0は最後まで

var MASTER_PAGE_LABEL_NORMAL = "";
var MASTER_PAGE_LABEL_MAP = "map";
var MASTER_PAGE_LABEL_PP = "pp";
var MASTER_PAGE_LABEL_SS = "ss";

var MASTER_PAGE_KEY = "masterPageKey";
var MASTER_PAGE_NAME_NORMAL = "M-マスター";
var MASTER_PAGE_NAME_MAP = "MAP-コマ図";
var MASTER_PAGE_NAME_PP = "PP-ポイント写真";
var MASTER_PAGE_NAME_SS = "SS-TC/SS";

var LABEL_PAGE_MASTER = 'page-master';
var LABEL_PAGE_LEG = 'leg';
var LABEL_PAGE_SECTION = 'section';
var LABEL_PAGE_FROM_POINT_NO = 'from_point_no';
var LABEL_PAGE_FROM_POINT_NAME = 'from_point_name';
var LABEL_PAGE_SS_NO = 'ss_no';
var LABEL_PAGE_SS_NAME = 'ss_name';
var LABEL_PAGE_SS_DIST = 'ss_dist';
var LABEL_PAGE_SS_AVERAGE = 'ss_average';
var LABEL_PAGE_SS_RECORD = 'ss_record';
var LABEL_PAGE_TO_POINT_NO = 'to_point_no';
var LABEL_PAGE_TO_POINT_NAME = 'to_point_name';
var LABEL_PAGE_LIAISON_DIST = 'liaison_dist';
var LABEL_PAGE_LIAISON_AVERAGE = 'liaison_average';
var LABEL_PAGE_TARGET_TIME = 'target_time';
var LABEL_PAGE_TOTAL = 'total';
var LABEL_PAGE_PARTIAL = 'partial';
var LABEL_PAGE_DIST_TO_GO = 'dist_to_go';
var LABEL_PAGE_DIRECTION_NO = 'direction_no';
var LABEL_PAGE_STAGE_TYPE = 'stage_type';
var LABEL_PAGE_LABEL = 'page_label';

var DATA_TEXT_ROW_LABELS = [
  LABEL_PAGE_MASTER,
  LABEL_PAGE_LEG,
  LABEL_PAGE_SECTION,
  LABEL_PAGE_FROM_POINT_NO,
  LABEL_PAGE_FROM_POINT_NAME,
  LABEL_PAGE_SS_NO,
  LABEL_PAGE_SS_NAME,
  LABEL_PAGE_SS_DIST,
  LABEL_PAGE_SS_AVERAGE,
  LABEL_PAGE_SS_RECORD,
  LABEL_PAGE_TO_POINT_NO,
  LABEL_PAGE_TO_POINT_NAME,
  LABEL_PAGE_LIAISON_DIST,
  LABEL_PAGE_LIAISON_AVERAGE,
  LABEL_PAGE_TARGET_TIME,
  LABEL_PAGE_TOTAL,
  LABEL_PAGE_PARTIAL,
  LABEL_PAGE_DIST_TO_GO,
  LABEL_PAGE_DIRECTION_NO,
  LABEL_PAGE_STAGE_TYPE,
  LABEL_PAGE_LABEL
];

var PAGE_HEADER_LABELS = [
  LABEL_PAGE_LEG,
  LABEL_PAGE_SECTION,
  LABEL_PAGE_FROM_POINT_NO,
  LABEL_PAGE_FROM_POINT_NAME,
  LABEL_PAGE_SS_NO,
  LABEL_PAGE_SS_NAME,
  LABEL_PAGE_SS_DIST,
  LABEL_PAGE_SS_AVERAGE,
  LABEL_PAGE_SS_RECORD,
  LABEL_PAGE_TO_POINT_NO,
  LABEL_PAGE_TO_POINT_NAME,
  LABEL_PAGE_LIAISON_DIST,
  LABEL_PAGE_LIAISON_AVERAGE,
  LABEL_PAGE_TARGET_TIME,
  LABEL_PAGE_STAGE_TYPE,
  LABEL_PAGE_LABEL
];

var PAGE_HEADER_LABELS_COUNT = PAGE_HEADER_LABELS.length;

var EXTEND_SRINGS = {};
EXTEND_SRINGS[LABEL_PAGE_SS_DIST] = 'km';
EXTEND_SRINGS[LABEL_PAGE_SS_AVERAGE] = 'km/h';
EXTEND_SRINGS[LABEL_PAGE_LIAISON_DIST] = 'km';
EXTEND_SRINGS[LABEL_PAGE_LIAISON_AVERAGE] = 'km/h';

//----------------------------------------------------------

Array.prototype.indexOf = function (item) {
  var index = 0, length = this.length;
  for (; index < length; index++) {
    if (this[index] === item)
      return index;
  }
  return -1;
};

//----------------------------------------------------------

// 指定データがlineの何列目にあるかを取得する
function getLineData(lineData, keyName) {
  var index = DATA_TEXT_ROW_LABELS.indexOf(keyName);
  var value = lineData[index];
  return value ? value : '';
}

//----------------------------------------------------------

// roadbook.txtのデータを整形する
function getFileData() {
  const dataFile = File.openDialog("Select a .csv file to open:");
  if (!dataFile.exists) {
    alert(filePath + "が見つかりません。");
    return [];
  }

  // ファイルを読み込む
  dataFile.open("r");
  var lines = [];
  while (!dataFile.eof) {
    lines.push(dataFile.readln());
  }
  dataFile.close();

  return lines;
}

//----------------------------------------------------------

function formatLineDataForPages(lines) {
  if (lines.length < 1) {
    return [];
  }

  var pages = [];
  var currentPage = 0;
  var leg = 0;
  var linesMax = lines.length;
  var headerData = {};
  for (var i = 1; i < linesMax; i++) {
    var lineArray = lines[i].split('\t');
    var pageMasterNames = getLineData(lineArray, LABEL_PAGE_MASTER);
    var directionNo = getLineData(lineArray, LABEL_PAGE_DIRECTION_NO);
    var directionNoOfPage = directionNo % 6;
    var pageMasterNameArray = pageMasterNames.split(',');

    // ページのヘッダー情報を各ページに保存させるために先頭のヘッダー情報を保持しておく
    if (!directionNo || ["", "1"].indexOf(directionNo) >= 0) {
      headerData = {};
      for (var j = 0; j < PAGE_HEADER_LABELS_COUNT; j++) {
        var targetDataLabel = PAGE_HEADER_LABELS[j];
        headerData[targetDataLabel] = getLineData(lineArray, targetDataLabel);
      }
    }

    // TODO headerDataが正しくコピーされているかどうかが怪しい
    var inserHeaderData = headerData;

    var maxPageMasterNameArray = pageMasterNameArray.length;
    for (var j = 0; j < maxPageMasterNameArray; j++) {
      switch (pageMasterNameArray[j]) {
        case MASTER_PAGE_LABEL_MAP:
          if (directionNoOfPage === 1) {
            pages.push({masterPageName: MASTER_PAGE_NAME_MAP, headerData: inserHeaderData, data: []});
            currentPage = pages.length - 1;
          }
          pages[currentPage]['data'].push(lineArray);
          break;
        case MASTER_PAGE_LABEL_PP:
          pages.push({masterPageName: MASTER_PAGE_NAME_PP, headerData: inserHeaderData, data: lineArray});
          break;
        case MASTER_PAGE_LABEL_SS:
          pages.push({masterPageName: MASTER_PAGE_NAME_SS, headerData: inserHeaderData, data: lineArray});
          break;
        default:
          pages.push({masterPageName: MASTER_PAGE_NAME_NORMAL, headerData: inserHeaderData, data: lineArray});
          break;
      }
    }
  }

  return pages;
}

//----------------------------------------------------------

function createPage(doc, masterSpreadsName) {
  if (!masterSpreadsName) {
    return null;
  }

  var parentPage = doc.masterSpreads.itemByName(masterSpreadsName);
  if (!parentPage.isValid) {
    alert("親ページ " + masterSpreadsName + " が見つかりません。");
    return null;
  }

  var newPage = doc.pages.add();
  newPage.appliedMaster = parentPage;

  // 全てのマスターページアイテムをオーバーライド
  var masterItems = newPage.masterPageItems;
  for (var ii = masterItems.length; ii >= 0; ii--) {
    try {
      masterItems[ii].override(newPage);
    } catch (e) {
    }
  }

  return newPage;
}

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

function setPageMapHeader(headerData, textFrames) {
  for (var j = 0; j < PAGE_HEADER_LABELS_COUNT; j++) {
    var targetDataLabel = PAGE_HEADER_LABELS[j];
    if (typeof textFrames[targetDataLabel] !== 'undefined') {
      textFrames[targetDataLabel].contents = '';
      if(typeof headerData[targetDataLabel] === 'undefined' || headerData[targetDataLabel] === '') {
        continue;
      }
      textFrames[targetDataLabel].contents = 
        headerData[targetDataLabel] + 
        ((typeof EXTEND_SRINGS[targetDataLabel] !== 'undefined')? EXTEND_SRINGS[targetDataLabel] : '');
    } 
  }
}

//----------------------------------------------------------

function setMapPage(pageData, textFrames) {
  var boxData = pageData;
  var maxPageData = boxData.length;
  for (var j = 0; j < maxPageData; j++) {
    var lineData = boxData[j];
    var cellIndex = j + 1;
    if (typeof textFrames['total-' + cellIndex] !== 'undefined') textFrames['total-' + cellIndex].contents = getLineData(lineData, LABEL_PAGE_TOTAL);
    if (typeof textFrames['partial-' + cellIndex] !== 'undefined') textFrames['partial-' + cellIndex].contents = getLineData(lineData, LABEL_PAGE_PARTIAL);
    if (typeof textFrames['dist_to_go-' + cellIndex] !== 'undefined') textFrames['dist_to_go-' + cellIndex].contents = getLineData(lineData, LABEL_PAGE_DIST_TO_GO);
    if (typeof textFrames['direction_no-' + cellIndex] !== 'undefined') textFrames['direction_no-' + cellIndex].contents = getLineData(lineData, LABEL_PAGE_DIRECTION_NO);
  }
}

//----------------------------------------------------------

// 初期処理
(function () {
  // ファイルデータの読み込み
  var dataLines = getFileData();
  var pages = formatLineDataForPages(dataLines);

  // 今開いているinDesignドキュメントを取得
  var doc = app.activeDocument;


  // データを１行ずつ処理する
  var targetPage = null;

  var maxPages = pages.length;
  for (var i = 0; i < maxPages; i++) {

    if (EDIT_TARGET_START_PAGE > 0 && i < EDIT_TARGET_START_PAGE - 1) {
      continue;
    }

    if (EDIT_TARGET_END_PAGE > 0 && i > EDIT_TARGET_END_PAGE - 1) {
      continue;
    }

    var pageData = pages[i];

    if (IS_EDIT) {
      targetPage = doc.pages.item(i);
    } else {
      targetPage = createPage(doc, pageData.masterPageName);
    }
    if (!targetPage) {
      continue;
    }

    var formatTextFrames = getTextFramesInPage(targetPage);

    setPageMapHeader(pageData.headerData, formatTextFrames);

    switch (pageData.masterPageName) {
      case MASTER_PAGE_NAME_MAP:
        setMapPage(pageData.data, formatTextFrames);
        break;
    }

  }

  alert(dataLines.length + "行、の処理が完了しました。");
})();
