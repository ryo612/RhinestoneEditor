var polygonMode = 0; //多角形モード
var polygonN;

var side = 1; //片面=1 両面=2
var lensPattern = 1; //凸 or 凹
var dataPattern = 1;
var lensR; //レンズ半径
var focalLength; //焦点距離

var printThick; //印刷1回の厚さ
var maxThick; //印刷可能な厚さ
var refIndex; //屈折率

var r; //曲率半径
var lensH; //レンズ厚さ

var lensLayerN; //層の数（印刷回数）
var lensLayerR; //各層の直径

var layerNum = activeDocument.layers.length; //実行時のレイヤー数

var lensTipeNum = 0; //レンズの種類数
var topLayerNum = 0; //上部レイヤー数
var btmLayerNum = 0; //下部レイヤー数

var glossSpot;

var whiteColor = new CMYKColor();
whiteColor.cyan = 0;
whiteColor.magenta = 0;
whiteColor.yellow = 0;
whiteColor.black = 0;

var blackColor = new CMYKColor();
blackColor.cyan = 0;
blackColor.magenta = 0;
blackColor.yellow = 0;
blackColor.black = 100;

var milli = 72 / 25.4; //単位修正

var originalFileName;
var tmpFileName;
var folderName;

var cmd = ["Live Pathfinder Add",
    "Live Pathfinder Intersect",
    "Live Pathfinder Exclude",
    "Live Pathfinder Subtract",
    "Live Pathfinder Minus Back",
    "Live Pathfinder Divide",
    "Live Pathfinder Trim",
    "Live Pathfinder Merge",
    "Live Pathfinder Crop",
    "Live Pathfinder Outline",
    "Live Pathfinder Hard Mix",
    "Live Pathfinder Soft Mix",
    "Live Pathfinder Trap"
];

//塗り、線なしの指定
function setGlossSpot() {
    try {
        glossSpot = activeDocument.swatches.getByName('RDG_GLOSS');
        activeDocument.defaultFillColor = glossSpot.color;
    } catch (e) {
        var glossColor = new CMYKColor;
        glossColor.cyan = 50;
        glossColor.magenta = 25;
        glossColor.yellow = 25;
        glossColor.black = 0;

        var newSpot = activeDocument.spots.add();
        newSpot.name = 'RDG_GLOSS';
        newSpot.colorType = ColorModel.SPOT;
        newSpot.color = glossColor;

        glossSpot = activeDocument.swatches.getByName('RDG_GLOSS');
        activeDocument.defaultFillColor = glossSpot.color;
    }
    activeDocument.defaultStroked = false;
}

//activeDocument.changeMode(ChangeMode.CMYK);

//ファイルを上書き保存，未保存なら名前を付けて保存
function saveFile() {
    try {
        activeDocument.save();
    } catch (e) {
        filename = File.saveDialog("ファイルを保存"); //保存ダイアログ
        if (filename) {
            fileObj = new File(filename);
            var aiOpt = new IllustratorSaveOptions();
            var saveObj = activeDocument.saveAs(fileObj, aiOpt);
        }else{
            return false;    
        }
    }
    originalFileName = activeDocument.fullName;
    var originalFile = new File(originalFileName);
    
    return true;
}

function setPrintSetting() {
    app.executeMenuCommand('deselectall'); //全選択解除
    layerObj = activeDocument.layers;
    btmLayer = layerObj[layerObj.length - 1].name;
    if (btmLayer.slice(0, 4) == "印刷設定") {
        var c1 = btmLayer.indexOf('率');
        var c2 = btmLayer.indexOf(',');
        var c3 = btmLayer.indexOf('さ');
        var c4 = btmLayer.indexOf('m');
        refIndex = Number(btmLayer.slice(c1 + 1, c2));
        printThick = Number(btmLayer.slice(c3 + 1, c4));
    }
    for (var i = 0; i < layerNum; i++) {
        layerName = layerObj[i].name;
        var initial = layerName.slice(0, 1)
        if (initial == "#") {
            var comma = layerName.indexOf(',');
            lensTipeNum = Number(layerName.slice(1, comma)); //レンズの種類数
            topLayerNum = i;
            break;
        }
    }
    btmLayerNum = layerNum - lensTipeNum - topLayerNum - 1;
}


function processingByLensType(layer) {
    layerObj = activeDocument.layers;
    layerName = layerObj[layer].name;
    //各パラメータの取得
    var c1 = layerName.indexOf(',');
    var c2 = layerName.indexOf(',', c1 + 1);
    var c3 = layerName.indexOf('径');
    var c4 = layerName.indexOf('m');
    var c5 = layerName.indexOf('離');
    var c6 = layerName.indexOf('m', c5);
    var c7 = layerName.indexOf('さ');
    var c8 = layerName.indexOf('m', c7);
    var c9 = layerName.indexOf('径', c8);
    var c10 = layerName.indexOf('m', c9);
    // if (layerName.slice(-1) != "m") {
    //     if(layerName.slice(-1) == "0"){
    //         polygonMode = 2;
    //         polygonN = layerName.slice(c10 + 3);
    //     }else{
    //         polygonMode = 1;
    //         polygonN = layerName.slice(c10 + 3);
    //     }
    // }
    if (layerName.slice(-1) != "m") {
        polygonMode = 1;
        polygonN = layerName.slice(c10 + 3);
    }else{
        polygonMode = 0;
    }

    lensPattern = layerName.slice(c1 + 2, c1 + 3);
    lensR = Number(layerName.slice(c3 + 1, c4) / 2);
    focalLength = Number(layerName.slice(c5 + 1, c6));
    lensH = Number(layerName.slice(c7 + 1, c8));
    r = Number(layerName.slice(c9 + 1, c10));

    lensLayerN = Math.floor(lensH / printThick); //印刷回数の計算

    //各オブジェクトの位置を取得
    var lens = layerObj[layer].pathItems;
    var lensPosition = [];
    if (polygonMode == 0 /*|| polygonMode == 2*/) {
        for (var i = 0; i < lens.length; i++) {
            var carLensPosition = [
                Math.round(lens[i].geometricBounds[0] * 1000) / 1000,
                Math.round(lens[i].geometricBounds[1] * 1000) / 1000,
                Math.round(lens[i].geometricBounds[2] * 1000) / 1000,
                Math.round(lens[i].geometricBounds[3] * 1000) / 1000
            ];

            //中心座標
            lensPosition[i] = [];
            lensPosition[i][0] = Math.abs(carLensPosition[0] + carLensPosition[2]) / 2;
            lensPosition[i][1] = Math.abs(carLensPosition[1] + carLensPosition[3]) / 2;
        }
    } else {
        for (var i = 0; i < lens.length; i++) {
            var carLensPosition = [
                Math.round(lens[i].geometricBounds[0] * 1000) / 1000,
                Math.round(lens[i].geometricBounds[1] * 1000) / 1000,
                Math.round(lens[i].geometricBounds[2] * 1000) / 1000,
                Math.round(lens[i].geometricBounds[3] * 1000) / 1000
            ];

            //中心座標
            if (polygonN % 2 == 0) {
                lensPosition[i] = [];
                lensPosition[i][0] = Math.abs(carLensPosition[0] + carLensPosition[2]) / 2; //x座標
                lensPosition[i][1] = Math.abs(carLensPosition[1] + carLensPosition[3]) / 2; //y座標
            } else {
                lensPosition[i] = [];
                lensPosition[i][0] = Math.abs(carLensPosition[0] + carLensPosition[2]) / 2;
                //alert(lensPosition[i][0] / milli)
                lensPosition[i][1] = Math.abs(carLensPosition[1]) + (lensR * milli);
                //alert(lensPosition[i][1] / milli)
            }
        }
    }

    //レイヤーの生成
    curLayerNum = activeDocument.layers.length;
    if (curLayerNum < lensLayerN + layerNum) {
        for (var j = curLayerNum - layerNum; j < lensLayerN; j++) {
            newLayer = activeDocument.layers.add();
            newLayer.name = "MatteLayer_" + (j + 1);
        }
    }

    var lensLayerR = new Array(lensLayerN);
    var center = lensR * 72 / 25.4;
    /*if (polygonMode == 1) {

        //レイヤー毎の処理
        for (var j = 0; j < lensLayerN; j++) {

            //層の形状算出
            //lensLayerR[j] = Math.sqrt(Math.pow(r, 2) - Math.pow(r - lensH + printThick * (j + 1), 2)) * 144 / 25.4;
            lensLayerR[j] = (Math.sqrt(Math.pow(r, 2) - Math.pow(r - lensH + printThick * (0 + 1), 2)) * 144 / 25.4)*(lensLayerN-j)/lensLayerN;
            lensLayerR[j] = lensLayerR[j].toFixed(3); //小数点第3位まで

            //レイヤー内の各レンズの処理
            for (var i = 0; i < lens.length; i++) {

                //描画位置
                var positionX = lensPosition[i][0] - (lensLayerR[j] / 2);
                var positionY = lensPosition[i][1] - (lensLayerR[j] / 2);

                //層の描画
                layerObj = activeDocument.layers[activeDocument.layers.length - (j + 1) - layerNum];
                //pObj = layerObj.pathItems.ellipse(-positionY, positionX, lensLayerR[j], lensLayerR[j]);
                pObj = layerObj.pathItems.polygon(lensPosition[i][0], -lensPosition[i][1], (/*lensR*///lensLayerR[j]) /2/* milli*/, polygonN);
                // pObj = lens[i].duplicate(layerObj);
                // pObj.resize((lensLayerR[j]/lensLayerR[0])*100,(lensLayerR[j]/lensLayerR[0])*100);
                

                /*pObj.filled = true; //　塗りあり
                pObj.stroked = false; //　線なし
            }
        }

        
    } else */if(polygonMode == 1){//ピラミッド
        //レイヤー毎の処理
        for (var j = 0; j < lensLayerN; j++) {

            //層の形状算出
            //lensLayerR[j] = Math.sqrt(Math.pow(r, 2) - Math.pow(r - lensH + printThick * (j + 1), 2)) * 144 / 25.4;
            lensLayerR[j] = (Math.sqrt(Math.pow(r, 2) - Math.pow(r - lensH + printThick * (0 + 1), 2)) * 144 / 25.4)*(lensLayerN-j)/lensLayerN;
            lensLayerR[j] = lensLayerR[j].toFixed(3); //小数点第3位まで

            if(lensLayerR[j]/lensLayerR[0]<0.4){
                break;
            }

            //レイヤー内の各レンズの処理
            for (var i = 0; i < lens.length; i++) {

                //描画位置
                var positionX = lensPosition[i][0] - (lensLayerR[j] / 2);
                var positionY = lensPosition[i][1] - (lensLayerR[j] / 2);

                //層の描画
                layerObj = activeDocument.layers[activeDocument.layers.length - (j + 1) - layerNum];
                //pObj = layerObj.pathItems.ellipse(-positionY, positionX, lensLayerR[j], lensLayerR[j]);
                // pObj = layerObj.pathItems.polygon(lensPosition[i][0], -lensPosition[i][1], (/*lensR*/lensLayerR[j]) /2/* milli*/, polygonN);
                pObj = lens[i].duplicate(layerObj);
                pObj.resize((lensLayerR[j]/lensLayerR[0])*100,(lensLayerR[j]/lensLayerR[0])*100);
                

                // pObj.filled = true; //　塗りあり
                pObj.fillColor = glossSpot.color;
                pObj.stroked = false; //　線なし

            }
        }

    }else if (lensPattern == '凸') {

        //レイヤー毎の処理
        for (var j = 0; j < lensLayerN; j++) {

            //層の形状算出
            lensLayerR[j] = Math.sqrt(Math.pow(r, 2) - Math.pow(r - lensH + printThick * (j + 1), 2)) * 144 / 25.4;
            lensLayerR[j] = lensLayerR[j].toFixed(3); //小数点第3位まで

            //レイヤー内の各レンズの処理
            for (var i = 0; i < lens.length; i++) {

                //描画位置
                var positionX = lensPosition[i][0] - (lensLayerR[j] / 2);
                var positionY = lensPosition[i][1] - (lensLayerR[j] / 2);

                //層の描画
                layerObj = activeDocument.layers[activeDocument.layers.length - (j + 1) - layerNum];
                pObj = layerObj.pathItems.ellipse(-positionY, positionX, lensLayerR[j], lensLayerR[j]);

                pObj.filled = true; //　塗りあり
                pObj.stroked = false; //　線なし
            }
        }
    } else if (lensPattern == '凹') {
        var ouLens = lensR * 2 * 72 / 25.4

        //レイヤー毎の処理
        for (var j = 0; j < lensLayerN; j++) {

            //層の形状算出
            lensLayerR[j] = Math.sqrt(Math.pow(r, 2) - Math.pow(r - lensH + printThick * j, 2)) * 144 / 25.4;
            lensLayerR[j] = lensLayerR[j].toFixed(3); //小数点第3位まで


            //レイヤー内の各レンズの処理
            for (var i = 0; i < lens.length; i++) {

                //描画位置
                var positionX = lensPosition[i][0] - (lensLayerR[j] / 2);
                var positionY = lensPosition[i][1] - (lensLayerR[j] / 2);
                var positionX2 = lensPosition[i][0] - (ouLens / 2);
                var positionY2 = lensPosition[i][1] - (ouLens / 2);

                //層の描画
                layerObj = activeDocument.layers[activeDocument.layers.length - layerNum - lensLayerN + j];

                pObj = layerObj.pathItems.ellipse(-positionY, positionX, lensLayerR[j], lensLayerR[j]);

                pObj.filled = true; //　塗りあり
                pObj.stroked = false; //　線なし

                pObj2 = layerObj.pathItems.ellipse(-positionY2, positionX2, ouLens, ouLens);

                var gp = layerObj.groupItems.add();
                pObj.move(gp, ElementPlacement.PLACEATEND);
                pObj2.move(gp, ElementPlacement.PLACEATEND);
                gp.selected = true;
                app.executeMenuCommand(cmd[2]);
                app.executeMenuCommand('expandStyle');

            }
        }
    }
}


function main() {
    setGlossSpot();
    for (var i = 0; i < lensTipeNum; i++) {
        var pLayer = activeDocument.layers.length - (layerNum - lensTipeNum - topLayerNum + i + 1);
        processingByLensType(pLayer);
    }
}

//レイヤーの並べ替え
function layerSorting() {
    var layerN = activeDocument.layers.length;
    var layerObj = activeDocument.layers;
    for (var i = 0; i < layerN; i++) { //全レイヤーロック解除
        layerObj[i].locked = false;
    }

    activeDocument.layers[layerN - 1].remove(); //最下レイヤー削除

    // 光沢層の生成1、#レイヤーのグループ化
    var gObj;
    var glossLayer = activeDocument.layers.add();
    glossLayer.name = "GlossLayer_1";
    var gp = glossLayer.groupItems.add();
    for (var i = 0; i < lensTipeNum; i++) {
        var pathNum = activeDocument.layers[layerN - btmLayerNum - 1 - i].pathItems.length;
        for (var j = 0; j < pathNum; j++) {
            var gObj = activeDocument.layers[layerN - btmLayerNum - 1 - i].pathItems[0];
            gObj.fillColor = glossSpot.color;
            gObj.stroked = false;
            gObj.move(gp, ElementPlacement.PLACEATEND);
        }
    }

    //空になったレイヤー削除
    for (var i = 0; i < lensTipeNum; i++) {
        layerObj = activeDocument.layers;
        var layerN = activeDocument.layers.length;
        layerObj[layerN - btmLayerNum - 1].remove();
    }

    //上のCMYKレイヤー移動
    for (var i = 0; i < topLayerNum; i++) {
        layerObj = activeDocument.layers;
        var layerN = activeDocument.layers.length;
        layerObj[layerN - btmLayerNum - 1].zOrder(ZOrderMethod.BRINGTOFRONT);
    }

    // 光沢層の生成2
    var glossLayer2 = activeDocument.layers.add();
    glossLayer2.name = "GlossLayer_2";
    var copyItem = activeDocument.layers[topLayerNum + 1].groupItems[0].duplicate();
    copyItem.moveToEnd(glossLayer2);

    //app.executeMenuCommand('group');
    //selObj = activeDocument.selection;
    //selObj.move(gp, ElementPlacement.PLACEATEND);
}

function copyLine() {
    var layerN = activeDocument.layers.length;
    var gp = activeDocument.layers[layerN - 1].groupItems.add();
    var pathNum = activeDocument.layers[layerN - 1].pathItems.length;
    //最下レイヤーのパスをグループ化
    for (var i = 0; i < pathNum; i++) {
        var gObj = activeDocument.layers[layerN - 1].pathItems[0];
        gObj.filled = false;
        gObj.strokeColor = whiteColor;
        gObj.move(gp, ElementPlacement.PLACEATEND);
    }
    //最下レイヤーのパスを複製して各レイヤーに
    for (var i = 0; i < layerN - 1; i++) {
        var copyItem = activeDocument.layers[layerN - 1].groupItems[0].duplicate();
        copyItem.moveToEnd(activeDocument.layers[i]);
    }
    //最下レイヤーのパスの色を戻す
    for (var i = 0; i < pathNum; i++) {
        var sObj = activeDocument.layers[layerN - 1].groupItems[0].pathItems[i];
        sObj.strokeColor = blackColor;
    }
}


function expotFile() {
    //新規フォルダ、一時ファイルの作成
    var filePath = activeDocument.path;
    var onlyFileName = activeDocument.name.split(".");
    var folderName = filePath + "/" + onlyFileName[0];
    var folderObj = new Folder(folderName);
    folderObj.create();

    var tmpFileName = new File(folderName + "/" + "tmp" + ".ai");
    var aiOpt = new IllustratorSaveOptions();
    var tmpFile = activeDocument.saveAs(tmpFileName, aiOpt);

    //レイヤー毎に書き出し
    var expotLayerNum = activeDocument.layers.length;
    for (var i = expotLayerNum; i > 0; i--) {
        for (var j = expotLayerNum; j > 0; j--) {
            if (i != j) activeDocument.layers[j - 1].remove();
        }

        var saveEpsFile = new File(folderName + "/" + (expotLayerNum - i + 1) + "_" + activeDocument.layers[0].name + ".eps");
        var epsOpt = new EPSSaveOptions();
        var epsFile = activeDocument.saveAs(saveEpsFile, epsOpt);

        activeDocument.close(SaveOptions.DONOTSAVECHANGES);
        if (i == 1) {
            open(originalFileName);
            tmpFileName.remove();
        } else {
            open(tmpFileName);
        }
    }
}



function symbol() {
    var docRef = null;
    var symbolRef = null;
    var iCount = 0;

try {
docRef = app.activeDocument;
iCount = docRef.symbolItems.length;
for (var num = iCount-1; num>=0; num--){
     symbolRef = docRef.symbolItems[num];
     symbolRef.breakLink();
}

}
catch(err)
{
rs = 'ERROR: ' + (err.number & 0xFFFF) + ', ' + err.description;
alert(rs);
}

app.executeMenuCommand("selectall");
app.executeMenuCommand("ungroup");
app.executeMenuCommand("selectall");
app.executeMenuCommand("ungroup");
activeDocument.selection = null;

}



activeDocument.selection = null;
if(saveFile()==true){
    symbol();
setPrintSetting();
main();
layerSorting();
copyLine();
expotFile();
}