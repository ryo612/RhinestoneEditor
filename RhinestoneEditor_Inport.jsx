var docObj = activeDocument;

var whiteSpot;

//csv読み込みーーーーーーーーーーーーーーーーー
var num = [5];
var Num = [5];//x,y,r,pattern,pattern2
var n = 0;
inputFolder = Folder.selectDialog("フォルダを指定してください。");
fileObj = new File(inputFolder.getFiles("stone.csv"));
imgObj = inputFolder.getFiles("*.png");

flg = fileObj.open("r");
//ファイルを選択できた場合

if (flg) {
	//テキストファイルを１行ずつ読み込み最終行まで繰り返し

	while (!fileObj.eof) {

		//1行取得
		num[n] = fileObj.readln();
		Num[n] = num[n].split(",");
		n++;
		$.writeln(parseFloat(num[0]) + 0.2);
	}

} else {
	alert("ファイルが開けませんでした。");
}
fileObj.close();

//alert(Num[0][0]);

//csv読み込みーーーーーーーーーーーーーーーーー

var x = Num[0];
var y = Num[1];
var r = Num[2];
var p = Num[3];
var pp = Num[4];
var r_stock = [r[0]];
var r_boolean = true;


for (i = 0; i < r.length; i++) {
	for (j = 0; j < r_stock.length; j++) {
		if (r_stock[j] == r[i]) {
			r_boolean = false;
		}
	}
	if (r_boolean) {
		r_stock.push(r[i]);
	}
	r_boolean = true;
}


var setlay = docObj.layers.add();
setlay.zOrder(ZOrderMethod.SENDTOBACK);
setlay.name = "印刷設定 屈折率" + 1.50 + ", 層厚さ" + 0.035 + "mm";
setlay.locked = true;

var positionlay = docObj.layers.add();
positionlay.name = "枠";
var maxx = Math.max.apply(null, x);
var minx = Math.min.apply(null, x);
var maxy = Math.max.apply(null, y);
var miny = Math.min.apply(null, y);
// for (i = 0; i < x.length; i++) {
// 	if (maxx <= x[i]) {
// 		maxx = x[i];
// 	}
// 	if (minx >= x[i]) {
// 		minx = x[i];
// 	}
// 	if (maxy < y[i]) {
// 		maxy = y[i];
// 	}
// 	if (miny > y[i]) {
// 		miny = y[i];
// 	}
// }


pObj = docObj.pathItems.rectangle(-(miny / 2.0 - px(10)), minx / 2.0 - px(10), (maxx - minx) / 2 + px(20), (maxy - miny) / 2 + px(20));
pObj.filled = false;
//pObj.strokeColor = setRGBColor(255, 255, 255);

var whitelay = docObj.layers.add();
whitelay.name = "ホワイト";
var gp = app.activeDocument.groupItems.add();
for (i = 0; i < r.length; i++) {
	pObj = docObj.pathItems.ellipse(-y[i] / 2+px(r[i])/2.0, x[i] / 2-px(r[i])/2.0, px(r[i]), px(r[i]));//円を描画
	pObj.stroked = false;
	pObj.move(gp, ElementPlacement.PLACEATEND);//要素を挿入する場所→最後に配置
}
pObj = docObj.pathItems.rectangle(-(miny / 2.0 - px(10)), minx / 2.0 - px(10), (maxx - minx) / 2 + px(20), (maxy - miny) / 2 + px(20));
pObj.stroked = false;

pObj.move(gp, ElementPlacement.PLACEATEND);//要素を挿入する場所→最後に配置
gp.selected = true;
setWhiteSpot();

app.executeMenuCommand("Live Pathfinder Exclude");
app.executeMenuCommand('expandStyle');




var patternlay = docObj.layers.add();
patternlay.name = "底面パターン";

for (i = 0; i < r.length; i++) {
	for (j = 0; j < imgObj.length; j++) {
		var ll=parseInt(p[i])+parseInt(1);
		var s = "P" + parseInt(pp[i])+parseInt(ll) + ".png";
		//alert(s);
		if (imgObj[j].name == s) {
			Pfile = new File(imgObj[j]);
			Phaiti = docObj.placedItems.add();
			Phaiti.file = Pfile;
			Phaiti.position = [x[i] / 2.0-px(r[i])/2.0, -y[i] / 2.0+px(r[i])/2.0];
			Phaiti.width = px(r[i]);
			Phaiti.height = px(r[i]);
		}
	}
	//alert(s);
}


for (i = 0; i < r_stock.length; i++) {//レンズレイヤー作成
	var lay = docObj.layers.add();
	if (r_stock[i] <= 4) {
		var R = (Math.pow(r_stock[i] / 2, 2) + Math.pow(r_stock[i] / 2, 2)) / (r_stock[i] / 2 * 2);
		var f = R / (1.5 - 1);
		lay.name = "#" + (i + 1) + ", 凸, 直径" + r_stock[i] + "mm, 焦点距離" + f + "mm, 厚さ" + r_stock[i] / 2 + "mm, 曲率半径" + R + "mm";
	} else {
		var R = (Math.pow(r_stock[i] / 2, 2) + Math.pow(2, 2)) / (2 * 2);
		var f = R / (1.5 - 1);
		lay.name = "#" + (i + 1) + ", 凸, 直径" + r_stock[i] + "mm, 焦点距離" + f + "mm, 厚さ" + 2 + "mm, 曲率半径" + R + "mm";
	}
}


for (i = 0; i < r.length; i++) {
	var layerNum = activeDocument.layers.length;
	layerObj = docObj.layers;
	for (j = 0; j < layerNum; j++) {
		layerName = layerObj[j].name;
		var initial = layerName.slice(0, 1);
		if (initial == "#") {
			var kei = layerName.indexOf('径');
			var comma = layerName.indexOf('m');
			lensTipeNum = layerName.slice(kei + 1, comma); //レンズの種類数

			if (lensTipeNum == r[i]) {//直径が同じレイヤーに描画
				pObj = layerObj[j].pathItems.ellipse(-y[i] / 2.0+px(r[i])/2.0, x[i] / 2.0-px(r[i])/2.0, px(r[i]), px(r[i]));//円を描画
				pObj.filled = false;
				pObj.stroked = true;
			}
		}
	}
}

deleteLayer(docObj);


function px(a) {//mm→px
	return a * 72 / 25.4;
}

//空レイヤー削除関数
function deleteLayer(delDoc) {
	var delLen = delDoc.layers.length;
	for (var i = delLen - 1; i >= 0; i--) {
		var delLayer = delDoc.layers[i];
		var delObj = delLayer.pageItems.length;
		var initial = delLayer.name.slice(0, 1);
		if (delObj == 0 && initial != "印") {
			delLayer.locked = false;
			delLayer.remove();
		}
	}
}

//塗り、線なしの指定
function setWhiteSpot() {
    try {
        whiteSpot = activeDocument.swatches.getByName('RDG_WHITE');
        activeDocument.defaultFillColor = whiteSpot.color;
    } catch (e) {
        var whiteColor = new CMYKColor;
        whiteColor.cyan = 25;
        whiteColor.magenta = 25;
        whiteColor.yellow = 25;
        whiteColor.black = 0;

        var newSpot = activeDocument.spots.add();
        newSpot.name = 'RDG_WHITE';
        newSpot.colorType = ColorModel.SPOT;
        newSpot.color = whiteColor;

        whiteSpot = activeDocument.swatches.getByName('RDG_WHITE');
        activeDocument.defaultFillColor = whiteSpot.color;
    }
    activeDocument.defaultStroked = false;
}
