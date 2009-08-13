/*
regex_cellstyle.jsx
(c)2009 市川せうぞー
選択しているセル中の段落が正規表現にマッチしたら、指定のセルスタイルを適用します。

2009-08-10	ver.0.1	とりあえず。セルスタイルグループには対応しない。
2009-08-12	ver.0.2	セルスタイルグループに対応　http://d.hatena.ne.jp/seuzo/20090814/1250176724
	
*/

////////////////////////////////////////////エラー処理 
function myerror(mess) { 
  if (arguments.length > 0) { alert(mess); }
  exit();
}

////////////////////////////////////////////ダイアログ（テキストフィールドとポップアップ）
function popupDialog(my_Title, my_inputPrompt, my_input_default, my_popupPrompt, my_popupArray){
	var myDialog, my_popup_selectIndex, my_input_str;
	var ans_array = new Array();
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;//ダイアログの表示を有効に
	myDialog = app.dialogs.add({name:my_Title,canCancel:true});
	with(myDialog){
		with(dialogColumns.add()){
			with(borderPanels.add()){
				with(dialogColumns.add()){
					staticTexts.add({staticLabel:my_inputPrompt});// プロンプト
				}
				with(dialogColumns.add()){
					my_input_str = textEditboxes.add({editContents:my_input_default, minWidth:180});// テキストフィールド
				}
			}
			with(borderPanels.add()){
				with(dialogColumns.add()){
					staticTexts.add({staticLabel:my_popupPrompt});// プロンプト
				}
				with(dialogColumns.add()){
					my_popup_selectIndex = dropdowns.add({stringList:my_popupArray, selectedIndex:0});// ポップアップメニュー
				}
			}
		}
	}
	// ダイアログボックスを表示
	if(myDialog.show() === true){
		ans_array.push(my_input_str.editContents);
		ans_array.push(my_popup_selectIndex.selectedIndex);
		myDialog.destroy();//正常にダイアログを片付ける
		return ans_array//配列で返す
	} else {
		// ユーザが「キャンセル」をクリックしたので、メモリからダイアログボックスを削除
		myDialog.destroy();
		exit();//スクリプト終了
	}
}

////////////////////////////////////////////文字列を16進数にエスケープして、「\x{hex}」という形で返す。
function my_escape(str) {
	tmp_str = escape(str);
	return tmp_str.replace(/\%u([0-9A-F]+)/i, "\\x{$1}")
}
////////////////////////////////////////////正規表現検索でカタカナを16進に変換
function katakana2hex() {
	var find_str = app.findGrepPreferences.findWhat;
	while (/([ァ-ヴ])/.exec(find_str)) {
		find_str = find_str.replace(/([ァ-ヴ])/, my_escape(RegExp.$1));
	}
	app.findGrepPreferences.findWhat = find_str;//検索文字の設定
}

////////////////////////////////////////////正規表現検索
/*
my_range	検索置換の範囲
my_find	検索オブジェクト ex.) {findWhat:"(わたく?し|私)"}
my_change	置換オブジェクト ex.)  {changeTo:"ぼく"}
*/
function my_RegexFindChange(my_range, my_find) {
	//検索の初期化
	app.findGrepPreferences = NothingEnum.nothing;
	app.changeGrepPreferences = NothingEnum.nothing;
	//検索オプション
	app.findChangeGrepOptions.includeLockedLayersForFind = false;//ロックされたレイヤーをふくめるかどうか
	app.findChangeGrepOptions.includeLockedStoriesForFind = false;//ロックされたストーリーを含めるかどうか
	app.findChangeGrepOptions.includeHiddenLayers = false;//非表示レイヤーを含めるかどうか
	app.findChangeGrepOptions.includeMasterPages = false;//マスターページを含めるかどうか
	app.findChangeGrepOptions.includeFootnotes = false;//脚注を含めるかどうか
	app.findChangeGrepOptions.kanaSensitive = true;//カナを区別するかどうか
	app.findChangeGrepOptions.widthSensitive = true;//全角半角を区別するかどうか

	app.findGrepPreferences.properties = my_find;//検索の設定
	//app.changeGrepPreferences.properties = my_change;//置換の設定
	
	if (parseInt(app.version) === 5) {
		katakana2hex();//カタカナにマッチしないバグ回避
	}

	//my_range.changeGrep();//検索と置換の実行
	return my_range.findGrep();//検索のみの場合：マッチしたオブジェクトを返す
}


////////////////////////////////////////////以下メイン処理
//変数の宣言
var my_doc, my_selection, my_class, my_cellStyles, my_target_CellStyle, my_find_items, i, myError;
var my_cellStyles_names = new Array();

//InDesignで選択しているもののチェック
if (app.documents.length === 0) {myerror("ドキュメントが開かれていません")}
my_doc = app.activeDocument;
if (my_doc.selection.length === 0) {myerror("セルを選択してください")}
my_selection = my_doc.selection[0];
my_class =my_selection.constructor.name;
if ((my_class !== "Cell") && (my_class !== "Table")) {
	myerror("検索対象になるセルを選択してください");
}

//セルスタイルを取得
my_cellStyles = my_doc.allCellStyles;//ドキュメント中の全てのセルスタイルのリスト（オブジェクト）
for (i =0; i< my_cellStyles.length; i++) {
	my_cellStyles_names.push(my_cellStyles[i].name);//（名前のリスト）
}

//ダイアログを表示
var ans = popupDialog("正規表現でセルスタイル適用", "正規表現：", "^[0-9]+", "セルスタイル：", my_cellStyles_names);
my_target_CellStyle = my_cellStyles[ans[1]];//ターゲットセルスタイル（オブジェクト）を確定
//検索
my_find_items = my_RegexFindChange(my_selection, {findWhat:ans[0]});

//セルスタイル適用
for (i = 0; i < my_find_items.length; i++) {
	try {
		if (my_find_items[i].parent instanceof Cell) {
			my_find_items[i].parent.appliedCellStyle = my_target_CellStyle;
		}
	} catch (myError) {
		myerror("エラー：" + myError);
	}
}

alert("" + my_find_items.length + "箇所のセルに「" + my_cellStyles_names[ans[1]] + "」セルスタイルを適用しました。");