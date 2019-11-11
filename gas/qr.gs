/*
Google Spreadsheet上にあるURLに対して
- 短縮URLに変換
- QRコードを生成
- Google Driveの指定のフォルダ内に画像として保存
- フォルダ内の画像のリンクの書き込み
を行う

- Google DriveのフォルダのフォルダID
- bit.ly APIのアクセストークン
の取得が必要

Spreadsheetの
- B列の一部にid
- D列にtableId
- E列にURL
があり，idとtableIdをもとに画像のファイル名が決定され，F列に画像のリンクが書き込まれるという仕様
 */

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    var menu = ui.createMenu("Scripts")
        .addItem("QR Code Gen", "main")
        .addToUi();
}

function main() {

    const FOLDER_ID = "";  // https://drive.google.com/drive/folders/<FOLDER_ID>
    const QR_SIZE = 200;

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    var folder = DriveApp.getFolderById(FOLDER_ID);

    var values = sheet.getRange(1, 1, 100, 6).getValues();  // A列2行目からF列100行まで取得

    for(var row = 2-1; row < 100-1; row++){
        // E列にURLがあり，F列にQRコード画像を生成した場所のURLがなかった場合(=まだ生成されていない場合)のみ生成する
        if(values[row][5-1] != null && values[row][5-1] != ""){
            if(values[row][6-1] == null || values[row][6-1] == ""){
                // B列に値があればそのままidとして取得，なければ上の行を確認
                var id = "";
                var _row = row;
                while(_row > 0){
                    if(values[_row][2-1] != null && values[_row][2-1] != "") {
                        id = values[_row][2-1];
                        break;
                    }
                    _row--;
                }
                var tableId = values[row][4-1]; // D
                var url = values[row][5-1];  // E
                var shortenUrl = getShortenURL(url);
                var data = UrlFetchApp.fetch("http://chart.apis.google.com/chart?chs=" + QR_SIZE + "x" + QR_SIZE + "&cht=qr&chl=" + encodeURIComponent(shortenUrl));
                var image = data.getBlob().getAs("image/png").setName(String(id) + "_" + String(tableId) + ".png");
                var file = folder.createFile(image);
                var fileUrl = file.getUrl();
                sheet.getRange(row+1, 6).setValue(fileUrl);
            }
        }
    }
}

function getShortenURL(url){

    const ACCESS_TOKEN = "";

    var encodedUrl = encodeURIComponent(url);
    var options =
        {
            "method" : "post",
            "payload" : {}
        };

    var response = UrlFetchApp.fetch("https://api-ssl.bitly.com/v3/shorten?access_token="+ACCESS_TOKEN+"&longUrl=" + encodedUrl, options);
    var content = response.getContentText("UTF-8");
    var body = JSON.parse(content);

    console.log("Get shorten URL: " + body.data.url);
    return body.data.url;
}
