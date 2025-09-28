# vba-utils

よく使うVBAのライブラリ置き場

## モジュール名：SheetUtils

シート操作ユーティリティを集めたモジュールです。  

### public Sub initAllSheets

全シートの初期化を行います。

- 引数  
SheetProperties: シート用設定値を持つインスタンス

- 返り値  
なし

### public Sub dispNameDefinition
複製を重ねた資料のシートをコピーした際に  
>名前'あああ'は既に存在します。この名前にする場合は [はい] をクリックします。移動またはコピーを行うために'あああ'の名前を変更する場合は、 [いいえ] をクリックします。

といったアラートが出るが、名前の定義を見てもなぜか非表示になっているため強制的に表示させます。  

- 引数  
なし

- 返り値  
なし

　  



## クラス名: FileHandler

ファイル操作用モジュールです。  
FSOをラップしてOOPっぽい使い方ができるようになっています。

|プロシージャ種別|プロシージャ名|引数|返り値|説明|
|----|----|----|----|----|
|Sub|init|String: ファイルパス|-|コンストラクタ代わりに利用します。|
|Sub|openFile|WriteMode: 書き込みモード|-|ファイルを開きます。|
|Sub|closeFile|-|-|ファイルを閉じます。|
|Function|getFileName|-|String: ファイル名|ファイル名のgetterです。|
|Function|exists|-|Boolean: 存在する場合true|ファイルの存在チェックを行います。|
|Sub|delete|-|-|ファイルを削除します。|
|Function|readAll|-|String: ファイルの中身|ファイル全体を読み込みます。<br>ファイルサイズに注意。|
|Function|readLine|-|String: ファイルの中身|ファイルを1行ずつ読み込みます。<br>ループ内での利用を想定。|
|Sub|writeRaw|String: 書き込む内容|-|末尾の改行なしで書き込みます。|
|Sub|writeLine|String: 書き込む内容|-|末尾の改行ありで書き込みます。|
|Sub|appendLine|String: 書き込み内容|-|追記モードで書き込みます。|

　  



## クラス名: SheetProperties

シート用設定値を保持するクラスモジュールです。  

### フィールド

|プロパティ|説明|デフォルト値|
|----|----|----|
|zoomRate|拡大率|0(変更しない)|
|scrollColumn|列のスクロール位置|1|
|scrollRow|行のスクロール位置|1|
|displayPageBreaks|改ページ罫線の表示可否|False|
|displayGridlines|罫線の表示可否|False|
