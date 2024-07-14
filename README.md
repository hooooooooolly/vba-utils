# vba-utils
よく使うVBAのライブラリ置き場

## モジュール名：CmnUtils
共通モジュール

### public Sub initAllSheets
拡大率100%で全シートの初期化を行う.

### public Sub initAllSheetsWithSpecified
任意の拡大率で全シートの初期化を行う.
※デフォルト値: 130%

### public Sub dispNameDefinition
複製を重ねた資料のシートをコピーした際に
>名前'あああ'は既に存在します。この名前にする場合は [はい] をクリックします。移動またはコピーを行うために'あああ'の名前を変更する場合は、 [いいえ] をクリックします。

といったアラートが出るが、名前の定義を見てもなぜか非表示になっているため強制的に表示させるマクロ.
