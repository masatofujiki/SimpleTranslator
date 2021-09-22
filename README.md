# SimpleTranslator(英語翻訳ツール)

Excelに貼り付けた外国語をSelenium BasicとDeepLでお手軽に翻訳してくれるツールを作りました。

環境設定はお手軽ではないですが、一度、環境設定を済ませばChrome Driverの更新以外は設定不要です。

## ■背景
現在は特許関連の仕事をしており、検索システムを利用して特許文献を検索しています。

検索結果として出力される特許文献は日本語だけでなく英語などの外国語の特許語文献もあり、

外国語の特許文献を読んで技術内容を理解することがあります。

## ■課題
外国語の特許文献の理解を補助するために検索システムには翻訳ツールが実装されていますが、

用意された翻訳ツールの翻訳精度が低く、外国語の特許文献の翻訳結果を理解するのに時間が掛かります。


## ■目的
翻訳ツールを利用して外国語の原文とDeepLで翻訳した翻訳文とを並べて表示することにより、

精度が高い翻訳文と原文とを比較できる翻訳ツールを提供することです。


## ■必要なもの
### OS
- Microsoft Windows 10
### ソフトウェア
- Google Chrome
- Microsoft Excel 2016 または Microsoft Excel 2019
- Selenium Basic
- Chrome Driver

## ■Selenium Basicを動作させるまでの環境設定
Selenium Basicがインストールされ、動作している環境なら以下は必要ありません。次に進んでください。
1. [SeleniumBasicをインストールしてExcel(VBA)からWebスクレイピングを行うまでのチュートリアル][a]のページを開きます。
2. [上記記事][a]を参照してSelenium Basicをダウンロード＆インストールします。
3. [上記記事][a]を参照してChrome Driverのダウンロードし、ダウンロードしたChrome DriverをSelenium Basicをインストールしたフォルダにコピー＆上書きします。
4. [上記記事][a]を参照してSelenium Basicの動作確認を行います。
5. 「.Net Fremework 3.5」がインストールされていないために起こるエラーが発生したときは、[上記記事][a]の一番下のリンクを参照して「.Net Fremework 3.5」をインストールします。
6. ここで「.Net Fremework 3.5」をインストールできないときは次の8.の手順を実行します。
7. [Windows10に.NET3.5をインストールする方法！][b]のページを参照して「.Net Fremework 3.5」をインストールします。変更したレジストリの値を戻すまできっちりやってください。

## ■Simple Translatorの使い方
1. 最新版:CrazyTranslator.zip
2. SimpleTranslator.xlsmを開きます。
3. 表示順「原文→翻訳文」「翻訳文→原文」のどちらにするかのラジオボタンを選択します。
4. ブラウザに表示された英文をコピーしExcelのシートA2に貼り付けます。
5. 翻訳(HTML)ボタンを押します。
6. プログレスバーが表示されるのでしばらく待ちます。
7. 翻訳が完成するとブラウザが立ち上がり翻訳結果が出力されます。
8. 翻訳結果はHTMLの形式でアプリケーションと同じ場所のディレクトリに作成されます。

## ■注意
※Google ChromeがアップデートされるとChrome Driverが動かなくなります。

※このときは[上記記事][a]を参照してGoogle Chromeのバージョンに合わせたChrome Driverを上書き更新してください。

## ■ダウンロードページ
1. [Selenium Basic](https://florentbr.github.io/SeleniumBasic/)
2. [Chrome Driver](https://chromedriver.chromium.org/downloads)

## ■参考ページ
1. [SeleniumBasicをインストールしてExcel(VBA)からWebスクレイピングを行うまでのチュートリアル][a]
2. [Windows10に.NET3.5をインストールする方法！][b]

[a]:https://lil.la/archives/3436
[b]:https://bgt-48.blogspot.com/2019/04/windows10net35.html
