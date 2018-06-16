ExcelSudoku
===========

Excel で[数独（ナンプレ）](https://ja.wikipedia.org/wiki/%E6%95%B0%E7%8B%AC)を解く実装例。  
[数独をExcelのVBAで解くのをやってみた](https://gist.github.com/furyutei/192911043d16f4793c7f21655482aca1)  で公開していたソースを元に整理＆クラスモジュール化したもの。  

使い方
------
[サンプルの Excel ファイル](https://github.com/furyutei/ExcelSudoku/blob/master/ExcelSudoku.xlsm)を開き、左上9x9のセル範囲(A1:I9)に解きたい数独の問題を入力して、[開始]ボタンを押す。  
※[リセット]ボタンを押すと、文字色が青(vbBlue)のセル（試行セル）がクリアされる。  

パフォーマンスについて
----------------------
数独の問題によってばらつきはあるが、ほぼ[m-haketa/suudoku: 数独を解くプログラム](https://github.com/m-haketa/suudoku)や[Excelで数独解析ソフト](http://excel.syogyoumujou.com/freesoft/analysis_sudoku.html)と同等程度のパフォーマンスは出るようになった……と思ったが、測定条件が良かっただけで、実際はこれらよりもかなり遅い模様。なかなか難しい……。  
→ Version 0.0.1.4 になって、ようやく安定した結果がでるようになった、かも。  

[250問連続解析サンプルソフト](http://excel.syogyoumujou.com/freesoft/analysis_sudoku.html#saisoku)と同じことを自分の環境で実施した場合、  
※実効環境： Intel(R) Core(TM) i7-3820QM CPU @ 2.70GHz / Windows 10 Pro(64ビット) / Excel 2010(32ビット)   

| |本実装 Version 0.0.1.6|[Excelで数独解析ソフトVer.2.3.0](http://excel.syogyoumujou.com/freesoft/dl_soft/analysis_sudoku/index.html)|
|:-:|:------:|:-------------------------:|
|[250問連続解析サンプルソフト](http://excel.syogyoumujou.com/freesoft/analysis_sudoku.html#saisoku)上の問題|約0.43秒|約0.56秒|
|[ナンプレ京（数独）](http://nanpre.adg5.com/index.php)上の問題(Lv2→1(200問＋50問))|約0.22秒|約0.18秒|
|[ナンプレ京（数独）](http://nanpre.adg5.com/index.php)上の問題(Lv3→2(200問＋50問))|約0.28秒|約0.21秒|
|[ナンプレ京（数独）](http://nanpre.adg5.com/index.php)上の問題(Lv4→3(200問＋50問))|約0.30秒|約0.31秒|
|[ナンプレ京（数独）](http://nanpre.adg5.com/index.php)上の問題(Lv5→4(200問＋50問))|約0.52秒|約0.58秒|
|[ナンプレ京（数独）](http://nanpre.adg5.com/index.php)上の問題(Lv6→5(200問＋50問))|約0.62秒|約0.78秒|
|[ナンプレ京（数独）](http://nanpre.adg5.com/index.php)上の問題(Lv7→6(200問＋50問))|約0.72秒|約0.88秒|

のような結果になった（2018/06/16現在）。  
Lv4～5辺りで、[Excelで数独解析ソフト](http://excel.syogyoumujou.com/freesoft/analysis_sudoku.html)と成績が逆転している感じか？  

なお、250問連続解析については、[250問連続解析サンプルソフト](http://excel.syogyoumujou.com/freesoft/analysis_sudoku.html#saisoku)で使用されている問題の入った Excel ファイルは著作権の関係で掲載できない。  
代替として、[ナンプレ京（数独）無料パズルゲーム 10000問以上](http://nanpre.adg5.com/index.php) 上の[数独問題を CSV ファイルとしてダウンロードできるスクリプトを用意した（使い方はスクリプトソース内に記載）](https://github.com/furyutei/ExcelSudoku/blob/master/src/js/DownloadSudoku250Csv.js)。  
これでダウンロードしたCSVファイルを、[ExcelSudokuTry250.xlsm](https://github.com/furyutei/ExcelSudoku/blob/master/ExcelSudokuTry250.xlsm)に読ませることで、250問連続解析が実行できる。  

その他
------
- ソースコードは、Excel の VBA からエクスポートしたもの（文字コード：シフトJIS・改行：CR+LF）。  

参考
----
- [数独（ナンプレ）を解くアルゴリズムの要点とパフォーマンスの検証№1｜VBAサンプル集](https://excel-ubara.com/excelvba5/EXCELVBA231.html)  
- [エクセルのサンプルダウンロード｜エクセルの神髄](https://excel-ubara.com/excel_download.html)  
- [m-haketa/suudoku: 数独を解くプログラム](https://github.com/m-haketa/suudoku)  
- [Excelで数独解析ソフト](http://excel.syogyoumujou.com/freesoft/analysis_sudoku.html)  

ライセンス
----------
[The MIT License](https://github.com/furyutei/ExcelSudoku/blob/master/LICENSE)  
