ExcelSudoku
===========

Excel で[数独（ナンプレ）](https://ja.wikipedia.org/wiki/%E6%95%B0%E7%8B%AC)を解く実装例。  
[数独をExcelのVBAで解くのをやってみた](https://gist.github.com/furyutei/192911043d16f4793c7f21655482aca1)  で公開していたソースを元に整理＆クラスモジュール化したもの。  

使い方
------
サンプルの Excel ファイルを開き、左上9x9のセル範囲(A1:I9)に解きたい数独の問題を入力して、[開始]ボタンを押す。  
※なお、当該セル範囲で文字色が青(vbBlue)になっている場合、そのセルはクリアされるため注意。  

パフォーマンスについて
----------------------
数独の問題によってばらつきはあるが、ほぼ[m-haketa/suudoku: 数独を解くプログラム](https://github.com/m-haketa/suudoku)や[Excelで数独解析ソフト](http://excel.syogyoumujou.com/freesoft/analysis_sudoku.html)と同等程度のパフォーマンスは出るようになった……と思ったが、測定条件が良かっただけで、実際はこれらよりもかなり遅い模様。なかなか難しい……。  

[250問連続解析サンプルソフト](http://excel.syogyoumujou.com/freesoft/analysis_sudoku.html#saisoku)と同じことを自分の環境で実施した場合、  

| |本実装|[250問連続解析サンプルソフト](http://excel.syogyoumujou.com/freesoft/analysis_sudoku.html#saisoku)|
|:-:|:------:|:-------------------------:|
|[250問連続解析サンプルソフト](http://excel.syogyoumujou.com/freesoft/analysis_sudoku.html#saisoku)上の問題|約0.59秒|約0.60秒|
|[ナンプレ京（数独）](http://nanpre.adg5.com/index.php)上の問題(Lv1(200問)～2(50問))|約0.31秒|約0.15秒|
|[ナンプレ京（数独）](http://nanpre.adg5.com/index.php)上の問題(Lv6(50問)～7(200問))|約2.94秒|約0.90秒|

のような結果になった（2018/06/13現在）。  
※実効環境： Intel(R) Core(TM) i7-3820QM CPU @ 2.70GHz / Windows 10 Pro(64ビット) / Excel 2010(32ビット)   

その他
------
- ソースコードは、Excel の VBA からエクスポートしたもの（文字コード：シフトJIS・改行：CR+LF）。  
- 250問連続解析については、[250問連続解析サンプルソフト](http://excel.syogyoumujou.com/freesoft/analysis_sudoku.html#saisoku)で使用されている問題の入った Excel ファイルは著作権の関係で掲載できない。
  代替として、[ナンプレ京（数独）無料パズルゲーム 10000問以上](http://nanpre.adg5.com/index.php) 上の[数独問題を CSV ファイルとしてダウンロードできるスクリプトを用意した（使い方はスクリプトソース内に記載）](https://github.com/furyutei/ExcelSudoku/blob/master/src/js/DownloadSudoku256Csv.js)。  
  これでダウンロードしたCSVファイルを、[ExcelSudokuTry250.xlsm](https://github.com/furyutei/ExcelSudoku/blob/master/ExcelSudokuTry250.xlsm)に読ませることで、250問連続解析が実行できる。  

参考
----
- [数独（ナンプレ）を解くアルゴリズムの要点とパフォーマンスの検証№1｜VBAサンプル集](https://excel-ubara.com/excelvba5/EXCELVBA231.html)  
- [エクセルのサンプルダウンロード｜エクセルの神髄](https://excel-ubara.com/excel_download.html)  
- [m-haketa/suudoku: 数独を解くプログラム](https://github.com/m-haketa/suudoku)  
- [Excelで数独解析ソフト](http://excel.syogyoumujou.com/freesoft/analysis_sudoku.html)  

ライセンス
----------
[The MIT License](https://github.com/furyutei/ExcelSudoku/blob/master/LICENSE)  
