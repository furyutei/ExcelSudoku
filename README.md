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
数独の問題によってばらつきはあるが、ほぼ[m-haketa/suudoku: 数独を解くプログラム](https://github.com/m-haketa/suudoku)や[Excelで数独解析ソフト](http://excel.syogyoumujou.com/freesoft/analysis_sudoku.html)と同等程度のパフォーマンスは出るようになった。  

[250問連続解析サンプルソフト](http://excel.syogyoumujou.com/freesoft/analysis_sudoku.html#saisoku)と同じことを自分の環境で実施した場合、  

|本実装  |250問連続解析サンプルソフト|
|:------:|:-------------------------:|
|約0.49秒|約0.62秒                   |

のような結果になった。
※実効環境： Intel(R) Core(TM) i7-3820QM CPU @ 2.70GHz / Windows 10 Pro(64ビット) / Excel 2010(32ビット)   

その他
------
- ソースコードは、Excel の VBA からエクスポートしたもの
- 250問連続解析については、[250問連続解析サンプルソフト](http://excel.syogyoumujou.com/freesoft/analysis_sudoku.html#saisoku)で使用されている問題の入った Excel ファイルは著作権の関係で掲載できないため、ソースコード(使用しているのは[Try250.bas](https://github.com/furyutei/ExcelSudoku/blob/master/src/Modules/Try250.bas)と[ClassSudoku.cls](https://github.com/furyutei/ExcelSudoku/blob/master/src/ClassModules/ClassSudoku.cls))のみ公開としている。  

参考
----
- [数独（ナンプレ）を解くアルゴリズムの要点とパフォーマンスの検証№1｜VBAサンプル集](https://excel-ubara.com/excelvba5/EXCELVBA231.html)  
- [エクセルのサンプルダウンロード｜エクセルの神髄](https://excel-ubara.com/excel_download.html)  
- [m-haketa/suudoku: 数独を解くプログラム](https://github.com/m-haketa/suudoku)  
- [Excelで数独解析ソフト](http://excel.syogyoumujou.com/freesoft/analysis_sudoku.html)  

ライセンス
----------
[The MIT License](https://github.com/furyutei/ExcelSudoku/blob/master/LICENSE)  
