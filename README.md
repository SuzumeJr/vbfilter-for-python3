vbfilter-for-python3
====================
VisualBasic6.0以前用Doxygenフィルターです。  
python2.0のvbfilter.pyをpython3.0で動くようにし、  
独自の機能を追加しています。
Python3.2での動作確認は出来ています。  

####使い方
INPUT_FILTERに'vbfilter.py'と指定するか'vbfilter.py C 'と指定して下さい  

####注意
日本語基準になっているので、他の言語の方はコード内の下記部分を  
自国で使用している文字コードに合わせて変更して下さい  
>src_encoding = "cp932"
