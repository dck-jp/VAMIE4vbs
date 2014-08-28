# VAMIE 4 vbs

VAMIE (VAMIE4vbs) is VBScript Auto Mation for Internet Explorer, inspired By PAMIE & Selenium

## コンセプト
- 非プログラマの方が、内部実装について深く理解しなくても、感覚的にIEの自動制御ができるように。
- プログラマの方が、PAMIE(Pytthon)やSAMIE(Perl)、Selenium へ/からの移行が、そんなに違和感なくできるように。

## (プログラマ向け)設計方針的な話
- 1ファイルで完結
- VBA版とメソッドを共通化
- より複雑な制御をしたい人向けには、RAWアクセス手段を提供
- (任意のJavaScriptコードの実行や、Document objectにアクセスするためのプロパティを用意)

## 使い方
+ 書こうとしているコード(eg. main.vbs) と同じディレクトリに VAMIE.vbs を置きます
+ main.vbsに Sample.vbsのImport関数をコピーします
+ あとはUsage見て IEの制御部分を書いてね

## 使用例 （リファレンス代わり）
Please see the [Sample.vbs](https://github.com/dck-jp/VAMIE4vbs/blob/master/Sample.vbs).  
使い方は、[Sample.vbs](https://github.com/dck-jp/VAMIE4vbs/blob/master/Sample.vbs)を見てください。

## License
This source code is under [MIT License](https://github.com/dck-jp/VBAFramework/blob/master/LICENSE)  
ソースコードは[MIT License](https://github.com/dck-jp/VBAFramework/blob/master/LICENSE)で配布しています。

てきとうによろしくどうぞー。

## Author
D*isuke YAMAKAWA @ [ClockAhead](http://www.clockahead.com/)
