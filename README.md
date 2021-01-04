<!-- markdownlint-disable MD033 -->
<!-- markdownlint-disable MD041 -->
# Mobile Environment Creation Support Script<br>
<p>
<img src="https://2.bp.blogspot.com/-n4C4AJ2oZpY/VtogDwW202I/AAAAAAAA4cI/TYjhMG6L0_U/s800/usb_memory_stick.png" width=60 height=60 >
<span style="font-family: Hiragino Maru Gothic ProN; font-size: 40pt; font-weight: bold; color: pink;">
携帯環境作成支援スクリプト
</span></p>
USBメモリや仮想ディスクなど携帯可能な環境でポータブル版のソフトウェアを活かすための環境作成をお手伝いします。<br>
We can help you create an environment to take advantage of the portable version of the software in a portable environment such as USB memory or virtual disk.<br><br>

## ⧉できること
- スクリプトが実行されたドライブまたは指定したドライブを基準にした環境設定ができます。
- 設定した環境変数はシャットダウン後に設定前の状態に戻ります。
- デスクトップにショートカットを作成できます。
- 相対パス指定のシンボリックリンクを簡単に作成できます。
- 仮想ディスクファイル（VHD, VHDX）を指定ドライブにマウント、アンマウントができます。
- プログラムの実行ができます。
- 引数またはファイルをドラックアンドドロップして実行すると引数を実行、または引数を指定してスクリプトを実行できます。

## ⧉What I can do.
- You can set the environment based on the drive where the script was executed or the drive you specify.
- Environment variables you set will be restored to their previous state after shutdown.
- Allows you to create shortcuts on the desktop.
- You can easily create symbolic links with relative paths.
- Virtual disk files (VHD, VHDX) can be mounted and unmounted on the specified drive.
- Execute programs.
- Drag and drop arguments or files to execute them, or specify arguments to execute a script.

## ⧉使用方法
- "MECSS_jp.cmd"、"MECSS_en.cmd"が実行ファイルです。
- 実行すると最初に設定ファイルが作成されます。
- 設定ファイルに使用方法と使用例が記載されています。
## ⧉How to use.
- "MECSS_jp.cmd" and "MECSS_en.cmd" are the executable files.
- When you run them, the configuration file will be created first.
- Configuration file contains usage instructions and examples.

## ⧉スクリプトについて
- スクリプトはバッチファイル形式とJScriptのミックスで書かれています。
- 作成した関数はそれぞれ単独でも動作できるよう、なるべく汎用的に作成しました。
- 関数には実装しているが使用していない機能や使用していない関数もあります。
- たとえば仮想ディスクの最適化でデフラグ後、バイナリ"00"のデータを空きスペースに書き込むと効果が高くなると聞いて作成したBase64を利用してバイナリのファイルを作成する関数は、実際にはほとんど効果がなかったのでコマンドの実装を見送りましたが、スクリプトは残しています。バイナリデータを書き込む機能を少し変更すれば[データ完全消去][]っぽいこともでます。
- JavaScript(JScript)で作成したプログラムの2作目であるため、かなり稚拙なプログラムであることは理解しています。今後ものんびり精進いたします。自由に改編してご利用いただければ幸いです。

## ⧉About Script
- The script is written in a mix of batch file format and JScript.
- The functions created are as generic as possible so that they can work independently of each other.
- Some functions are implemented but not used, and some functions are not used.
- For example, the function to create a binary file using Base64, which was created after hearing that writing binary "00" data to free space after defragmentation in virtual disk optimization is more effective, was not implemented because it had little effect in practice.The script is still there.If you modify the function to write binary data a little, you can do something like complete data erasure.
- As this is the second program I have created in JavaScript (JScript), I understand that it is a very poor program.I will continue to work on it at my leisure. I hope you will feel free to modify it and use it as you wish.

## ⧉ライセンス -License-
Copyright(c) 2021 Aromatibus (https://github.com/Aromatibus)<br>
Released under the MIT license. See https://opensource.org/licenses/MIT<br>
Translated with www.DeepL.com/Translator (free version)<br>


[データ完全消去]: https://ja.wikipedia.org/wiki/データの完全消去
