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
- The environment variables you set will be restored to their previous state after shutdown.
- Allows you to create shortcuts on the desktop.
- You can easily create symbolic links with relative paths.
- Virtual disk files (VHD, VHDX) can be mounted and unmounted on the specified drive.
- Execute programs.
- Drag and drop arguments or files to execute them, or specify arguments to execute a script.


## ⧉使用方法
- "MECSS_jp.cmd"、"MECSS_en.cmd"が実行ファイルです。
- 実行すると最初に設定ファイルが作成されます。
- 設定ファイルに使用方法と使用例が記載されています。
- スクリプトはバッチファイル形式とJScriptのミックスで書かれています。
- 関数はそれぞれ単独でも動作できるよう、なるべく汎用的に作成しましたので自由に改編して使用してください。

## ⧉How to use.
- "MECSS_jp.cmd" and "MECSS_en.cmd" are the executable files.
- When you run them, the configuration file will be created first.
- The configuration file contains usage instructions and examples.
- The script is written in a mix of batch file format and JScript.
- The functions have been created to be as generic as possible so that they can work independently, so feel free to modify and use them as you like.

## ⧉スクリプトについて
JavaScript(JScript)で作成したプログラムの2作目であるため、かなり稚拙なプログラムであることは理解しています。自由に改編してください。

## ⧉ライセンス
Copyright(c) 2021 Aromatibus (https://github.com/Aromatibus)<br>
Released under the MIT license. See https://opensource.org/licenses/MIT<br>
Translated with www.DeepL.com/Translator (free version)<br>
