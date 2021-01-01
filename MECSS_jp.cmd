@if(0)==(0) rem /* Start of Batch
@echo off
setlocal
call :SetESC
title 携帯環境作成支援スクリプト
echo.
echo %ESC%[102m携帯環境作成支援スクリプト%ESC%[0m
echo.
echo USBメモリなど携帯環境でポータブルソフトを活かすための環境作成をお手伝いします。
echo Copyright(c) 2021 Aromatibus (https://github.com/Aromatibus)
echo.
:: Switch to JScript mode
cscript.exe //nologo //E:JScript "%~f0" %*
echo.
if %errorlevel% equ 0 (
  echo %ESC%[41;97mスクリプトにエラーが発生しました。%ESC%[0m
  echo.
  pause
  goto :eof
)
if %errorlevel% equ 1 (
  echo %ESC%[102m処理が終了しました。%ESC%[0m
  echo.
  timeout /t 5 /nobreak
) else (
  echo %ESC%[105m
  echo.
  timeout /t 2 /nobreak
)
goto :eof
:SetESC
for /F "tokens=1,2 delims=#" %%a ^
in ('"prompt #$H#$E# & echo on & for %%b in (1) do rem"') ^
do (set ESC=%%b)
@exit /b
gap
rem End of Batch */
@end
/*** Start of JScript <JavaScript ECMAScript3?> ***/

// 実行状況をコマンドプロンプトのウインドウに表示するフラグ
var DebugFlg = true;
// 外部に設定ファイルを作らずに内部の情報で動作させるフラグ
var OneFileMode = false;

// Polyfill trim
if (!String.prototype.trim) {
  String.prototype.trim = function () {
    return this.replace(/^[\s　\uFEFF\xA0]+|[\s　\uFEFF\xA0]+$/g, '');
  };
}
// Polyfill indexOf
var NotFound = -1;
if (!Array.prototype.indexOf) {
  Array.prototype.indexOf = function(obj, start){
    for(var i = (start || 0), j = this.length; i < j; i++){
      if(this[i] === obj){ return i }
    }
    return NotFound;
  };
}
// Polyfill repeat
if (!String.prototype.repeat) {
  String.prototype.repeat = function(count) {
    return Array(count * 1 + 1).join(this);
  };
}
// Polyfill isArray
if (!Array.isArray) {
  Array.isArray = function(arg) {
    return Object.prototype.toString.call(arg) === '[object Array]';
  };
}

// Main
function main() {
  // Batch %errorlevel%
  var ErrorlevelABNormal = 0; // Script terminated abnormally.(not use)
  var ErrorlevelNormal = 1;
  var ErrorlevelRebootAdmin = 2;
  var ErrorlevelMakeConfig = 3;
  // オブジェクト変数
  var ws = new ActiveXObject("WScript.Shell");
  var wseProcess = ws.Environment("Process");
  var wseSystem = ws.Environment("System");
  var wseUser = ws.Environment("User");
  var wseVolatile = ws.Environment("Volatile");
  var wsn = new ActiveXObject("WScript.Network");
  var sfo = new ActiveXObject("Scripting.FileSystemObject");
  // 環境参照変数
  var RunningDrv = sfo.GetDriveName(WScript.ScriptFullName);
  var ReferenceDrv = RunningDrv;
  var RunningFolder = sfo.GetParentFolderName(WScript.ScriptFullName);
  var RunningFileName = sfo.GetBaseName(WScript.ScriptFullName);
  var ConfigFileName = RunningFileName + ".txt";
  var ScriptArgArray = GetArrayArgs();
  var AdminFlg = isAdmin();
  // 定数
  var RegFound = 0; // 発見すると 0 以上になります。
  var RegNotFound = -1;
  var MODE = {
    READING: 1,
    WRITING: 2,
    APPENDING: 8,
    CREATE: true,
    NOT_CREATE: false,
    WAIT: true,
    NOT_WAIT: false
  }
  var WINSTYLE = {
    HIDDEN: 0,
    NORMAL: 1,
    MINIMUM: 2,
    MAXIMUM: 3,
    INACTIVE_NORMAL: 4,
    LAST_TIME: 5,
    INACTIVE_MINIMUM: 7
  }
  // 変数
  var ReadData;
  var CommandName;
  var ArgArray = [];
  var ArgLength;
  var RegEx;
  var TmpPATH = "";
  var ExeArgFlg = false;
  // カレントフォルダをスクリプトのフォルダにセット
  ws.CurrentDirectory = RunningFolder;
  // 設定ファイルの作成・読み込み
  if (OneFileMode) {
    ConfigText = CreateConfigFile("");
  } else {
    if (!sfo.FileExists(ConfigFileName)) {
      CreateConfigFile(ConfigFileName);
      return ErrorlevelMakeConfig;
    }
    var otf = sfo.OpenTextFile(ConfigFileName, MODE.READING);
    var ConfigText = otf.ReadAll();
    otf.Close();
  }
  // 設定ファイルを整形
  // ファイル末尾にCRLFを付加しておく
  ConfigText = ConfigText + "\r\n";
  // CRLFに統一
  ConfigText = ConfigText.replace(/(\r|\n|\r\n)/g, "\r\n");
  // 行頭、行末の空白除去(TRIM)
  ConfigText = ConfigText.replace(/(\r\n)[\s　\uFEFF\xA0]+/g, "\r\n");
  ConfigText = ConfigText.replace(/[\s　\uFEFF\xA0]+(\r\n)/g, "\r\n");
  // コメント行削除
  ConfigText = ConfigText.replace(/^#.*$(\r\n)/gm, "");
  // 空行削除
  ConfigText = ConfigText.replace(/^$(\r\n)/g, "");
  ConfigText = ConfigText.replace(/(\r\n)^$/m, "");
  // 優先するコマンドを実行
  RegEx = /^RunAdministrator\s*=\s*true$/gmi;
  if (ConfigText.search(RegEx) !== RegNotFound) {
    if (!(AdminFlg & 2)) {
      DebugMsg("Re-execute with administrative privileges.\n");
      ExecuteCommand(WScript.ScriptFullName, ScriptArgArray, "RunAs");
      return ErrorlevelRebootAdmin;
    }
  }
  // 設定情報を行ごとに配列に分解する
  ConfigText = ConfigText.split("\r\n");
  // プロセス環境変数の"PATH"の末尾に";"がない場合、";"が付ける
  if (wseProcess.item("PATH").slice(-1) !== "\;") { TmpPATH = "\;"; }
  // ユーザー権限、実行プロセス権限の情報表示
  DebugMsg("User Name    = " + wsn.UserName);
  if (AdminFlg & 1) {
    DebugMsg("User Account = Administrator");
  } else {
    DebugMsg("User Account = User");
  }
  DebugMsg("Computer     = " + wsn.ComputerName);
  if (AdminFlg & 2) {
    DebugMsg("Process      = Administrator\n");
  } else {
    DebugMsg("Process      = User\n");
  }
  // 引数情報表示
  if (ScriptArgArray.length === 0) {
    DebugMsg("Arguments nothing.\n");
  } else {
    DebugMsg("Arguments = [" + ScriptArgArray + "]\n");
  }
  // 環境変数情報表示
  DebugMsg("Reference Drive = \"" + ReferenceDrv + "\"\n");
  DebugMsg("Process PATH =\n[" + wseProcess.item("PATH") + "]\n");
  DebugMsg("System PATH =\n[" + wseSystem.item("PATH") + "]\n");
  DebugMsg("User PATH =\n[" + wseUser.item("PATH") + "]\n");
  DebugMsg("Volatile PATH =\n[" + wseVolatile.item("PATH") + "]\n");
  // 設定情報を配列ごとに処理
  for (var i=0; i < ConfigText.length; i++){
    ReadData = ConfigText[i];
//    DebugMsg("Command [" + NumberFormat(i, 2) + "] : " + ReadData);
    // 環境変数"PATH"へ変数を追加するための事前処理
    if (sfo.FolderExists(SameDrv(ReadData, ReferenceDrv))) {
      ReadData = SameDrv(ReadData, ReferenceDrv);
      RegEx = new RegExp(ReadData.replace(/\\/g, "\\\\"), "gi");
      if (wseProcess.item("PATH").search(RegEx) === RegNotFound) {
        TmpPATH = TmpPATH + ReadData + ";";
        DebugMsg("PATH Attached.   = [" + ReadData + "]");
      } else {
        DebugMsg("PATH Unattached. = [" + ReadData + "]");
      }
    }
    // Volatile環境変数の削除
    RegEx = /^RemoveVolatile\s*=\s*true$/i;
    if (ReadData.search(RegEx) === RegFound) {
      ws.Exec("cmd /c reg delete "
          + "\"HKEY_CURRENT_USER\\Volatile Environment\"");
      DebugMsg("Registry \"HKEY_CURRENT_USER\\Volatile Environment\""
          + " removed.");
    }
    // 名有りの環境変数の削除
    RegEx = /^(.*)\s*:=:$/gi;
    if (ReadData.search(RegEx) === RegFound) {
      ws.Exec("cmd /c reg delete "
          + "\"HKEY_CURRENT_USER\\Volatile Environment\" /v \""
          + ReadData.replace(RegEx, "$1")
          + "\" /f");
      wseVolatile.item(ReadData.replace(RegEx, "$1")) = "";
      wseProcess.item(ReadData.replace(RegEx, "$1")) = "";
      DebugMsg(ReadData.replace(RegEx, "$1") + "=(Removed)");
    }
    // 名有りの環境変数の追加
    RegEx = /^(.*)\s*:=:\s*(.*)$/gi;
    if (ReadData.search(RegEx) === RegFound) {
      ReadData = SameDrv(ReadData, ReferenceDrv);
      wseVolatile.item(ReadData.replace(RegEx, "$1"))
          = ReadData.replace(RegEx, "$2");
      wseProcess.item(ReadData.replace(RegEx, "$1"))
          = ReadData.replace(RegEx, "$2");
      DebugMsg("\n"
          + ReadData.replace(RegEx, "$1") + "="
          + ReadData.replace(RegEx, "$2"));
    }
    // 参照するドライブを変更
    RegEx = /^ReferenceDrive\s*=$/gi;
    if (ReadData.search(RegEx) === RegFound) {
      ReferenceDrv = RunningDrv;
      DebugMsg("Set the referenced drive to \"" + ReferenceDrv + "\".\n");
    }
    RegEx = /^ReferenceDrive\s*=\s*(\"|\')?([A-Z]):?\\?(\"|\')?$/gi;
    if (ReadData.search(RegEx) === RegFound) {
      RegEx = ReadData.replace(RegEx, "$2").toUpperCase() + ":";
      if (!sfo.FolderExists(RegEx)) {
        DebugMsg("The drive specified in "
            + "\"ReferenceDrive\" is not found."
            + " [" + RegEx + "]\n");
       } else {
        ReferenceDrv = RegEx;
        DebugMsg("Set the referenced drive to \"" + ReferenceDrv + "\"\n");
      }
    }
    // 指定秒数の待機処理
    RegEx = /^WaitSec\((\d+)\)$/gi;
    if (ReadData.search(RegEx) === RegFound) {
      DebugMsg("");
      WaitSecond(parseInt(ReadData.replace(RegEx, "$1")))
    }
    // デスクトップへショートカットを作成
    CommandName = "DesktopShortcut";
    ArgLength = 2;
    if (SplitCommand(CommandName, ReadData, ArgLength) !== false) {
      ArgArray = SplitCommand(CommandName, ReadData, ArgLength);
      ArgArray = SameDrv(ArgArray, ReferenceDrv);
      MakeDesktopShortcut(ArgArray[0], ArgArray[1]);
    }
    // シンボリックリンクを作成 ※プロセスの管理者権限が必要
    CommandName = "RelativeMKLink";
    ArgLength = 2;
    if (SplitCommand(CommandName, ReadData, ArgLength) !== false) {
      ArgArray = SplitCommand(CommandName, ReadData, ArgLength);
      ArgArray = SameDrv(ArgArray, ReferenceDrv);
      RelativeMKLink(ArgArray[0], ArgArray[1]);
    }
    // 仮想ディスクのマウント処理 ※プロセスの管理者権限が必要
    CommandName = "VDiskAttach";
    ArgLength = 2;
    if (SplitCommand(CommandName, ReadData, ArgLength) !== false) {
      ArgArray = SplitCommand(CommandName, ReadData, ArgLength);
      VDiskAttach(SameDrv(ArgArray[0], ReferenceDrv), ArgArray[1]);
    }
    // 仮想ディスクのアンマウント処理 ※要プロセスの管理者権限
    CommandName = "VDiskDetach";
    ArgLength = 1;
    if (SplitCommand(CommandName, ReadData, ArgLength) !== false) {
      ArgArray = SplitCommand(CommandName, ReadData, ArgLength);
      ArgArray = SameDrv(ArgArray[0], ReferenceDrv);
      VDiskDetach(ArgArray[0]);
    }
    // 指定文字列を実行
    CommandName = "Execute";
    ArgLength = 1;
    if (SplitCommand(CommandName, ReadData, ArgLength) !== false) {
      ArgArray = SplitCommand(CommandName, ReadData, ArgLength);
      ArgArray = SameDrv(ArgArray, ReferenceDrv);
      DebugMsg("\n" + CommandName + " : [" + ArgArray + "]");
      // 環境変数"PATH"へ揮発属性で追加
      if (TmpPATH !== "") {
        wseVolatile.item("PATH") = wseVolatile.item("PATH") + TmpPATH;
        wseProcess.item("PATH") = wseProcess.item("PATH") + TmpPATH;
        DebugMsg("\nAttach Volatile PATH = \n["
            + wseVolatile.item("PATH") + "]\n");
        TmpPATH = "";
      }
      // ファイルがあるか確認後、実行
      if (sfo.FileExists(ArgArray)) {
        ExecuteCommand(ArgArray);
      } else {
        DebugMsg("Execute : Error Execute File Not Found.\n");
      }
    }
    // 指定文字列を引数付きで実行
    CommandName = "ExecuteArg";
    ArgLength = 1;
    if (SplitCommand(CommandName, ReadData, ArgLength) !== false) {
      ArgArray = SplitCommand(CommandName, ReadData, ArgLength);
      ArgArray = SameDrv(ArgArray, ReferenceDrv);
      ExeArgFlg = true;
      DebugMsg("\n" + CommandName + " : [" + ArgArray + "]");
      // 環境変数PATHへ揮発属性で追加
      if (TmpPATH !== "") {
        wseVolatile.item("PATH") = wseVolatile.item("PATH") + TmpPATH;
        wseProcess.item("PATH") = wseProcess.item("PATH") + TmpPATH;
        DebugMsg("\nAttach Volatile PATH = \n["
            + wseVolatile.item("PATH") + "]\n");
        TmpPATH = "";
      }
      // ファイルがあるか確認後、実行
      if (sfo.FileExists(ArgArray)) {
        if (ScriptArgArray.length != 0) {
          ExecuteCommand(ArgArray, ScriptArgArray);
        } else {
          DebugMsg("No Arguments.\n");
          ExecuteCommand(ArgArray);
        }
      } else {
        DebugMsg("Error Execute File Not Found.\n");
      }
    }
  }
  // 環境変数PATHへ揮発属性で追加
  if (TmpPATH !== "") {
    wseVolatile.item("PATH") = wseVolatile.item("PATH") + ";" + TmpPATH;
    wseProcess.item("PATH") = wseProcess.item("PATH") + ";" + TmpPATH;
    DebugMsg("\nAttach Volatile PATH = \n["
        + wseVolatile.item("PATH") + "]\n");
  }
  // 引数があり、設定情報に "ExecuteArg "がない場合は、引数を実行します
  if ((ScriptArgArray.length > 0) && (!ExeArgFlg)) {
    if (ScriptArgArray.length > 1) {
      ExecuteCommand(ScriptArgArray[0], ScriptArgArray.slice(1));
    } else {
      ExecuteCommand(ScriptArgArray);
    }
  }
  return ErrorlevelNormal;
}

// デバッグ向け情報の表示
function DebugMsg(TextMessage, opt1, opt2) {
  DebugFlg = (typeof DebugFlg !== "undefined") ? DebugFlg : false;
  if (DebugFlg !== true) { return; }
  var ws = new ActiveXObject("WScript.Shell");
  var sfo = new ActiveXObject("Scripting.FileSystemObject");
  var FILE = {
    READING: 1,
    WRITING: 2,
    APPENDING: 8,
    CREATE: true,
    NOT_CREATE: false
  }
  ws.CurrentDirectory = sfo.GetParentFolderName(WScript.ScriptFullName);
  var opt = [];
  opt.push((typeof opt1 !== "undefined")
      ? opt1.toString().toUpperCase() : "");
  opt.push((typeof opt2 !== "undefined")
      ? opt2.toString().toUpperCase() : "");
  if(opt[0] === "FILE") {
    if (!sfo.FileExists(opt[1])) {
      wsOut = sfo.OpenTextFile(opt[1], FILE.WRITING, FILE.CREATE);
    } else {
      wsOut = sfo.OpenTextFile(opt[1], FILE.APPENDING, FILE.NOT_CREATE);
    }
    wsOut.WriteLine(TextMessage);
    wsOut.Close();
    return;
  }
  var wsOut = WScript.StdOut;
  if(opt.indexOf("STDERR") > NotFound) { wsOut = WScript.StdErr; }
  if(opt.indexOf("STDOUT") > NotFound) { wsOut = WScript.StdOut; }
  if(opt.indexOf("WRITE") > NotFound) {
    wsOut.Write(TextMessage);
  } else {
    wsOut.WriteLine(TextMessage);
  }
}

// 整数は３桁区切り指定小数未満は切捨て指定文字数に揃えます。数値以外は指定文字数に揃えます。
function NumberFormat(FloatNum, Digits, DecPoint) {
  DecPoint =
      (typeof DecPoint !== "undefined") ? DecPoint : 0;
  if (DecPoint > 0) { Digits = Digits - DecPoint - 1; }
      var RegEx = new RegExp(/^[+,-]?([1-9]\d*|0)(\.\d+)?$/);
  if (RegEx.test(FloatNum)) {
    var IntNum = Math.floor(parseFloat(FloatNum, 10));
    var IntNumStr = IntNum.toString()
        .replace( /(\d)(?=(\d\d\d)+(?!\d))/g, '$1,');
    var SmallNum = (parseFloat(FloatNum, 10) - IntNum); // Bugがある?
    var SmallNumStr = "";
    if (DecPoint !== 0) {
      SmallNumStr = ("." + SmallNum.toString().slice(2, DecPoint + 2)
                  + "0".repeat(DecPoint)).slice(0, DecPoint + 1);
    }
    var SameDigitChar = "";
    if ((Digits - IntNumStr.length) > 0) {
      SameDigitChar = " ".repeat(Digits - IntNumStr.length);
    }
    return (SameDigitChar + IntNumStr + SmallNumStr);
  } else {
    var SameDigitChar = " ".repeat(Digits - FloatNum.toString().length);
    return (SameDigitChar + FloatNum.toString());
  }
}

// バイト文字列を数値に変換、オプションで単位変換
function ConvByteStr(ByteStr, DataType) {
  DataType = (typeof DataType !== "undefined") ? DataType : "Byte";
  var RegEx = new RegExp("(bit|Byte|kB|KB|MB|GB|TB|KiB|MiB|GiB|TiB)", "i");
  if (!RegEx.test(DataType.toString())) { DataType = "Byte"; }
  var ByteUnits = {
    bit: 8,
    kB: 1000,
    MB: Math.pow(1000, 2),
    GB: Math.pow(1000, 3),
    TB: Math.pow(1000, 4),
    KB: Math.pow(2, 10),
    KiB: Math.pow(2, 10),
    MiB: Math.pow(2, 20),
    GiB: Math.pow(2, 30),
    TiB: Math.pow(2, 40)
  }
  var RegEx = new RegExp("^(\\d{1,3}(,?\\d{3})*(\\.\\d+)?)(\\s+)?"
            + "(kB|KB|MB|GB|TB|KiB|MiB|GiB|TiB)?$", "g");
  var ResultRegEx = RegEx.exec(ByteStr);
  try {
      var ByteValue = parseFloat(ResultRegEx[1], 10);
  } catch (error) {
    return false;
  }
  switch (ResultRegEx[ResultRegEx.length - 1]) {
    case "kB":
      ByteValue = ByteValue * ByteUnits.kB;
      break;
    case "KB":
      ByteValue = ByteValue * ByteUnits.KB;
      break;
    case "KiB":
      ByteValue = ByteValue * ByteUnits.KiB;
      break;
    case "MB":
      ByteValue = ByteValue * ByteUnits.MB;
      break;
    case "MiB":
      ByteValue = ByteValue * ByteUnits.MiB;
      break;
    case "GB":
      ByteValue = ByteValue * ByteUnits.GB;
      break;
    case "GiB":
      ByteValue = ByteValue * ByteUnits.GiB;
      break;
    case "TB":
      ByteValue = ByteValue * ByteUnits.TB;
      break;
    case "TiB":
      ByteValue = ByteValue * ByteUnits.TiB;
      break;
    case "":
      ByteValue = ByteValue;
      break;
    Default:
      return false;
  }
  // DataTypeに応じて戻り値を返します。
  switch (DataType) {
    case "bit":
      ByteValue = ByteValue * ByteUnits.bit;
      break;
    case "Byte":
      ByteValue = ByteValue;
      break;
    case "kB":
      ByteValue = ByteValue / ByteUnits.kB;
      break;
    case "KB":
      ByteValue = ByteValue / ByteUnits.KB;
      break;
    case "KiB":
      ByteValue = ByteValue / ByteUnits.KiB;
      break;
    case "MB":
      ByteValue = ByteValue / ByteUnits.MB;
      break;
    case "MiB":
      ByteValue = ByteValue / ByteUnits.MiB;
      break;
    case "GB":
      ByteValue = ByteValue / ByteUnits.GB;
      break;
    case "GiB":
      ByteValue = ByteValue / ByteUnits.GiB;
      break;
    case "TB":
      ByteValue = ByteValue / ByteUnits.TB;
      break;
    case "TiB":
      ByteValue = ByteValue / ByteUnits.TiB;
      break;
    Default:
      return false;
  }
  return ByteValue;
}

// 指定秒待機
function WaitSecond(WaitSec, TextMsg) {
  TextMsg = (typeof TextMsg !== "undefined") ? TextMsg : " 秒待機します";
  var wsStdOut = WScript.StdOut;
  var StartMsec = new Date();
  WaitMsec = parseInt(WaitSec) * 1000;
  do {
    WaitSec
        = parseInt((WaitMsec - (new Date() - StartMsec)) / 1000 + 0.9);
    if (TextMsg !== "") {
      DebugMsg("\r" + WaitSec + TextMsg, "StdOut", "Write");
    }
  } while (WaitSec);
}

// 引数を配列化
function GetArrayArgs(){
  var ScriptArgArray = [];
  for(var i = 0; i < WScript.Arguments.Count(); i++){
    ScriptArgArray.push(WScript.Arguments(i));
  }
  return ScriptArgArray;
}

// コマンドの引数を配列化
function SplitCommand(CommandName, CommandText, ArgLength) {
  ArgLength = (typeof ArgLength !== "undefined") ? ArgLength : "";
  var RegNotFound = -1;
  var RegEx = new RegExp("^" + CommandName + "\\(\\s*?(.*)\\s*?\\)$", "i");
  if (CommandText.search(RegEx) === RegNotFound) { return false; }
  var ArgText = CommandText.replace(RegEx, "$1");
  if (ArgLength === 0) { return ArgText; }
  var ArgArray = ArgText.split(",");
  var RegEx = new RegExp("((\"|\')?(\\w+)(\"|\')?)", "gi");
  for (var i=0; i < ArgArray.length; i++){
    ArgArray[i] = ArgArray[i].replace(RegEx, "$3").trim();
  }
  if (ArgLength !== "") { ArgArray.length = ArgLength; }
  return ArgArray;
}

// 文字列中のドライブを指定したドライブに変更
function SameDrv(TextStr, ReferenceDrv) {
  var sfo = new ActiveXObject("Scripting.FileSystemObject");
  ReferenceDrv = (typeof ReferenceDrv !== "undefined")
      ? ReferenceDrv
      : sfo.GetDriveName(WScript.ScriptFullName);
  var RegEx = /^[a-z]/i;
  if (RegEx.test(ReferenceDrv)) {
    if (ReferenceDrv.length === 1) {
      ReferenceDrv = ReferenceDrv + ":";
    }
    ReferenceDrv = sfo.GetDriveName(ReferenceDrv) + "\\";
    if (Array.isArray(TextStr)) {
      for (var i = 0; i < TextStr.length; i++) {
        TextStr[i] = TextStr[i].replace(/^\\/g, ReferenceDrv);
        TextStr[i] = TextStr[i].replace(/[a-z]:\\/gi, ReferenceDrv);
      }
    } else {
        TextStr = TextStr.replace(/^\\/g, ReferenceDrv);
        TextStr = TextStr.replace(/[a-z]:\\/gi, ReferenceDrv);
    }
  }
  return TextStr;
}

// 管理者権限情報を取得
function isAdmin() {
  var wsn = new ActiveXObject("WScript.Network");
  var ResultCode = 0;
  if (isAdminUser(wsn.UserName)) { ResultCode += 1; }
  if (isAdminProcess()) { ResultCode += 2; };
  return ResultCode;
}

// ログインユーザーの管理者権限情報を取得
function isAdminUser(UserName) {
  var ws = new ActiveXObject("WScript.Shell");
  var ResultString;
  var ResultArray = [];
  var wslocal = ws.Exec("net localgroup Administrators");
  var wsdomain = ws.Exec("net localgroup Administrators /domain");
  ResultString = wslocal.StdOut.ReadAll();
  ResultString += wsdomain.StdOut.ReadAll();
  ResultArray = ResultString.toString().split("\r\n");
  if(ResultArray.indexOf(UserName) === -1) {
    return false;
  }
  return true;
}

// 実行中プロセスの管理者権限取得
function isAdminProcess() {
  var ws = new ActiveXObject("WScript.Shell");
  var AppName = "openfiles";
  var ResultCode = false;
  var MODE = {
    WAIT: true,
    NOT_WAIT: false
  }
  var WINSTYLE = {
    HIDDEN: 0,
    NORMAL: 1,
    MINIMUM: 2,
    MAXIMUM: 3,
    INACTIVE_NORMAL: 4,
    LAST_TIME: 5,
    INACTIVE_MINIMUM: 7
  }
  if(ws.Run(AppName, WINSTYLE.HIDDEN, MODE.WAIT) === 0) {
    ResultCode = true;
  }
  return(ResultCode);
}

// 新規プロセスで指定ファイルを実行
function ExecuteCommand(ExeFName, AryArgs, Opt0, Opt1, Opt2) {
/*
** バグと思われますが、ShellExecute関数とPowershell start-processコマンド
** 共にRunAsによる管理者権限を付与して実行すると引数の扱いが通常と変わってしまいます。
** いくつかのパターンをマイクロソフトに報告済みです。
** ここではShellExecute関数からPowershellを管理者権限で実行し、Powershellに
** 希望するコマンドと引数を持たせてを実行し回避しています。
*/
  var sa = new ActiveXObject("Shell.Application");
  var sfo = new ActiveXObject("Scripting.FileSystemObject");
  var vOperation = {
    EDIT: "edit", // not use
    FIND: "find", // not use
    OPEN: "open", // Run or Open
    PRINT: "print", // not use
    PROPERTIES: "properties", // not use
    RUNAS: "runas" // Aministrator
  }
  var vShow = {
    HIDDEN: 0,
    NORMAL: 1,
    MINIMUM: 2,
    MAXIMUM: 3,
    INACTIVE_NORMAL: 4,
    LAST_TIME: 5,
    INACTIVE_MINIMUM: 7,
    APPLICATION: 10
  }
  var pwsWin = {
    NoWin: "-NoNewWindow",
    Normal: "-WindowStyle Normal",
    Hidden: "-WindowStyle Hidden",
    Minimized: "-WindowStyle Minimized",
    Maximized: "-WindowStyle Maximized"
  }
  if (!sfo.FileExists(ExeFName)) { return false; }
  var sFile = "powershell.exe";
  AryArgs = (typeof AryArgs !== "undefined") ? AryArgs : "";
  AryArgs = (AryArgs instanceof Array)
           ? AryArgs.join("\\\" \\\"") : AryArgs;
  AryArgs = "\\\"" + AryArgs + "\\\"";
  var vDirectory = sfo.GetParentFolderName(WScript.ScriptFullName);
  Opt0 = (typeof Opt0 !== "undefined") ? Opt0 : "";
  Opt1 = (typeof Opt1 !== "undefined") ? Opt1 : "";
  Opt2 = (typeof Opt2 !== "undefined") ? Opt2 : "";
  var Opt_Ary = [Opt0.trim(), Opt1.trim(), Opt2.trim()];
  var OptArgs = "";
  var ExeFlg = vOperation.OPEN;
  var WinFlg = pwsWin.Normal;
  for(var i = 0; i < Opt_Ary.length; i++) {
    switch (Opt_Ary[i].toString().toUpperCase()) {
      case "RUNAS":
        ExeFlg = vOperation.RUNAS;
        break;
      case "NOWIN":
      case "NOWINDOW":
      case "NONEWWINDOW":
        WinFlg = pwsWin.NoWin;
        break;
      case "NORMAL":
        WinFlg = pwsWin.Normal;
        break;
      case "HIDDEN":
        WinFlg = pwsWin.Hidden;
        break;
      case "MINIMIZED":
      case "MIN":
        WinFlg = pwsWin.Minimized;
        break;
      case "MAXIMIZED":
      case "MAX":
        WinFlg = pwsWin.Maximized;
        break;
      default:
        OptArgs = Opt_Ary[i].trim();
    }
  }
  var vArgs = "start-process "
            + "-FilePath '" + ExeFName + "' "
            + "-WorkingDirectory '" + vDirectory + "' "
            + WinFlg;
  OptArgs = ("\\\"" + OptArgs + "\\\" "
           + AryArgs).replace(/\\\"\\\"/gm, "").trim();
  if (OptArgs !== "") {
    vArgs = vArgs + " -ArgumentList '" + OptArgs + "'";
  }
  DebugMsg("Operation = " + ExeFlg);
  DebugMsg("Command : " + sFile + " " + vArgs);
  DebugMsg("");
  return sa.ShellExecute(sFile, vArgs, vDirectory, ExeFlg, vShow.HIDDEN);
}

// url接続先のhtmlからタイトルを取得
function GetHPTitle(url){
  var RegEx = /^https?:\/\/[\w!\?/\+\-_~=;\.,\*&@#\$%\(\)'\[\]]+$/g;
  if (url.search(RegEx) === 0) {
    var http = WScript.CreateObject("Msxml2.ServerXMLHTTP");
    http.open("GET", url, false);
    try {
      http.send();
    } catch (e) {
      return false;
    }
    var htmlfile = WScript.CreateObject("htmlfile");
    htmlfile.write("<meta http-equiv='x-ua-compatible' content='IE=11'>"
        + http.responseText);
    var gethp = htmlfile.parentWindow;
    var hptitle = gethp.document.querySelector("title");
    return hptitle.textContent;
  } else {
    return false;
  }
}

// デスクトップにショートカット作成
function MakeDesktopShortcut(ShortcutName, TargetFile) {
  var ws = new ActiveXObject("WScript.Shell");
  var sfo = new ActiveXObject("Scripting.FileSystemObject");
  var WINSTYLE = {
    HIDDEN: 0,
    NORMAL: 1,
    MINIMUM: 2,
    MAXIMUM: 3,
    INACTIVE_NORMAL: 4,
    LAST_TIME: 5,
    INACTIVE_MINIMUM: 7
  }
  var Shortcut_Full_Name = false;
  // ショートカット名がクオーテーションで囲まれていた場合、除去
  var RegEx = /^["']?(.*?)["']?$/g;
  ShortcutName = ShortcutName.replace(RegEx, "$1");
  // ＵＲＬ文字列が判定
  var RegEx = /^https?:\/\/[\w!\?/\+\-_~=;\.,\*&@#\$%\(\)'\[\]]+$/g;
  if (TargetFile.search(RegEx) === 0) {
    // ショートカット名が空だった場合、インターネットからタイトルを取得
    if (ShortcutName === "") {
      ShortcutName = GetHPTitle(TargetFile);
    }
    // ＵＲＬとしてショートカット名を設定
    Shortcut_Full_Name = ws.SpecialFolders("Desktop") + "\\"
        + ShortcutName + ".url";
  // ＵＲＬではない場合、ファイルまたはフォルダの存在をチェックしターゲット名を取得
  } else if (sfo.FileExists(TargetFile) || sfo.FolderExists(TargetFile)) {
    if (ShortcutName === "") {
      ShortcutName = sfo.GetBaseName(TargetFile);
    }
    // ファイルまたはフォルダとしてショートカット名を設定
    Shortcut_Full_Name = ws.SpecialFolders("Desktop") + "\\"
        + ShortcutName + ".lnk";
  } else {
    DebugMsg("Shortcut Target Not Found.");
  }
  // すでに同じショートカット名のファイルがある場合、削除
  if (sfo.FileExists(Shortcut_Full_Name)) {
    sfo.DeleteFile(Shortcut_Full_Name, true);
  }
  // ショートカット作成
  if (Shortcut_Full_Name !== false) {
    DebugMsg("Shortcut Created.");
    var wscs = ws.CreateShortcut(Shortcut_Full_Name);
    wscs.TargetPath = TargetFile;
    if (sfo.GetExtensionName(Shortcut_Full_Name) === "lnk") {
      wscs.WindowStyle = WINSTYLE.MAXIMUM;
      wscs.IconLocation = TargetFile + ", 0";
      wscs.WorkingDirectory = ws.SpecialFolders("Desktop");
      wscs.Description = "Created by "
          + sfo.GetBaseName(WScript.ScriptFullName);
    }
    try {
      wscs.Save();
    } catch(e) {
      DebugMsg("Error(" + e + ")");
    }
  }
  DebugMsg("  Shortcut File Name : " + Shortcut_Full_Name);
  DebugMsg("  Target File Name   : " + TargetFile);
  DebugMsg("");
}

// 相対パスでシンボリックリンク作成
function RelativeMKLink(LinkFile, TargetFile) {
  var ws = new ActiveXObject("WScript.Shell");
  var sfo = new ActiveXObject("Scripting.FileSystemObject");
  var RunningDrv = sfo.GetDriveName(WScript.ScriptFullName) + "\\";
  var MKLink_Cmd = "cmd /c mklink";
  var ErrFlg = 0;
  var MODE = {
    WAIT: true,
    NOT_WAIT: false
  }
  var WINSTYLE = {
    HIDDEN: 0,
    NORMAL: 1,
    MINIMUM: 2,
    MAXIMUM: 3,
    INACTIVE_NORMAL: 4,
    LAST_TIME: 5,
    INACTIVE_MINIMUM: 7
  }
  // ドライブ名を環境に合わせて変更
  LinkFile = LinkFile.replace(/^\\/, "");
  LinkFile = LinkFile.replace(/^[a-z]:\\/gi, "");
  TargetFile = TargetFile.replace(/^\\/, "");
  TargetFile = TargetFile.replace(/^[a-z]:\\/gi, "");
  DebugMsg("\nRelative MKLink");
  DebugMsg("  Link   File : " + RunningDrv + LinkFile);
  DebugMsg("  Target File : " + RunningDrv + TargetFile);
  // 実行中のプロセスが管理者権限を持っているか確認
  if(!isAdminProcess()) {
    DebugMsg("\nError : Process Not Administrator.");
    ErrFlg += 1;
  }
  // リンクファイルが存在するか確認
  if (sfo.FileExists(RunningDrv + LinkFile)) {
    DebugMsg("\nError : Exist Link File.");
    ErrFlg += 2;
  }
  // リンクファイルのフォルダが存在するか確認
  if (!sfo.FolderExists(sfo.GetParentFolderName(RunningDrv + LinkFile))) {
    DebugMsg("\nError : Link Folder Not Found."
        + sfo.GetParentFolderName(RunningDrv + LinkFile));
    ErrFlg += 4;
  }
  // ターゲットファイルが存在するか確認
  if (!sfo.FileExists(RunningDrv + TargetFile)) {
    DebugMsg("\nError : Target File Not Found.");
    ErrFlg += 8;
  }
  // エラーがなければリンク作成
  if(ErrFlg === 0) {
    // リンクファイルとターゲットファイルが相対位置になるようにフォルダの指定を整形
    MKLink_Cmd = MKLink_Cmd + " \""
        + RunningDrv + LinkFile + "\" \""
        + Array(LinkFile.split("\\").length).join("..\\")
        + TargetFile + "\"";
    DebugMsg("  Command     : " + MKLink_Cmd);
    ws.Run(MKLink_Cmd, WINSTYLE.HIDDEN, MODE.NOT_WAIT);
  }
  DebugMsg("");
}

// 仮想ファイルをマウント、または最適化（コンパクト）します
  var ws = new ActiveXObject("WScript.Shell");
  var sfo = new ActiveXObject("Scripting.FileSystemObject");
  var VDiskCMD = {
    CMD: "cmd /c (",
    SelectFile: "echo select vdisk file=",
    SelectVolume: "echo select volume=",
    Attach: "&&echo attach vdisk",
    AttachReadOnly: "&&echo attach vdisk readonly",
    Detach: "&&echo detach vdisk",
    Assign: "&&echo assign letter=",
    VolumeList: "&&echo list volume",
    DetailDisk: "&&echo detail disk",
    DetailVDISK: "&&echo detail vdisk",
    Compact: "&&echo compact vdisk",
    Exit: "&&echo exit)|Diskpart.exe"
  }
  var VDError = {
    NotAdministrator: -1,
    FileNotFound: -2,
    UseDriveLetter: -8
  }
  var NotFound = -1;
  var ResultStr;
  var RegEx;
  var ResultRegEx;
  // プロセスが管理者権限を持っているか確認
  if (isAdminProcess() !== true) {
    DebugMsg("\nError : Process Not Administrator.");
    return VDError.NotAdministrator;
  }
  // オプションの確認、エラーの場合でもスルーしています。
  VDOpt = (typeof VDOpt !== "undefined")
         ? VDOpt.toString().toUpperCase() : "";
  if (VDOpt !== "COMPACT") {
    if (VDOpt.length === 1) { VDOpt = VDOpt + ":"; }
    VDOpt = sfo.GetDriveName(VDOpt);
  }
  // 仮想ファイルが絶対PATHではない場合、カレントフォルダに設定
  if (sfo.GetParentFolderName(VDiskName) === "") {
    VDiskName = sfo.GetParentFolderName(WScript.ScriptFullName)
        + "\\" + VDiskName;
  }
  // 仮想ファイルを確認
  if (!sfo.FileExists(VDiskName)) {
    DebugMsg("\nError : VDISK Not Found. [" + VDiskName + "]\n");
    return VDError.FileNotFound;
  }
  // DiskPartのコマンド作成
  var VolumeCheck
      = VDiskCMD.CMD
      + VDiskCMD.SelectFile + "\"" + VDiskName + "\""
      + VDiskCMD.Detach
      + VDiskCMD.VolumeList
      + VDiskCMD.Exit;
  var AttachVDisk
      = VDiskCMD.CMD
      + VDiskCMD.SelectFile + "\"" + VDiskName + "\""
      + VDiskCMD.Detach
      + VDiskCMD.Attach
      + VDiskCMD.VolumeList
      + VDiskCMD.DetailDisk
      + VDiskCMD.Exit;
  var AttachROVDisk
      = VDiskCMD.CMD
      + VDiskCMD.SelectFile + "\"" + VDiskName + "\""
      + VDiskCMD.Detach
      + VDiskCMD.AttachReadOnly
      + VDiskCMD.VolumeList
      + VDiskCMD.DetailDisk
      + VDiskCMD.Exit;
  var DetachVDisk
      = VDiskCMD.CMD
      + VDiskCMD.SelectFile + "\"" + VDiskName + "\""
      + VDiskCMD.Detach
      + VDiskCMD.Exit;
  // 使用中のボリューム番号、ドライブレターを確認
  var VolumeNums = [];
  var ConnectedDrvs = [];
  RegEx = new RegExp("\\s+Volume\\s+(\\d+)\\s+([A-Z])?\\s+", "gi");
  ResultStr = ws.Exec(VolumeCheck).StdOut.ReadAll();
  while ((ResultRegEx = RegEx.exec(ResultStr)) !== null) {
    VolumeNums.push(ResultRegEx[1]);
    ConnectedDrvs.push(ResultRegEx[2]);
  }
  // 指定したドライブレターがすでに使用されているか確認
  if (ConnectedDrvs.indexOf(VDOpt) > NotFound) {
    DebugMsg("The specified drive \"" + VDOpt + "\" is in use.");
    VDError.UseDriveLetter;
  }
  // 仮想ファイルをマウントしボリューム番号を取得
  if (VDOpt !== "COMPACT") {
    ResultStr = ws.Exec(AttachVDisk).StdOut.ReadAll();
  } else {
    ResultStr = ws.Exec(AttachROVDisk).StdOut.ReadAll();
  }
  var i = 0;
  while ((ResultRegEx = RegEx.exec(ResultStr)) !== null) {
    if (VolumeNums[i] !== ResultRegEx[1]) { break; }
    i++;
  }
  // DiskPartのコマンド作成
  var CompactVDisk
      = VDiskCMD.CMD
      + VDiskCMD.SelectFile + "\"" + VDiskName + "\""
      + VDiskCMD.SelectVolume + ResultRegEx[1]
      + VDiskCMD.CompactVDisk
      + VDiskCMD.DetailDisk
      + VDiskCMD.Detach
      + VDiskCMD.Exit;
  var AssignVDisk
      = VDiskCMD.CMD
      + VDiskCMD.SelectVolume + ResultRegEx[1]
      + VDiskCMD.Assign + VDOpt
      + VDiskCMD.DetailDisk
      + VDiskCMD.Exit;
  // ドライブレターが指定されていた場合、ドライブレターを変更
  if (VDOpt !== "") {
    if (VDOpt === "COMPACT") {
      ResultStr = ws.Exec(CompactVDisk).StdOut.ReadAll();
      VDOpt = ResultRegEx[2];
    } else {
      ResultStr = ws.Exec(AssignVDisk).StdOut.ReadAll();
    }
  } else {
    VDOpt = ResultRegEx[2];
  }
  var FileInfo = sfo.GetFile(VDiskName);
  DebugMsg("Attach VDisk.");
  DebugMsg("VDisk File Name             : " + VDiskName);
  DebugMsg(" File Type                  : " + FileInfo.Type);
  DebugMsg(" File Size                  : "
      + NumberFormat(ConvByteStr(FileInfo.Size, "GiB"), 9, 3)
      + " GiB ("
      + NumberFormat(FileInfo.Size, 16) + " Byte)");
  DebugMsg("");
  var DriveInfo = sfo.GetDrive(VDOpt);
  DebugMsg("Mounted on Drive            : " + DriveInfo.DriveLetter);
  DebugMsg(" File System                : " + DriveInfo.FileSystem);
  DebugMsg(" Drive Total Size           : "
      + NumberFormat(ConvByteStr(DriveInfo.TotalSize, "GiB"), 9, 3)
      + " GiB ("
      + NumberFormat(DriveInfo.TotalSize, 16) + " Byte)");
  DebugMsg(" Drive Free Space           : "
      + NumberFormat(ConvByteStr(DriveInfo.FreeSpace, "GiB"), 9, 3)
      + " GiB ("
      + NumberFormat(DriveInfo.TotalSize, 16) + " Byte)");
  DebugMsg("");
  return VDOpt;
}

// 仮想ファイルをアンマウント
function VDiskDetach(VDiskName) {
  var ws = new ActiveXObject("WScript.Shell");
  var sfo = new ActiveXObject("Scripting.FileSystemObject");
  var ResultStr;
  var RegEx;
  var ResultRegEx;
  var VolumeNo;
  var VDiskCMD = {
    CMD: "cmd /c (",
    SelectFile: "echo select vdisk file=",
    SelectVolume: "echo select volume=",
    Attach: "&&echo attach vdisk",
    AttachReadOnly: "&&echo attach vdisk readonly",
    Detach: "&&echo detach vdisk",
    Assign: "&&echo assign letter=",
    VolumeList: "&&echo list volume",
    DetailDisk: "&&echo detail disk",
    DetailVDISK: "&&echo detail vdisk",
    Compact: "&&echo compact vdisk",
    Exit: "&&echo exit)|Diskpart.exe"
  }
  var VDError = {
    NotAdministrator: -1,
    FileNotFound: -2,
    OptStrError: -4,
    UseDriveLetter: -8
  }
  // プロセスが管理者権限を持っているか確認
  if (isAdminProcess() !== true) {
    return VDError.NotAdministrator;
  }
  // 仮想ファイルを確認
  if (sfo.GetParentFolderName(VDiskName) === "") {
    VDiskName = sfo.GetParentFolderName(WScript.ScriptFullName)
        + "\\" + VDiskName;
  }
  if (!sfo.FileExists(VDiskName)) {
    DebugMsg("\nError : VDisk Not Found. [" + VDiskName + "]\n");
    return VDError.FileNotFound;
  }
  // DiskPartのコマンド作成
  var DetachVDisk
      = VDiskCMD.CMD
      + VDiskCMD.SelectFile + "\"" + VDiskName + "\""
      + VDiskCMD.Detach
      + VDiskCMD.Exit;
  ResultStr = ws.Exec(DetachVDisk).StdOut.ReadAll();
  DebugMsg("Detach VDisk.");
  DebugMsg("VDisk File Name             : " + VDiskName);
  DebugMsg("");
}

// 設定情報のファイル作成
function CreateConfigFile(ConfigFileName) {
  var ConfigText = function() {/*
#
# ◆◇◆ 携帯環境作成支援スクリプト ◆◇◆
# -Mobile Environment Creation Support Script-
# Copyright(c) 2021 Aromatibus (https://github.com/Aromatibus)
#
# USBメモリや仮想ディスクなど携帯可能な環境でポータブル版のソフトウェアを活かすための
# 環境作成をお手伝いします。
#
# ◆できること
# 　・スクリプトが実行されたドライブまたは指定したドライブを基準にした環境設定ができます。
# 　・設定した環境変数はシャットダウン後に設定前の状態に戻ります。
# 　・デスクトップにショートカットを作成できます。
# 　・相対パス指定のシンボリックリンクを簡単に作成できます。
# 　・仮想ディスクファイル（VHD, VHDX）を指定ドライブにマウント、アンマウントができます。
# 　・プログラムの実行ができます。
# 　・引数またはファイルをドラックアンドドロップして実行すると引数を実行
# 　　または引数を指定してスクリプトを実行できます。
# 　※設定した環境変数は同プロセス、および以前に実行されたプロセスには影響しません。
# 　※設定した環境変数はVolatile環境変数に記録されるため、シャットダウン後に設定前の状態に
# 　　戻る（削除される）はずですが、戻らない場合が確認されています。
# 　　その場合は本スクリプトで削除することができます。
#
# ◆スクリプトのファイル名は変更できます。
# 　・使用する環境やソフトウェアに合わせた名前で使用できます。
# 　・設定ファイルは"スクリプト名.txt"が読み込まれます。
# 　・拡張子はバッチファイルの"CMD"と"BAT"、JScriptの"JS"が使用できます。
# 　　バッチファイルの拡張子はWindows2000以降では両方使用できます。
# 　※２種類のバッチファイルにはエラーレベルの扱いに違いがあります。
# 　・拡張子を"JS"にする場合は、実行中のメッセージが表示できないためスクリプトの
# 　　"var Debug_Flg = true;"の"true"を"false"に変更しないとエラーになります。
#
# ◆設定ファイルについて
# 　スクリプト中の"var OneFileMode = false;"の"false"を"true"にすると
# 　設定ファイルを外部に作らずスクリプト内の設定情報を読み込みます。
#
# ◆環境変数"PATHEXT"について
# 　コマンドプロンプトでは同名のファイルがあった場合、実行される拡張子の優先順位は一般的に、
# 　"EXE" > "BAT" > "CMD" > "VBS" > "JS"になります。
# 　※詳細は環境変数"PATHEXT"を確認してください。
# 　・サンプルの設定ファイルでは環境変数"PATHEXT"を変更し、拡張子"CMD"が最優先に実行
# 　　されるように変更しています。
# 　・また、拡張子"PY"を加えて実行ファイルとして扱えるようにしています。
# 　・"CMD"を本スクリプトのようにJScriptタイプのバッチファイルと決めておくなどすると
# 　　区別しやすくなります。
#
# ◆設定ファイルの書き方は本設定ファイルを参考にしてください。
# 　"#"から始まる行はコメント行となります。
#
# ◆"RunAdministrator=true"の行があった場合
# 　実行したプロセスが管理者権限を持っていない場合、管理者権限で再実行されます。
# 　デフォルトの設定はオフ（false）です。
# 　管理者権限が必要なコマンドやプログラムを実行する場合に指定してください。
# 　※このコマンドは最優先で実行されます。
#
# ◆"ReferenceDrive=ドライブ名"の行があった場合
# 　基準となる参照先ドライブを変更します。
# 　デフォルトはスクリプトを実行したドライブです。
#
# ◆環境変数"PATH"の設定（フォルダを設定した場合）
# 　ドライブ名を変更し環境変数"PATH"に追記されます。
# 　"Execute"および"ExecuteArg"を実行しない限り環境変数"PATH"は最後に設定されます。
# 　ドライブ名はデフォルトではスクリプトを実行したドライブ名に書き換えられます。
# 　たとえばスクリプトが実行されたドライブがＤドライブだった場合は、
# 　a:\abc -> D:\abc
# 　\abc   -> D:\abc
# 　上記のようにドライブ名が置換または行の先頭が"\"で始まっていた場合は補完されます。
# 　環境変数"PATH"に同じフォルダがある、またはフォルダが存在しない場合は無視されます。
# 　・スクリプトで設定したフォルダの参照優先度は低く最後に参照されます。
# 　　同名のファイルが優先度の高い別のフォルダにある場合、そちらが優先されます。
# 　・アプリ実行エイリアスが有効の場合
# 　　Python.exeなど一部のソフトはアプリ実行エイリアスで設定された"PATH"が優先される
# 　　ため参照順があとになり指定した環境で実行できない場合があります。
# 　　事前に相対PATHで別名のシンボリックリンクを作るなど対策が必要です。
# 　※最初に作成される設定ファイルでは"PYT"の名前でシンボリックリンクを作成しています。
# 　　"Python.exe"の代わりに"PYT"を実行することで回避できます。
#
# ◆":=:"を含む行があった場合
# 　":=:"を境に左辺を環境変数名、右辺を変数として環境変数に設定します。
# 　環境変数名がすでに使用されている場合は上書きされますがシャットダウン後に戻ります。
# 　右辺を指定しない場合、環境変数の削除をします。
# 　※"PATH"の設定と同じようにドライブ名は書き換えられます。
# 　※ドライブ名の書き換えは変数のみです。
#
# ◆"RemoveVolatile=true"の行があった場合
# 　レジストリからVolatile環境変数をすべて削除します。
# 　再起動後、Volataile環境変数はリセットされる仕様のはずですが、リセットされない場合が
# 　あるために用意したコマンドです。
# 　通常は":=:"を使用して削除してください。
# 　※再起動するとレジストリのVolatile環境変数はデフォルトの値に戻ります。
# 　　コマンドの実行結果は自己責任になります。
#
# ◆"DesktopShortcut(文字列A, 文字列B)"の行があった場合
# 　ユーザーのデスクトップにショートカットを作成します。
# 　"PATH"の設定と同じようにドライブ名は書き換えられます。
# 　文字列Aはショートカット名になり、同名のショートカットがあった場合は上書きされます。
# 　ショートカット名を""囲みで省略すると、ファイル名、フォルダ名、ＵＲＬのタイトルになります。
# 　文字列Bはリンク元のファイル名、フォルダ名または"http"から始まるＵＲＬを指定します。
# 　※"PATH"の設定と同じようにドライブ名は書き換えられます。
#
# ◆"RelativeMKLink(リンクファイル名, リンク元ファイルの絶対パス)"の行があった場合
# 　リンクファイル名で指定した名前でシンボリックリンクを作成します。
# 　リンク元ファイルの絶対パスのファイルの位置をリンクファイル名との相対PATHでリンクが
# 　作成されるため、ドライブ名が変わった場合でもリンクは正常に機能します。
# 　※"PATH"の設定と同じようにドライブ名は書き換えられます。
# 　※管理者権限を持ったプロセスで実行する必要があります。
#
# ◆"VDiskAttch(仮想ディスクファイル名, 接続するドライブ名)"の行があった場合
# 　仮想ディスクファイル名を接続するドライブ名のドライブにマウントします。
# 　仮想ディスクファイル名にパスを省略するとスクリプトと同じフォルダから読み込みます。
# 　接続するドライブ名に"Compact"を指定すると仮想ディスクを最適化します。
# 　※Diskpartコマンドで作成した仮想ディスクファイル、VHD, VHDXで動作検証しています。
# 　※"PATH"の設定と同じようにドライブ名は書き換えられます。
# 　※管理者権限を持ったプロセスで実行する必要があります。
#
# ◆"VDiskDetch(仮想ディスクファイル)"の行があった場合
# 　仮想ディスクファイルをアンマウントします。
# 　仮想ディスクファイル名にパスを省略するとスクリプトと同じフォルダから読み込みます。
# 　※"PATH"の設定と同じようにドライブ名は書き換えられます。
# 　※管理者権限を持ったプロセスで実行する必要があります。
#
# ◆"WaitSec(数字)"の行があった場合
# 　数字の秒数のあいだ処理を停止し待機します。
#
# ◆"Execute(文字列)"の行があった場合
# 　直前までに設定された環境変数を引継ぎ、文字列を実行します。
# 　以降、設定された環境変数は影響しません。
# 　アプリケーション以外は関連付けされた拡張子に合わせて文字列を実行します。
# 　※"PATH"の設定と同じようにドライブ名は書き換えられます。
#
# ◆"ExecuteArg(文字列)"の行があった場合
# 　直前までに設定された環境変数を引継ぎ、引数付きで文字列を実行します。
# 　以降、設定された環境変数は影響しません。
# 　アプリケーション以外は関連付けされた拡張子に合わせて文字列を実行します。
# 　引数を指定されていない場合は引数無しで実行されます。
# 　※"PATH"の設定と同じようにドライブ名は書き換えられます。
#
# ◆引数、またはファイルをドラッグアンドドロップして実行された場合
# 　ドライブ名は書き換えません。
# 　ExecuteArg(文字列)の引数になります。
# 　ExecuteArg(文字列)が使用されていない場合、第一引数の文字列を処理の最後に実行します。
# 　第二引数以降は第一引数の引数になります。
#
# ◆ライセンス
# 　Copyright(c) 2021 Aromatibus (https://github.com/Aromatibus)
# 　Released under the MIT license.
# 　See https://opensource.org/licenses/MIT
#

# 管理者権限で実行
RunAdministrator=true

# アプリ実行エイリアス回避用、Python.exeのエイリアス作成
#RelativeMKLink(a:\Python\PYT, a:\Python\Python.exe)

# 環境変数 "PATH" へフォルダを追加
\VSCode
\VSCode\MinGW\bin
\VSCode\node.js
\VSCode\Git\bin
\VSCode\Git\cmd
\Python
\Python\Scripts

# 環境変数を追加
NODE_PATH:=:a:\VSCode\node.js\node_modules\npm\node_modules;a:\VSCode\node.js\node_modules\npm
GIT.PATH:=:a:\VSCode\git\bin\git.exe

# 実行できる拡張子と優先度を変更 > 優先(CMD)、追加(PY)
PATHEXT:=:.CMD;.COM;.EXE;.BAT;.VBS;.VBE;.JS;.JSE;.WSF;.WSH;.MSC;.PY

# デスクトップへショートカットを作成
DesktopShortcut(VSCode Portable, a:\VSCode\Code.exe)
DesktopShortcut("", a:\Source)
DesktopShortcut("", https://github.com/Aromatibus)

# 仮想ディスクを接続し仮想ディスクからファイルを起動
ReferenceDrive = c
vdiskattach(a:\Users\Aromatibus\Documents\test.vhd, z)
ReferenceDrive = z
execute(a:\TestEnv.cmd)

# 参照先をデフォルトに戻す
ReferenceDrive =

# NexusFontを起動、引数なし
Execute(a:\Fonts\NexusFont\NexusFont.exe)

# NexusFontの起動待ち
WaitSec(5)

# VSCodeを起動、引数あり
ExecuteArg(a:\VSCode\Code.exe)
*/}.toString().split("\r\n").slice(1,-1).join("\r\n");
  var ws = new ActiveXObject("WScript.Shell");
  var sfo = new ActiveXObject("Scripting.FileSystemObject");
  var MODE = {
    READING: 1,
    WRITING: 2,
    APPENDING: 8,
    CREATE: true,
    NOT_CREATE: false,
    WAIT: true,
    NOT_WAIT: false
  }
  var WINSTYLE = {
    HIDDEN: 0,
    NORMAL: 1,
    MINIMUM: 2,
    MAXIMUM: 3,
    INACTIVE_NORMAL: 4,
    LAST_TIME: 5,
    INACTIVE_MINIMUM: 7
  }
  if (ConfigFileName === "") { return ConfigText; }
  var otf = sfo.OpenTextFile(ConfigFileName, MODE.WRITING, MODE.CREATE);
  otf.Write(ConfigText);
  otf.Close();
  DebugMsg("Created a configuration file. : " + ConfigFileName);
  ws.Run(ConfigFileName, WINSTYLE.MAXIMUM, MODE.NOT_WAIT);
}

WScript.Quit(main());
