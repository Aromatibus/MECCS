@if(0)==(0) rem /* Start of Batch
@echo off
setlocal
call :SetESC
title �g�ъ��쐬�x���X�N���v�g
echo.
echo %ESC%[102m�g�ъ��쐬�x���X�N���v�g%ESC%[0m
echo.
echo USB�������Ȃǌg�ъ��Ń|�[�^�u���\�t�g�����������߂̊��쐬������`�����܂��B
echo Copyright(c) 2021 Aromatibus (https://github.com/Aromatibus)
echo.
:: Switch to JScript mode
cscript.exe //nologo //E:JScript "%~f0" %*
echo.
if %errorlevel% equ 0 (
  echo %ESC%[41;97m�X�N���v�g�ɃG���[���������܂����B%ESC%[0m
  echo.
  pause
  goto :eof
)
if %errorlevel% equ 1 (
  echo %ESC%[102m�������I�����܂����B%ESC%[0m
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

// ���s�󋵂��R�}���h�v�����v�g�̃E�C���h�E�ɕ\������t���O
var DebugFlg = true;
// �O���ɐݒ�t�@�C������炸�ɓ����̏��œ��삳����t���O
var OneFileMode = false;

// Polyfill trim
if (!String.prototype.trim) {
  String.prototype.trim = function () {
    return this.replace(/^[\s�@\uFEFF\xA0]+|[\s�@\uFEFF\xA0]+$/g, '');
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
  // �I�u�W�F�N�g�ϐ�
  var ws = new ActiveXObject("WScript.Shell");
  var wseProcess = ws.Environment("Process");
  var wseSystem = ws.Environment("System");
  var wseUser = ws.Environment("User");
  var wseVolatile = ws.Environment("Volatile");
  var wsn = new ActiveXObject("WScript.Network");
  var sfo = new ActiveXObject("Scripting.FileSystemObject");
  // ���Q�ƕϐ�
  var RunningDrv = sfo.GetDriveName(WScript.ScriptFullName);
  var ReferenceDrv = RunningDrv;
  var RunningFolder = sfo.GetParentFolderName(WScript.ScriptFullName);
  var RunningFileName = sfo.GetBaseName(WScript.ScriptFullName);
  var ConfigFileName = RunningFileName + ".txt";
  var ScriptArgArray = GetArrayArgs();
  var AdminFlg = isAdmin();
  // �萔
  var RegFound = 0; // ��������� 0 �ȏ�ɂȂ�܂��B
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
  // �ϐ�
  var ReadData;
  var CommandName;
  var ArgArray = [];
  var ArgLength;
  var RegEx;
  var TmpPATH = "";
  var ExeArgFlg = false;
  // �J�����g�t�H���_���X�N���v�g�̃t�H���_�ɃZ�b�g
  ws.CurrentDirectory = RunningFolder;
  // �ݒ�t�@�C���̍쐬�E�ǂݍ���
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
  // �ݒ�t�@�C���𐮌`
  // �t�@�C��������CRLF��t�����Ă���
  ConfigText = ConfigText + "\r\n";
  // CRLF�ɓ���
  ConfigText = ConfigText.replace(/(\r|\n|\r\n)/g, "\r\n");
  // �s���A�s���̋󔒏���(TRIM)
  ConfigText = ConfigText.replace(/(\r\n)[\s�@\uFEFF\xA0]+/g, "\r\n");
  ConfigText = ConfigText.replace(/[\s�@\uFEFF\xA0]+(\r\n)/g, "\r\n");
  // �R�����g�s�폜
  ConfigText = ConfigText.replace(/^#.*$(\r\n)/gm, "");
  // ��s�폜
  ConfigText = ConfigText.replace(/^$(\r\n)/g, "");
  ConfigText = ConfigText.replace(/(\r\n)^$/m, "");
  // �D�悷��R�}���h�����s
  RegEx = /^RunAdministrator\s*=\s*true$/gmi;
  if (ConfigText.search(RegEx) !== RegNotFound) {
    if (!(AdminFlg & 2)) {
      DebugMsg("Re-execute with administrative privileges.\n");
      ExecuteCommand(WScript.ScriptFullName, ScriptArgArray, "RunAs");
      return ErrorlevelRebootAdmin;
    }
  }
  // �ݒ�����s���Ƃɔz��ɕ�������
  ConfigText = ConfigText.split("\r\n");
  // �v���Z�X���ϐ���"PATH"�̖�����";"���Ȃ��ꍇ�A";"���t����
  if (wseProcess.item("PATH").slice(-1) !== "\;") { TmpPATH = "\;"; }
  // ���[�U�[�����A���s�v���Z�X�����̏��\��
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
  // �������\��
  if (ScriptArgArray.length === 0) {
    DebugMsg("Arguments nothing.\n");
  } else {
    DebugMsg("Arguments = [" + ScriptArgArray + "]\n");
  }
  // ���ϐ����\��
  DebugMsg("Reference Drive = \"" + ReferenceDrv + "\"\n");
  DebugMsg("Process PATH =\n[" + wseProcess.item("PATH") + "]\n");
  DebugMsg("System PATH =\n[" + wseSystem.item("PATH") + "]\n");
  DebugMsg("User PATH =\n[" + wseUser.item("PATH") + "]\n");
  DebugMsg("Volatile PATH =\n[" + wseVolatile.item("PATH") + "]\n");
  // �ݒ����z�񂲂Ƃɏ���
  for (var i=0; i < ConfigText.length; i++){
    ReadData = ConfigText[i];
//    DebugMsg("Command [" + NumberFormat(i, 2) + "] : " + ReadData);
    // ���ϐ�"PATH"�֕ϐ���ǉ����邽�߂̎��O����
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
    // Volatile���ϐ��̍폜
    RegEx = /^RemoveVolatile\s*=\s*true$/i;
    if (ReadData.search(RegEx) === RegFound) {
      ws.Exec("cmd /c reg delete "
          + "\"HKEY_CURRENT_USER\\Volatile Environment\"");
      DebugMsg("Registry \"HKEY_CURRENT_USER\\Volatile Environment\""
          + " removed.");
    }
    // ���L��̊��ϐ��̍폜
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
    // ���L��̊��ϐ��̒ǉ�
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
    // �Q�Ƃ���h���C�u��ύX
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
    // �w��b���̑ҋ@����
    RegEx = /^WaitSec\((\d+)\)$/gi;
    if (ReadData.search(RegEx) === RegFound) {
      DebugMsg("");
      WaitSecond(parseInt(ReadData.replace(RegEx, "$1")))
    }
    // �f�X�N�g�b�v�փV���[�g�J�b�g���쐬
    CommandName = "DesktopShortcut";
    ArgLength = 2;
    if (SplitCommand(CommandName, ReadData, ArgLength) !== false) {
      ArgArray = SplitCommand(CommandName, ReadData, ArgLength);
      ArgArray = SameDrv(ArgArray, ReferenceDrv);
      MakeDesktopShortcut(ArgArray[0], ArgArray[1]);
    }
    // �V���{���b�N�����N���쐬 ���v���Z�X�̊Ǘ��Ҍ������K�v
    CommandName = "RelativeMKLink";
    ArgLength = 2;
    if (SplitCommand(CommandName, ReadData, ArgLength) !== false) {
      ArgArray = SplitCommand(CommandName, ReadData, ArgLength);
      ArgArray = SameDrv(ArgArray, ReferenceDrv);
      RelativeMKLink(ArgArray[0], ArgArray[1]);
    }
    // ���z�f�B�X�N�̃}�E���g���� ���v���Z�X�̊Ǘ��Ҍ������K�v
    CommandName = "VDiskAttach";
    ArgLength = 2;
    if (SplitCommand(CommandName, ReadData, ArgLength) !== false) {
      ArgArray = SplitCommand(CommandName, ReadData, ArgLength);
      VDiskAttach(SameDrv(ArgArray[0], ReferenceDrv), ArgArray[1]);
    }
    // ���z�f�B�X�N�̃A���}�E���g���� ���v�v���Z�X�̊Ǘ��Ҍ���
    CommandName = "VDiskDetach";
    ArgLength = 1;
    if (SplitCommand(CommandName, ReadData, ArgLength) !== false) {
      ArgArray = SplitCommand(CommandName, ReadData, ArgLength);
      ArgArray = SameDrv(ArgArray[0], ReferenceDrv);
      VDiskDetach(ArgArray[0]);
    }
    // �w�蕶��������s
    CommandName = "Execute";
    ArgLength = 1;
    if (SplitCommand(CommandName, ReadData, ArgLength) !== false) {
      ArgArray = SplitCommand(CommandName, ReadData, ArgLength);
      ArgArray = SameDrv(ArgArray, ReferenceDrv);
      DebugMsg("\n" + CommandName + " : [" + ArgArray + "]");
      // ���ϐ�"PATH"�֊��������Œǉ�
      if (TmpPATH !== "") {
        wseVolatile.item("PATH") = wseVolatile.item("PATH") + TmpPATH;
        wseProcess.item("PATH") = wseProcess.item("PATH") + TmpPATH;
        DebugMsg("\nAttach Volatile PATH = \n["
            + wseVolatile.item("PATH") + "]\n");
        TmpPATH = "";
      }
      // �t�@�C�������邩�m�F��A���s
      if (sfo.FileExists(ArgArray)) {
        ExecuteCommand(ArgArray);
      } else {
        DebugMsg("Execute : Error Execute File Not Found.\n");
      }
    }
    // �w�蕶����������t���Ŏ��s
    CommandName = "ExecuteArg";
    ArgLength = 1;
    if (SplitCommand(CommandName, ReadData, ArgLength) !== false) {
      ArgArray = SplitCommand(CommandName, ReadData, ArgLength);
      ArgArray = SameDrv(ArgArray, ReferenceDrv);
      ExeArgFlg = true;
      DebugMsg("\n" + CommandName + " : [" + ArgArray + "]");
      // ���ϐ�PATH�֊��������Œǉ�
      if (TmpPATH !== "") {
        wseVolatile.item("PATH") = wseVolatile.item("PATH") + TmpPATH;
        wseProcess.item("PATH") = wseProcess.item("PATH") + TmpPATH;
        DebugMsg("\nAttach Volatile PATH = \n["
            + wseVolatile.item("PATH") + "]\n");
        TmpPATH = "";
      }
      // �t�@�C�������邩�m�F��A���s
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
  // ���ϐ�PATH�֊��������Œǉ�
  if (TmpPATH !== "") {
    wseVolatile.item("PATH") = wseVolatile.item("PATH") + ";" + TmpPATH;
    wseProcess.item("PATH") = wseProcess.item("PATH") + ";" + TmpPATH;
    DebugMsg("\nAttach Volatile PATH = \n["
        + wseVolatile.item("PATH") + "]\n");
  }
  // ����������A�ݒ���� "ExecuteArg "���Ȃ��ꍇ�́A���������s���܂�
  if ((ScriptArgArray.length > 0) && (!ExeArgFlg)) {
    if (ScriptArgArray.length > 1) {
      ExecuteCommand(ScriptArgArray[0], ScriptArgArray.slice(1));
    } else {
      ExecuteCommand(ScriptArgArray);
    }
  }
  return ErrorlevelNormal;
}

// �f�o�b�O�������̕\��
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

// �����͂R����؂�w�菬�������͐؎̂Ďw�蕶�����ɑ����܂��B���l�ȊO�͎w�蕶�����ɑ����܂��B
function NumberFormat(FloatNum, Digits, DecPoint) {
  DecPoint =
      (typeof DecPoint !== "undefined") ? DecPoint : 0;
  if (DecPoint > 0) { Digits = Digits - DecPoint - 1; }
      var RegEx = new RegExp(/^[+,-]?([1-9]\d*|0)(\.\d+)?$/);
  if (RegEx.test(FloatNum)) {
    var IntNum = Math.floor(parseFloat(FloatNum, 10));
    var IntNumStr = IntNum.toString()
        .replace( /(\d)(?=(\d\d\d)+(?!\d))/g, '$1,');
    var SmallNum = (parseFloat(FloatNum, 10) - IntNum); // Bug������?
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

// �o�C�g������𐔒l�ɕϊ��A�I�v�V�����ŒP�ʕϊ�
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
  // DataType�ɉ����Ė߂�l��Ԃ��܂��B
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

// �w��b�ҋ@
function WaitSecond(WaitSec, TextMsg) {
  TextMsg = (typeof TextMsg !== "undefined") ? TextMsg : " �b�ҋ@���܂�";
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

// ������z��
function GetArrayArgs(){
  var ScriptArgArray = [];
  for(var i = 0; i < WScript.Arguments.Count(); i++){
    ScriptArgArray.push(WScript.Arguments(i));
  }
  return ScriptArgArray;
}

// �R�}���h�̈�����z��
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

// �����񒆂̃h���C�u���w�肵���h���C�u�ɕύX
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

// �Ǘ��Ҍ��������擾
function isAdmin() {
  var wsn = new ActiveXObject("WScript.Network");
  var ResultCode = 0;
  if (isAdminUser(wsn.UserName)) { ResultCode += 1; }
  if (isAdminProcess()) { ResultCode += 2; };
  return ResultCode;
}

// ���O�C�����[�U�[�̊Ǘ��Ҍ��������擾
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

// ���s���v���Z�X�̊Ǘ��Ҍ����擾
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

// �V�K�v���Z�X�Ŏw��t�@�C�������s
function ExecuteCommand(ExeFName, AryArgs, Opt0, Opt1, Opt2) {
/*
** �o�O�Ǝv���܂����AShellExecute�֐���Powershell start-process�R�}���h
** ����RunAs�ɂ��Ǘ��Ҍ�����t�^���Ď��s����ƈ����̈������ʏ�ƕς���Ă��܂��܂��B
** �������̃p�^�[�����}�C�N���\�t�g�ɕ񍐍ς݂ł��B
** �����ł�ShellExecute�֐�����Powershell���Ǘ��Ҍ����Ŏ��s���APowershell��
** ��]����R�}���h�ƈ������������Ă����s��������Ă��܂��B
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

// url�ڑ����html����^�C�g�����擾
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

// �f�X�N�g�b�v�ɃV���[�g�J�b�g�쐬
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
  // �V���[�g�J�b�g�����N�I�[�e�[�V�����ň͂܂�Ă����ꍇ�A����
  var RegEx = /^["']?(.*?)["']?$/g;
  ShortcutName = ShortcutName.replace(RegEx, "$1");
  // �t�q�k�����񂪔���
  var RegEx = /^https?:\/\/[\w!\?/\+\-_~=;\.,\*&@#\$%\(\)'\[\]]+$/g;
  if (TargetFile.search(RegEx) === 0) {
    // �V���[�g�J�b�g�����󂾂����ꍇ�A�C���^�[�l�b�g����^�C�g�����擾
    if (ShortcutName === "") {
      ShortcutName = GetHPTitle(TargetFile);
    }
    // �t�q�k�Ƃ��ăV���[�g�J�b�g����ݒ�
    Shortcut_Full_Name = ws.SpecialFolders("Desktop") + "\\"
        + ShortcutName + ".url";
  // �t�q�k�ł͂Ȃ��ꍇ�A�t�@�C���܂��̓t�H���_�̑��݂��`�F�b�N���^�[�Q�b�g�����擾
  } else if (sfo.FileExists(TargetFile) || sfo.FolderExists(TargetFile)) {
    if (ShortcutName === "") {
      ShortcutName = sfo.GetBaseName(TargetFile);
    }
    // �t�@�C���܂��̓t�H���_�Ƃ��ăV���[�g�J�b�g����ݒ�
    Shortcut_Full_Name = ws.SpecialFolders("Desktop") + "\\"
        + ShortcutName + ".lnk";
  } else {
    DebugMsg("Shortcut Target Not Found.");
  }
  // ���łɓ����V���[�g�J�b�g���̃t�@�C��������ꍇ�A�폜
  if (sfo.FileExists(Shortcut_Full_Name)) {
    sfo.DeleteFile(Shortcut_Full_Name, true);
  }
  // �V���[�g�J�b�g�쐬
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

// ���΃p�X�ŃV���{���b�N�����N�쐬
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
  // �h���C�u�������ɍ��킹�ĕύX
  LinkFile = LinkFile.replace(/^\\/, "");
  LinkFile = LinkFile.replace(/^[a-z]:\\/gi, "");
  TargetFile = TargetFile.replace(/^\\/, "");
  TargetFile = TargetFile.replace(/^[a-z]:\\/gi, "");
  DebugMsg("\nRelative MKLink");
  DebugMsg("  Link   File : " + RunningDrv + LinkFile);
  DebugMsg("  Target File : " + RunningDrv + TargetFile);
  // ���s���̃v���Z�X���Ǘ��Ҍ����������Ă��邩�m�F
  if(!isAdminProcess()) {
    DebugMsg("\nError : Process Not Administrator.");
    ErrFlg += 1;
  }
  // �����N�t�@�C�������݂��邩�m�F
  if (sfo.FileExists(RunningDrv + LinkFile)) {
    DebugMsg("\nError : Exist Link File.");
    ErrFlg += 2;
  }
  // �����N�t�@�C���̃t�H���_�����݂��邩�m�F
  if (!sfo.FolderExists(sfo.GetParentFolderName(RunningDrv + LinkFile))) {
    DebugMsg("\nError : Link Folder Not Found."
        + sfo.GetParentFolderName(RunningDrv + LinkFile));
    ErrFlg += 4;
  }
  // �^�[�Q�b�g�t�@�C�������݂��邩�m�F
  if (!sfo.FileExists(RunningDrv + TargetFile)) {
    DebugMsg("\nError : Target File Not Found.");
    ErrFlg += 8;
  }
  // �G���[���Ȃ���΃����N�쐬
  if(ErrFlg === 0) {
    // �����N�t�@�C���ƃ^�[�Q�b�g�t�@�C�������Έʒu�ɂȂ�悤�Ƀt�H���_�̎w��𐮌`
    MKLink_Cmd = MKLink_Cmd + " \""
        + RunningDrv + LinkFile + "\" \""
        + Array(LinkFile.split("\\").length).join("..\\")
        + TargetFile + "\"";
    DebugMsg("  Command     : " + MKLink_Cmd);
    ws.Run(MKLink_Cmd, WINSTYLE.HIDDEN, MODE.NOT_WAIT);
  }
  DebugMsg("");
}

// ���z�t�@�C�����}�E���g�A�܂��͍œK���i�R���p�N�g�j���܂�
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
  // �v���Z�X���Ǘ��Ҍ����������Ă��邩�m�F
  if (isAdminProcess() !== true) {
    DebugMsg("\nError : Process Not Administrator.");
    return VDError.NotAdministrator;
  }
  // �I�v�V�����̊m�F�A�G���[�̏ꍇ�ł��X���[���Ă��܂��B
  VDOpt = (typeof VDOpt !== "undefined")
         ? VDOpt.toString().toUpperCase() : "";
  if (VDOpt !== "COMPACT") {
    if (VDOpt.length === 1) { VDOpt = VDOpt + ":"; }
    VDOpt = sfo.GetDriveName(VDOpt);
  }
  // ���z�t�@�C�������PATH�ł͂Ȃ��ꍇ�A�J�����g�t�H���_�ɐݒ�
  if (sfo.GetParentFolderName(VDiskName) === "") {
    VDiskName = sfo.GetParentFolderName(WScript.ScriptFullName)
        + "\\" + VDiskName;
  }
  // ���z�t�@�C�����m�F
  if (!sfo.FileExists(VDiskName)) {
    DebugMsg("\nError : VDISK Not Found. [" + VDiskName + "]\n");
    return VDError.FileNotFound;
  }
  // DiskPart�̃R�}���h�쐬
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
  // �g�p���̃{�����[���ԍ��A�h���C�u���^�[���m�F
  var VolumeNums = [];
  var ConnectedDrvs = [];
  RegEx = new RegExp("\\s+Volume\\s+(\\d+)\\s+([A-Z])?\\s+", "gi");
  ResultStr = ws.Exec(VolumeCheck).StdOut.ReadAll();
  while ((ResultRegEx = RegEx.exec(ResultStr)) !== null) {
    VolumeNums.push(ResultRegEx[1]);
    ConnectedDrvs.push(ResultRegEx[2]);
  }
  // �w�肵���h���C�u���^�[�����łɎg�p����Ă��邩�m�F
  if (ConnectedDrvs.indexOf(VDOpt) > NotFound) {
    DebugMsg("The specified drive \"" + VDOpt + "\" is in use.");
    VDError.UseDriveLetter;
  }
  // ���z�t�@�C�����}�E���g���{�����[���ԍ����擾
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
  // DiskPart�̃R�}���h�쐬
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
  // �h���C�u���^�[���w�肳��Ă����ꍇ�A�h���C�u���^�[��ύX
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

// ���z�t�@�C�����A���}�E���g
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
  // �v���Z�X���Ǘ��Ҍ����������Ă��邩�m�F
  if (isAdminProcess() !== true) {
    return VDError.NotAdministrator;
  }
  // ���z�t�@�C�����m�F
  if (sfo.GetParentFolderName(VDiskName) === "") {
    VDiskName = sfo.GetParentFolderName(WScript.ScriptFullName)
        + "\\" + VDiskName;
  }
  if (!sfo.FileExists(VDiskName)) {
    DebugMsg("\nError : VDisk Not Found. [" + VDiskName + "]\n");
    return VDError.FileNotFound;
  }
  // DiskPart�̃R�}���h�쐬
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

// �ݒ���̃t�@�C���쐬
function CreateConfigFile(ConfigFileName) {
  var ConfigText = function() {/*
#
# ������ �g�ъ��쐬�x���X�N���v�g ������
# -Mobile Environment Creation Support Script-
# Copyright(c) 2021 Aromatibus (https://github.com/Aromatibus)
#
# USB�������≼�z�f�B�X�N�Ȃǌg�щ\�Ȋ��Ń|�[�^�u���ł̃\�t�g�E�F�A�����������߂�
# ���쐬������`�����܂��B
#
# ���ł��邱��
# �@�E�X�N���v�g�����s���ꂽ�h���C�u�܂��͎w�肵���h���C�u����ɂ������ݒ肪�ł��܂��B
# �@�E�ݒ肵�����ϐ��̓V���b�g�_�E����ɐݒ�O�̏�Ԃɖ߂�܂��B
# �@�E�f�X�N�g�b�v�ɃV���[�g�J�b�g���쐬�ł��܂��B
# �@�E���΃p�X�w��̃V���{���b�N�����N���ȒP�ɍ쐬�ł��܂��B
# �@�E���z�f�B�X�N�t�@�C���iVHD, VHDX�j���w��h���C�u�Ƀ}�E���g�A�A���}�E���g���ł��܂��B
# �@�E�v���O�����̎��s���ł��܂��B
# �@�E�����܂��̓t�@�C�����h���b�N�A���h�h���b�v���Ď��s����ƈ��������s
# �@�@�܂��͈������w�肵�ăX�N���v�g�����s�ł��܂��B
# �@���ݒ肵�����ϐ��͓��v���Z�X�A����шȑO�Ɏ��s���ꂽ�v���Z�X�ɂ͉e�����܂���B
# �@���ݒ肵�����ϐ���Volatile���ϐ��ɋL�^����邽�߁A�V���b�g�_�E����ɐݒ�O�̏�Ԃ�
# �@�@�߂�i�폜�����j�͂��ł����A�߂�Ȃ��ꍇ���m�F����Ă��܂��B
# �@�@���̏ꍇ�͖{�X�N���v�g�ō폜���邱�Ƃ��ł��܂��B
#
# ���X�N���v�g�̃t�@�C�����͕ύX�ł��܂��B
# �@�E�g�p�������\�t�g�E�F�A�ɍ��킹�����O�Ŏg�p�ł��܂��B
# �@�E�ݒ�t�@�C����"�X�N���v�g��.txt"���ǂݍ��܂�܂��B
# �@�E�g���q�̓o�b�`�t�@�C����"CMD"��"BAT"�AJScript��"JS"���g�p�ł��܂��B
# �@�@�o�b�`�t�@�C���̊g���q��Windows2000�ȍ~�ł͗����g�p�ł��܂��B
# �@���Q��ނ̃o�b�`�t�@�C���ɂ̓G���[���x���̈����ɈႢ������܂��B
# �@�E�g���q��"JS"�ɂ���ꍇ�́A���s���̃��b�Z�[�W���\���ł��Ȃ����߃X�N���v�g��
# �@�@"var Debug_Flg = true;"��"true"��"false"�ɕύX���Ȃ��ƃG���[�ɂȂ�܂��B
#
# ���ݒ�t�@�C���ɂ���
# �@�X�N���v�g����"var OneFileMode = false;"��"false"��"true"�ɂ����
# �@�ݒ�t�@�C�����O���ɍ�炸�X�N���v�g���̐ݒ����ǂݍ��݂܂��B
#
# �����ϐ�"PATHEXT"�ɂ���
# �@�R�}���h�v�����v�g�ł͓����̃t�@�C�����������ꍇ�A���s�����g���q�̗D�揇�ʂ͈�ʓI�ɁA
# �@"EXE" > "BAT" > "CMD" > "VBS" > "JS"�ɂȂ�܂��B
# �@���ڍׂ͊��ϐ�"PATHEXT"���m�F���Ă��������B
# �@�E�T���v���̐ݒ�t�@�C���ł͊��ϐ�"PATHEXT"��ύX���A�g���q"CMD"���ŗD��Ɏ��s
# �@�@�����悤�ɕύX���Ă��܂��B
# �@�E�܂��A�g���q"PY"�������Ď��s�t�@�C���Ƃ��Ĉ�����悤�ɂ��Ă��܂��B
# �@�E"CMD"��{�X�N���v�g�̂悤��JScript�^�C�v�̃o�b�`�t�@�C���ƌ��߂Ă����Ȃǂ����
# �@�@��ʂ��₷���Ȃ�܂��B
#
# ���ݒ�t�@�C���̏������͖{�ݒ�t�@�C�����Q�l�ɂ��Ă��������B
# �@"#"����n�܂�s�̓R�����g�s�ƂȂ�܂��B
#
# ��"RunAdministrator=true"�̍s���������ꍇ
# �@���s�����v���Z�X���Ǘ��Ҍ����������Ă��Ȃ��ꍇ�A�Ǘ��Ҍ����ōĎ��s����܂��B
# �@�f�t�H���g�̐ݒ�̓I�t�ifalse�j�ł��B
# �@�Ǘ��Ҍ������K�v�ȃR�}���h��v���O���������s����ꍇ�Ɏw�肵�Ă��������B
# �@�����̃R�}���h�͍ŗD��Ŏ��s����܂��B
#
# ��"ReferenceDrive=�h���C�u��"�̍s���������ꍇ
# �@��ƂȂ�Q�Ɛ�h���C�u��ύX���܂��B
# �@�f�t�H���g�̓X�N���v�g�����s�����h���C�u�ł��B
#
# �����ϐ�"PATH"�̐ݒ�i�t�H���_��ݒ肵���ꍇ�j
# �@�h���C�u����ύX�����ϐ�"PATH"�ɒǋL����܂��B
# �@"Execute"�����"ExecuteArg"�����s���Ȃ�������ϐ�"PATH"�͍Ō�ɐݒ肳��܂��B
# �@�h���C�u���̓f�t�H���g�ł̓X�N���v�g�����s�����h���C�u���ɏ����������܂��B
# �@���Ƃ��΃X�N���v�g�����s���ꂽ�h���C�u���c�h���C�u�������ꍇ�́A
# �@a:\abc -> D:\abc
# �@\abc   -> D:\abc
# �@��L�̂悤�Ƀh���C�u�����u���܂��͍s�̐擪��"\"�Ŏn�܂��Ă����ꍇ�͕⊮����܂��B
# �@���ϐ�"PATH"�ɓ����t�H���_������A�܂��̓t�H���_�����݂��Ȃ��ꍇ�͖�������܂��B
# �@�E�X�N���v�g�Őݒ肵���t�H���_�̎Q�ƗD��x�͒Ⴍ�Ō�ɎQ�Ƃ���܂��B
# �@�@�����̃t�@�C�����D��x�̍����ʂ̃t�H���_�ɂ���ꍇ�A�����炪�D�悳��܂��B
# �@�E�A�v�����s�G�C���A�X���L���̏ꍇ
# �@�@Python.exe�Ȃǈꕔ�̃\�t�g�̓A�v�����s�G�C���A�X�Őݒ肳�ꂽ"PATH"���D�悳���
# �@�@���ߎQ�Ə������ƂɂȂ�w�肵�����Ŏ��s�ł��Ȃ��ꍇ������܂��B
# �@�@���O�ɑ���PATH�ŕʖ��̃V���{���b�N�����N�����ȂǑ΍􂪕K�v�ł��B
# �@���ŏ��ɍ쐬�����ݒ�t�@�C���ł�"PYT"�̖��O�ŃV���{���b�N�����N���쐬���Ă��܂��B
# �@�@"Python.exe"�̑����"PYT"�����s���邱�Ƃŉ���ł��܂��B
#
# ��":=:"���܂ލs���������ꍇ
# �@":=:"�����ɍ��ӂ����ϐ����A�E�ӂ�ϐ��Ƃ��Ċ��ϐ��ɐݒ肵�܂��B
# �@���ϐ��������łɎg�p����Ă���ꍇ�͏㏑������܂����V���b�g�_�E����ɖ߂�܂��B
# �@�E�ӂ��w�肵�Ȃ��ꍇ�A���ϐ��̍폜�����܂��B
# �@��"PATH"�̐ݒ�Ɠ����悤�Ƀh���C�u���͏����������܂��B
# �@���h���C�u���̏��������͕ϐ��݂̂ł��B
#
# ��"RemoveVolatile=true"�̍s���������ꍇ
# �@���W�X�g������Volatile���ϐ������ׂč폜���܂��B
# �@�ċN����AVolataile���ϐ��̓��Z�b�g�����d�l�̂͂��ł����A���Z�b�g����Ȃ��ꍇ��
# �@���邽�߂ɗp�ӂ����R�}���h�ł��B
# �@�ʏ��":=:"���g�p���č폜���Ă��������B
# �@���ċN������ƃ��W�X�g����Volatile���ϐ��̓f�t�H���g�̒l�ɖ߂�܂��B
# �@�@�R�}���h�̎��s���ʂ͎��ȐӔC�ɂȂ�܂��B
#
# ��"DesktopShortcut(������A, ������B)"�̍s���������ꍇ
# �@���[�U�[�̃f�X�N�g�b�v�ɃV���[�g�J�b�g���쐬���܂��B
# �@"PATH"�̐ݒ�Ɠ����悤�Ƀh���C�u���͏����������܂��B
# �@������A�̓V���[�g�J�b�g���ɂȂ�A�����̃V���[�g�J�b�g���������ꍇ�͏㏑������܂��B
# �@�V���[�g�J�b�g����""�݂͂ŏȗ�����ƁA�t�@�C�����A�t�H���_���A�t�q�k�̃^�C�g���ɂȂ�܂��B
# �@������B�̓����N���̃t�@�C�����A�t�H���_���܂���"http"����n�܂�t�q�k���w�肵�܂��B
# �@��"PATH"�̐ݒ�Ɠ����悤�Ƀh���C�u���͏����������܂��B
#
# ��"RelativeMKLink(�����N�t�@�C����, �����N���t�@�C���̐�΃p�X)"�̍s���������ꍇ
# �@�����N�t�@�C�����Ŏw�肵�����O�ŃV���{���b�N�����N���쐬���܂��B
# �@�����N���t�@�C���̐�΃p�X�̃t�@�C���̈ʒu�������N�t�@�C�����Ƃ̑���PATH�Ń����N��
# �@�쐬����邽�߁A�h���C�u�����ς�����ꍇ�ł������N�͐���ɋ@�\���܂��B
# �@��"PATH"�̐ݒ�Ɠ����悤�Ƀh���C�u���͏����������܂��B
# �@���Ǘ��Ҍ������������v���Z�X�Ŏ��s����K�v������܂��B
#
# ��"VDiskAttch(���z�f�B�X�N�t�@�C����, �ڑ�����h���C�u��)"�̍s���������ꍇ
# �@���z�f�B�X�N�t�@�C������ڑ�����h���C�u���̃h���C�u�Ƀ}�E���g���܂��B
# �@���z�f�B�X�N�t�@�C�����Ƀp�X���ȗ�����ƃX�N���v�g�Ɠ����t�H���_����ǂݍ��݂܂��B
# �@�ڑ�����h���C�u����"Compact"���w�肷��Ɖ��z�f�B�X�N���œK�����܂��B
# �@��Diskpart�R�}���h�ō쐬�������z�f�B�X�N�t�@�C���AVHD, VHDX�œ��쌟�؂��Ă��܂��B
# �@��"PATH"�̐ݒ�Ɠ����悤�Ƀh���C�u���͏����������܂��B
# �@���Ǘ��Ҍ������������v���Z�X�Ŏ��s����K�v������܂��B
#
# ��"VDiskDetch(���z�f�B�X�N�t�@�C��)"�̍s���������ꍇ
# �@���z�f�B�X�N�t�@�C�����A���}�E���g���܂��B
# �@���z�f�B�X�N�t�@�C�����Ƀp�X���ȗ�����ƃX�N���v�g�Ɠ����t�H���_����ǂݍ��݂܂��B
# �@��"PATH"�̐ݒ�Ɠ����悤�Ƀh���C�u���͏����������܂��B
# �@���Ǘ��Ҍ������������v���Z�X�Ŏ��s����K�v������܂��B
#
# ��"WaitSec(����)"�̍s���������ꍇ
# �@�����̕b���̂������������~���ҋ@���܂��B
#
# ��"Execute(������)"�̍s���������ꍇ
# �@���O�܂łɐݒ肳�ꂽ���ϐ������p���A����������s���܂��B
# �@�ȍ~�A�ݒ肳�ꂽ���ϐ��͉e�����܂���B
# �@�A�v���P�[�V�����ȊO�͊֘A�t�����ꂽ�g���q�ɍ��킹�ĕ���������s���܂��B
# �@��"PATH"�̐ݒ�Ɠ����悤�Ƀh���C�u���͏����������܂��B
#
# ��"ExecuteArg(������)"�̍s���������ꍇ
# �@���O�܂łɐݒ肳�ꂽ���ϐ������p���A�����t���ŕ���������s���܂��B
# �@�ȍ~�A�ݒ肳�ꂽ���ϐ��͉e�����܂���B
# �@�A�v���P�[�V�����ȊO�͊֘A�t�����ꂽ�g���q�ɍ��킹�ĕ���������s���܂��B
# �@�������w�肳��Ă��Ȃ��ꍇ�͈��������Ŏ��s����܂��B
# �@��"PATH"�̐ݒ�Ɠ����悤�Ƀh���C�u���͏����������܂��B
#
# �������A�܂��̓t�@�C�����h���b�O�A���h�h���b�v���Ď��s���ꂽ�ꍇ
# �@�h���C�u���͏��������܂���B
# �@ExecuteArg(������)�̈����ɂȂ�܂��B
# �@ExecuteArg(������)���g�p����Ă��Ȃ��ꍇ�A�������̕�����������̍Ō�Ɏ��s���܂��B
# �@�������ȍ~�͑������̈����ɂȂ�܂��B
#
# �����C�Z���X
# �@Copyright(c) 2021 Aromatibus (https://github.com/Aromatibus)
# �@Released under the MIT license.
# �@See https://opensource.org/licenses/MIT
#

# �Ǘ��Ҍ����Ŏ��s
RunAdministrator=true

# �A�v�����s�G�C���A�X���p�APython.exe�̃G�C���A�X�쐬
#RelativeMKLink(a:\Python\PYT, a:\Python\Python.exe)

# ���ϐ� "PATH" �փt�H���_��ǉ�
\VSCode
\VSCode\MinGW\bin
\VSCode\node.js
\VSCode\Git\bin
\VSCode\Git\cmd
\Python
\Python\Scripts

# ���ϐ���ǉ�
NODE_PATH:=:a:\VSCode\node.js\node_modules\npm\node_modules;a:\VSCode\node.js\node_modules\npm
GIT.PATH:=:a:\VSCode\git\bin\git.exe

# ���s�ł���g���q�ƗD��x��ύX > �D��(CMD)�A�ǉ�(PY)
PATHEXT:=:.CMD;.COM;.EXE;.BAT;.VBS;.VBE;.JS;.JSE;.WSF;.WSH;.MSC;.PY

# �f�X�N�g�b�v�փV���[�g�J�b�g���쐬
DesktopShortcut(VSCode Portable, a:\VSCode\Code.exe)
DesktopShortcut("", a:\Source)
DesktopShortcut("", https://github.com/Aromatibus)

# ���z�f�B�X�N��ڑ������z�f�B�X�N����t�@�C�����N��
ReferenceDrive = c
vdiskattach(a:\Users\Aromatibus\Documents\test.vhd, z)
ReferenceDrive = z
execute(a:\TestEnv.cmd)

# �Q�Ɛ���f�t�H���g�ɖ߂�
ReferenceDrive =

# NexusFont���N���A�����Ȃ�
Execute(a:\Fonts\NexusFont\NexusFont.exe)

# NexusFont�̋N���҂�
WaitSec(5)

# VSCode���N���A��������
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
