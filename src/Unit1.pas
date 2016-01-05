// NI FastPaste - program to quickly insert predefined text strings

// Copyright (C) 2002-2015 - Nikolai Ivanov

// This program is free software; you can redistribute it and/or
// modify it under the terms of the GNU General Public License
// as published by the Free Software Foundation; either version 2
// of the License, or (at your option) any later version.

// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.

// You should have received a copy of the GNU General Public License
// along with this program; if not, write to the Free Software Foundation,
// 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
//

unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, AppEvnts, ExtCtrls, Menus, ComCtrls, ToolWin, StdCtrls,
  ImgList, ActiveX, ShlObj, IniFiles, ShellApi, clipbrd, Registry, DKLang;

type
  TForm1 = class(TForm)
    TrayIcon1: TTrayIcon;
    ApplicationEvents1: TApplicationEvents;
    PopupMenu1: TPopupMenu;
    mSettings: TMenuItem;
    mExit: TMenuItem;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ScrollBox1: TScrollBox;
    ImageList1: TImageList;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    ImageList2: TImageList;
    ImageList3: TImageList;
    GroupBox1: TGroupBox;
    Method: TComboBox;
    dklcMain: TDKLanguageController;
    N6: TMenuItem;
    mLanguage: TMenuItem;
    Donate1: TMenuItem;
    procedure mSettingsClick(Sender: TObject);
    procedure mExitClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ApplicationEvents1Minimize(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure ToolButton1Click(Sender: TObject);
    procedure Edit1Change(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure ToolButton3Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Donate1Click(Sender: TObject);
    procedure UpdateStateNotify(Sender: TObject);
  private
    { Private declarations }
    procedure HideForm;
    procedure ShowForm;
    function FileVersion: string;
    function GetSpecFolder(nFolder: integer): string;
    procedure SetControls(Disable: Boolean = true);
    procedure RegisterHotKeys;
    procedure UnRegisterHotKeys;
    procedure OnHotKey(var Msg: TWMHotKey); message WM_HOTKEY;
    procedure WMQueryEndSession(var Msg: TWMQueryEndSession); message WM_QUERYENDSESSION;
    // Updates form interface
    procedure UpdateState;
    // Language item click handler
    procedure LanguageItemClick(Sender: TObject);

  public
    { Public declarations }
    procedure LoadSettings;
    procedure SaveSettings;
    procedure ClearSettings;
    procedure PasteInFocus(const s: string; const methd: integer);
    procedure PostKeyEx32(key: Word; const shift: TShiftState; specialkey: Boolean);
    // messages
    procedure ShowMessage(Msg: string);
    procedure ShowError(Msg: string);
    function ShowConfirm(Msg: string): Boolean;
    function ShowConfirmYNC(Msg: string): byte;
  end;

const
  reccnt = 100;
  RegK = '\Software\Microsoft\Windows\CurrentVersion\Run';
  RegV = 'NI FastPaste';

var
  Form1: TForm1;
  closeapp: Boolean = false; // closing app
  dataloading: Boolean = false; // loading data
  // some dirs
  AppDir, AppData, IniF: string;
  PAutoStart: Boolean; // auto start
  MyHotKeys: array [1 .. reccnt] of integer; // hotkeys
  gb: array [1 .. reccnt] of TGroupBox;
  ed: array [1 .. reccnt] of TEdit;
  lb: array [1 .. reccnt] of TLabel;
  hk: array [1 .. reccnt] of THotKey;

implementation

{$R *.dfm}

// messages
procedure TForm1.ShowMessage(Msg: string); // info
begin
  try
    MessageBox(0, PChar(Msg), PChar(Form1.Caption), MB_OK + MB_ICONINFORMATION)
  except
  end;
end;

procedure TForm1.ShowError(Msg: string); // error
begin
  try
    MessageBox(0, PChar(Msg), PChar(Form1.Caption), MB_OK + MB_ICONERROR)
  except
  end;
end;

function TForm1.ShowConfirm(Msg: string): Boolean; // confirm
begin
  Result := false;
  try
    if MessageBox(0, PChar(Msg), PChar(Form1.Caption),
      MB_YESNO + MB_ICONQUESTION) = ID_YES then
      Result := true
    else
      Result := false
  except
  end;
end;

function TForm1.ShowConfirmYNC(Msg: string): byte;
// confirm Yes / No / Cancel
begin
  Result := 2;
  try
    case MessageBox(0, PChar(Msg), PChar(Form1.Caption),
      MB_YESNOCANCEL + MB_ICONQUESTION) of
      ID_YES:
        Result := 1;
      ID_NO:
        Result := 0;
    else
      Result := 2
    end;
  except
  end;
end;

// version - begin
function GetVersion(const FileName: String = '';
  const Fmt: String = '%d.%d.%d.%d'): String;
var
  sFileName: String;
  iBufferSize: DWORD;
  iDummy: DWORD;
  pBuffer: Pointer;
  pFileInfo: Pointer;
  iVer: array [1 .. 4] of Word;
begin
  // set default value
  Result := '';
  // get filename of exe/dll if no filename is specified
  sFileName := FileName;
  if (sFileName = '') then
  begin
    // prepare buffer for path and terminating #0
    SetLength(sFileName, MAX_PATH + 1);
    SetLength(sFileName, GetModuleFileName(hInstance, PChar(sFileName),
      MAX_PATH + 1));
  end;
  // get size of version info (0 if no version info exists)
  iBufferSize := GetFileVersionInfoSize(PChar(sFileName), iDummy);
  if (iBufferSize > 0) then
  begin
    GetMem(pBuffer, iBufferSize);
    try
      // get fixed file info (language independent)
      GetFileVersionInfo(PChar(sFileName), 0, iBufferSize, pBuffer);
      VerQueryValue(pBuffer, '\', pFileInfo, iDummy);
      // read version blocks
      iVer[1] := HiWord(PVSFixedFileInfo(pFileInfo)^.dwFileVersionMS);
      iVer[2] := LoWord(PVSFixedFileInfo(pFileInfo)^.dwFileVersionMS);
      iVer[3] := HiWord(PVSFixedFileInfo(pFileInfo)^.dwFileVersionLS);
      iVer[4] := LoWord(PVSFixedFileInfo(pFileInfo)^.dwFileVersionLS);
    finally
      FreeMem(pBuffer);
    end;
    // format result string
    Result := Format(Fmt, [iVer[1], iVer[2], iVer[3], iVer[4]]);
  end;
end;

function TForm1.FileVersion: string;
var
  v: string;
  i: integer;
begin
  try
    v := '';
    v := GetVersion;
    if length(v) > 4 then
      for i := 1 to 2 do
        if copy(v, length(v) - 1, 2) = '.0' then
          delete(v, length(v) - 1, 2);
  finally
    Result := v
  end;
end;
// version end

// lang begin
procedure TForm1.UpdateState;
const
  awsModified: Array[Boolean] of UnicodeString = ('', '*');
  //---
  procedure UpdateLanguageMark;
  var
    i: Integer;
    CurLang: LANGID; // To avoid excess synch calls
  begin
    CurLang := LangManager.LanguageID;
    for i := 0 to mLanguage.Count-1 do
      with mLanguage[i] do Checked := Tag=CurLang;
    // groupboxes & labels
    for i := 1 to reccnt do
    begin
      gb[i].Caption := DKLangConstW('Sstring') + ' ' + IntToStr(i);
	    lb[i].Caption := DKLangConstW('Shotkey') + ':';
    end;
  end;
  //---
begin
  // Update language menu
  UpdateLanguageMark;
  Method.ItemIndex := Method.Tag;
end;

procedure TForm1.UpdateStateNotify(Sender: TObject);
begin
  UpdateState;
end;

procedure TForm1.LanguageItemClick(Sender: TObject);
  //--
  procedure SaveLang;
    var
      ini: TMemIniFile;
  begin
    ini := TMemIniFile.Create(IniF, TEncoding.UTF8);
    try
      ini.WriteInteger('Params', 'LanguageID', LangManager.LanguageID);
    finally
      ini.UpdateFile;
      ini.Free
    end;
  end;
  //--
begin
  Method.Tag := Method.ItemIndex;
  // We stored language ID in Tag of each menu item (which is Sender here)
  LangManager.LanguageID := (Sender as TComponent).Tag;
  UpdateState;
  SaveLang;
end;
// lang end

function TForm1.GetSpecFolder(nFolder: integer): string;
var
  Allocator: IMalloc;
  SpecialDir: PItemIdList;
  FBuf: array [0 .. MAX_PATH] of char;
begin
  Result := '';
  try
    if SHGetMalloc(Allocator) = NOERROR then
    begin
      SHGetSpecialFolderLocation(Form1.handle, nFolder, SpecialDir);
      SHGetPathFromIDList(SpecialDir, @FBuf[0]);
      Allocator.Free(SpecialDir);
      Result := string(FBuf);
    end;
  except
    Result := '';
  end;
end;

procedure TForm1.ApplicationEvents1Minimize(Sender: TObject);
begin
  HideForm;
end;

procedure TForm1.ClearSettings;
var
  i: integer;
begin
  dataloading := true;
  try
    for i := 1 to reccnt do
    begin
      if ed[i] <> nil then
        ed[i].Text := '';
      if hk[i] <> nil then
        hk[i].HotKey := 0;
    end;
    dataloading := false;
    SetControls();
  except
    on e: exception do
    begin
      dataloading := false;
      raise exception.Create(DKLangConstW('Serclear') + ':' + #13 + e.Message);
    end;
  end;
end;

procedure TForm1.Donate1Click(Sender: TObject);
begin
  ShellExecute(0, nil, 'http://www.kivlab.com/donate/', nil, nil, 1);
end;

procedure TForm1.Edit1Change(Sender: TObject);
begin
  if not dataloading then
  begin
    SetControls(false);
  end;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
var
  Reg: TRegistry;
  progpath: string;
begin
  { create/del autostart key }
  try
    if N5.Checked <> PAutoStart then
    begin { * } { save info into registry }
      progpath := Application.ExeName + ' /t';
      Reg := TRegistry.Create;
      try
        Reg.RootKey := HKEY_CURRENT_USER;
        if Reg.OpenKey(RegK, true) then
        begin
          try
            if N5.Checked then
            begin { create key }
              Reg.WriteString(RegV, progpath);
            end
            else
            begin { delete key }
              if Reg.ValueExists(RegV) then
                Reg.DeleteValue(RegV);
            end;
            Reg.CloseKey;
          except
          end;
        end;
      finally
        Reg.Free;
        inherited;
      end;
    end; { * }
  except
  end;
end;

procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if ToolButton2.Enabled then
    if ShowConfirm(DKLangConstW('Ssavchan')) then
      try
        SaveSettings;
      except
        on e: exception do
        begin
          CanClose := false;
          raise exception.Create(e.Message);
        end;
      end
    else
      LoadSettings;
  if closeapp then
    CanClose := true
  else
  begin
    CanClose := false;
    HideForm;
    RegisterHotKeys;
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  Reg: TRegistry;
  i: integer;
  s: string;
  // lang begin
  procedure CreateLanguageMenu;
    var
      i: Integer;
      mi: TMenuItem;
  begin
    for i := 0 to LangManager.LanguageCount-1 do begin
      mi := NewItem(LangManager.LanguageNames[i], 0, False, True, LanguageItemClick, 0, '');
      mi.Tag := LangManager.LanguageIDs[i];
      mLanguage.Add(mi);
    end;
  end;
  // lang end
begin
  HideForm;
  try
    AppDir := ExtractFilePath(Application.ExeName);
  except
  end;
  try
    AppData := GetSpecFolder(CSIDL_APPDATA) + '\NI FastPaste\';
    if not DirectoryExists(AppData) then
      if not ForceDirectories(AppData) then
        AppData := AppDir;
  except
  end;
  // lang begin
  try
    // Scan for language files in the app directory and register them in the LangManager object
    LangManager.ScanForLangFiles(ExtractFileDir(ParamStr(0)), '*.lng', False);
     // Create available languages menu
    CreateLanguageMenu;
     // Update interface elements
    UpdateState;
  except end;
  // lang end
  { autostart - check }
  try
    PAutoStart := false;
    Reg := TRegistry.Create;
    try
      Reg.RootKey := HKEY_CURRENT_USER;
      if Reg.OpenKey(RegK, true) then
      begin
        try
          if Reg.ValueExists(RegV) then
          begin
            N5.Checked := true;
            PAutoStart := true;
          end;
          Reg.CloseKey;
        except
        end;
      end;
    finally
      Reg.Free;
      inherited;
    end;
  except
  end;
  IniF := AppData + 'fastpaste.ini';
  Form1.Caption := 'NI FastPaste ' + FileVersion;
  // dynamically created components
  for i := 1 to reccnt do
  begin
    s := IntToStr(i);
    // groupbox
    gb[i] := TGroupBox.Create(ScrollBox1);
    gb[i].Caption := DKLangConstW('Sstring') + ' ' + s;
    gb[i].Height := 43;
    gb[i].Top := (i - 1) * gb[i].Height + 57;
    gb[i].Width := 630;
    gb[i].Parent := ScrollBox1;
    // edit
    ed[i] := TEdit.Create(gb[i]);
    ed[i].Parent := gb[i];
    ed[i].Height := 21;
    ed[i].Left := 11;
    ed[i].Top := 16;
    ed[i].Width := 390;
    ed[i].OnChange := Edit1Change;
    // label
    lb[i] := TLabel.Create(gb[i]);
    lb[i].Parent := gb[i];
    lb[i].Height := 13;
    lb[i].Left := 413;
    lb[i].Top := 18;
    lb[i].Width := 93;
    lb[i].Caption := DKLangConstW('Shotkey') + ':';
    // hotkey
    hk[i] := THotKey.Create(gb[i]);
    hk[i].Parent := gb[i];
    hk[i].Height := 19;
    hk[i].Left := 512;
    hk[i].Top := 16;
    hk[i].Width := 105;
    hk[i].HotKey := 0;
    hk[i].OnChange := Edit1Change;
  end;
  UnRegisterHotKeys;
  LoadSettings;
  RegisterHotKeys;
  if not FileExists(IniF) then
    TrayIcon1.ShowBalloonHint;
end;

procedure TForm1.HideForm;
begin
  Hide();
  WindowState := wsMinimized;
end;

procedure TForm1.LoadSettings;
var
  i: integer;
  ini: TMemIniFile;
  s: string;
begin
  ClearSettings;
  dataloading := true;
  ini := TMemIniFile.Create(IniF, TEncoding.UTF8);
  try
    i:=ini.ReadInteger('Settings', 'Method', 3);
    if (i < 0) or (i > Method.Items.Count) then i:=3;
    Method.ItemIndex := i;
    Method.Tag := i;
    for i := 1 to reccnt do
    begin
      s := IntToStr(i);
      if ed[i] <> nil then
        ed[i].Text := ini.ReadString('String' + s, 'Text', '');
      if hk[i] <> nil then
        hk[i].HotKey := ini.ReadInteger('String' + s, 'Hotkey', 0);
    end;
    // LanguageID
    LangManager.LanguageID := ini.ReadInteger('Params', 'LanguageID', LangManager.LanguageID);
    UpdateState;
  finally
    dataloading := false;
    SetControls();
    ini.Free
  end;
end;

procedure TForm1.mSettingsClick(Sender: TObject);
begin
  UnRegisterHotKeys;
  ShowForm;
end;

procedure TForm1.N3Click(Sender: TObject);
var
  s: string;
begin
  s := 'NI FastPaste ' + FileVersion + ' [Freeware]' + #13 + #13 +
    DKLangConstW('Spinfo1') + #13 + DKLangConstW('Spinfo2') + #13 + DKLangConstW('Spinfo3') + '   '
    + #13#13 + 'Web: http://www.kivlab.com' + #13#13 +
    DKLangConstW('Spinfo4') + ':' + #13 + IniF + #13#13 +
    'Copyright © 2002-2015 by Nikolay Ivanov. ' + #13#13 +
    'Third Party Components:' + #13 +
    '- DKLang Localization Package (http://www.dk-soft.org/)';
  ShowMessage(Format(s, [reccnt]));
end;

procedure TForm1.N4Click(Sender: TObject);
begin
  try
    ShellExecute(0, nil, 'http://www.kivlab.com', nil, nil, 1);
  except
  end;
end;

procedure TForm1.OnHotKey(var Msg: TWMHotKey);
var
  i: integer;
  s: string;
begin
  for i := 1 to reccnt do
    if Msg.HotKey = MyHotKeys[i] then
    begin
      s := IntToStr(i);
      if ed[i] <> nil then
        PasteInFocus(ed[i].Text, Method.ItemIndex);
      Break;
    end;
end;

procedure TForm1.PasteInFocus(const s: string; const methd: integer);
var
  ah: HWnd;
  vGuiInfo: TGUIThreadInfo;
  i: integer;
begin
  Clipboard.AsText := s;
  if length(s) < 1 then
    exit;
  if methd<3 then
  begin
    vGuiInfo.cbSize := SizeOf(TGUIThreadInfo);
    GetGUIThreadInfo(GetWindowThreadProcessId(GetForegroundWindow), vGuiInfo);
    ah := vGuiInfo.hwndFocus;
    if ah > 0 then
      case methd of
        0:  // character by character
            for i := 1 to length(s) do
            begin
              SendMessage(ah, WM_CHAR, ord(s[i]), 0);
              Application.ProcessMessages;
            end;
        1:  // the whole line
            SendMessage(ah, WM_SETTEXT, 0, lparam(LPCTSTR(s)));
        2:  // using the Clipboard
            sendmessage(ah,WM_PASTE,0,0);
      end;
  end;
  if methd=3 then PostKeyEx32(Ord('V'), [ssCtrl], False);
  {if methd=4 then
  begin
    PostKeyEx32(VK_INSERT, [ssShift], False);
    Application.ProcessMessages;
    PostKeyEx32(VK_INSERT, [], False);
  end;}
end;

procedure TForm1.PostKeyEx32(key: Word; const shift: TShiftState;
  specialkey: Boolean);
 {************************************************************
* Procedure PostKeyEx32
*
* Parameters:
*  key    : virtual keycode of the key to send. For printable
*           keys this is simply the ANSI code (Ord(character)).
*  shift  : state of the modifier keys. This is a set, so you
*           can set several of these keys (shift, control, alt,
*           mouse buttons) in tandem. The TShiftState type is
*           declared in the Classes Unit.
*  specialkey: normally this should be False. Set it to True to
*           specify a key on the numeric keypad, for example.
* Description:
*  Uses keybd_event to manufacture a series of key events matching
*  the passed parameters. The events go to the control with focus.
*  Note that for characters key is always the upper-case version of
*  the character. Sending without any modifier keys will result in
*  a lower-case character, sending it with [ssShift] will result
*  in an upper-case character!
************************************************************}
 type
   TShiftKeyInfo = record
     shift: Byte;
     vkey: Byte;
   end;
   byteset = set of 0..7;
 const
   shiftkeys: array [1..3] of TShiftKeyInfo =
     ((shift: Ord(ssCtrl); vkey: VK_CONTROL),
     (shift: Ord(ssShift); vkey: VK_SHIFT),
     (shift: Ord(ssAlt); vkey: VK_MENU));
 var
   flag: DWORD;
   bShift: ByteSet absolute shift;
   i: Integer;
 begin
   for i := 1 to 3 do
   begin
     if shiftkeys[i].shift in bShift then
       keybd_event(shiftkeys[i].vkey, MapVirtualKey(shiftkeys[i].vkey, 0), 0, 0);
   end; { For }
   if specialkey then
     flag := KEYEVENTF_EXTENDEDKEY
   else
     flag := 0;
   keybd_event(key, MapvirtualKey(key, 0), flag, 0);
   flag := flag or KEYEVENTF_KEYUP;
   keybd_event(key, MapvirtualKey(key, 0), flag, 0);
   for i := 3 downto 1 do
   begin
     if shiftkeys[i].shift in bShift then
       keybd_event(shiftkeys[i].vkey, MapVirtualKey(shiftkeys[i].vkey, 0),
         KEYEVENTF_KEYUP, 0);
   end; { For }
 end; { PostKeyEx32 }

procedure ShortCutToHotKey(HotKey: TShortCut; var Key: Word;
  var Modifiers: Uint);
var
  Shift: TShiftState;
begin
  ShortCutToKey(HotKey, Key, Shift);
  Modifiers := 0;
  if (ssShift in Shift) then
    Modifiers := Modifiers or MOD_SHIFT;
  if (ssAlt in Shift) then
    Modifiers := Modifiers or MOD_ALT;
  if (ssCtrl in Shift) then
    Modifiers := Modifiers or MOD_CONTROL;
end;

procedure TForm1.RegisterHotKeys;
var
  i: integer;
  s: string;
  k: Word;
  m: Uint;
begin
  for i := 1 to reccnt do
    try
      k := 0;
      s := IntToStr(i);
      if hk[i] <> nil then
        k := hk[i].HotKey;
      if k = 0 then
        Continue;
      // obtaining a unique identifier for shortcuts
      MyHotKeys[i] := GlobalAddAtom(PWideChar('NIHotKey' + s));
      // registration of hotkeys
      ShortCutToHotKey(hk[i].HotKey, k, m);
      RegisterHotKey(handle, MyHotKeys[i], m, k);
    except
    end;
end;

procedure TForm1.mExitClick(Sender: TObject);
begin
  closeapp := true;
  Close;
end;

procedure TForm1.SaveSettings;
var
  i: integer;
  ini: TMemIniFile;
  s: string;
begin
  ini := TMemIniFile.Create(IniF, TEncoding.UTF8);
  try
    ini.WriteInteger('Settings', 'Method', Method.ItemIndex);
    for i := 1 to reccnt do
    begin
      s := IntToStr(i);
      if ed[i] <> nil then
        ini.WriteString('String' + s, 'Text', ed[i].Text);
      if hk[i] <> nil then
        ini.WriteInteger('String' + s, 'Hotkey', hk[i].HotKey);
    end;
    // LanguageID
    ini.WriteInteger('Params', 'LanguageID', LangManager.LanguageID);
  finally
    ini.UpdateFile;
    SetControls();
    ini.Free
  end;
end;

procedure TForm1.SetControls(Disable: Boolean);
begin
  ToolButton2.Enabled := not Disable;
  ToolButton3.Enabled := not Disable;
end;

procedure TForm1.ShowForm;
begin
  Show();
  WindowState := wsNormal;
  Application.BringToFront();
end;

procedure TForm1.ToolButton1Click(Sender: TObject);
begin
  if ShowConfirm(DKLangConstW('Sconfclear')) then
    ClearSettings;
end;

procedure TForm1.ToolButton2Click(Sender: TObject);
begin
  SaveSettings;
end;

procedure TForm1.ToolButton3Click(Sender: TObject);
begin
  if ShowConfirm(DKLangConstW('Sconfload')) then
    LoadSettings;
end;

procedure TForm1.UnRegisterHotKeys;
var
  i: integer;
begin
  for i := 1 to reccnt do
    try
      // unreg hotkey
      UnRegisterHotKey(handle, MyHotKeys[i]);
      // del atom
      GlobalDeleteAtom(MyHotKeys[i]);
    except
    end;
end;

// exit from app
procedure TForm1.WMQueryEndSession(var Msg: TWMQueryEndSession);
begin
try
 Msg.Result := 1;
 mExitClick(Self);
except end;
end;

end.
