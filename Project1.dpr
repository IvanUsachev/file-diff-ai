program FileDiff;

{$APPTYPE CONSOLE}

uses
  System.SysUtils, System.Classes, System.Generics.Collections, System.Math, System.Hash,
  System.NetEncoding, ComObj, Variants, Windows, ActiveX, IOUtils,
  System.Net.HttpClient, System.JSON, System.RegularExpressions;

type
  TDiffTag = (dtEqual, dtAdded, dtDeleted, dtChanged);

  TLineInfo = record
    Text: string;
    ParagraphIndex: Integer;
    PageNumber: Integer;
    Article: string;
    IsArticleHeader: Boolean;
    ArticleLevel: Integer;
    ArticleNumber: string
  end;

  TDiffOp = record
    Tag: TDiffTag;
    Lines1: TArray<TLineInfo>;
    Lines2: TArray<TLineInfo>;
  end;

function LineInfosToStrings(const Arr: TArray<TLineInfo>): TArray<string>;
var
  i: Integer;
begin
  SetLength(Result, Length(Arr));
  for i := 0 to High(Arr) do
    Result[i] := Arr[i].Text;
end;

function FirstTextOrEmpty(const Arr: TArray<TLineInfo>): string;
begin
  if Length(Arr) > 0 then
    Result := Arr[0].Text
  else
    Result := '';
end;


// LCS-based diff computation with changed detection
function ComputeDiff(const A, B: TArray<TLineInfo>): TArray<TDiffOp>;
var
  Matrix: array of array of Integer;
  i, j: Integer;
  Ops: TList<TDiffOp>;
  Op: TDiffOp;

  function Max(a, b: Integer): Integer;
  begin
    if a > b then
      Result := a
    else
      Result := b;
  end;

begin
  SetLength(Matrix, Length(A) + 1, Length(B) + 1);

  for i := 0 to Length(A) do
    Matrix[i, 0] := 0;
  for j := 0 to Length(B) do
    Matrix[0, j] := 0;

  for i := 1 to Length(A) do
    for j := 1 to Length(B) do
      if A[i - 1].Text = B[j - 1].Text then
        Matrix[i, j] := Matrix[i - 1, j - 1] + 1
      else
        Matrix[i, j] := Max(Matrix[i - 1, j], Matrix[i, j - 1]);

  Ops := TList<TDiffOp>.Create;
  try
    i := Length(A);
    j := Length(B);

    while (i > 0) or (j > 0) do
    begin
      if (i > 0) and (j > 0) and (A[i - 1].Text = B[j - 1].Text) then
      begin
        Op.Tag := dtEqual;
        Op.Lines1 := [A[i - 1]];
        Op.Lines2 := [B[j - 1]];
        Ops.Insert(0, Op);
        Dec(i);
        Dec(j);
      end
      else if (i > 0) and (j > 0) and (Matrix[i, j] = Matrix[i - 1, j - 1]) then
      begin
        Op.Tag := dtChanged;
        Op.Lines1 := [A[i - 1]];
        Op.Lines2 := [B[j - 1]];
        Ops.Insert(0, Op);
        Dec(i);
        Dec(j);
      end
      else if (j > 0) and ((i = 0) or (Matrix[i, j - 1] >= Matrix[i - 1, j])) then
      begin
        Op.Tag := dtAdded;
        SetLength(Op.Lines1, 0);
        Op.Lines2 := [B[j - 1]];
        Ops.Insert(0, Op);
        Dec(j);
      end
      else if i > 0 then
      begin
        Op.Tag := dtDeleted;
        Op.Lines1 := [A[i - 1]];
        SetLength(Op.Lines2, 0);
        Ops.Insert(0, Op);
        Dec(i);
      end;
    end;

    Result := Ops.ToArray;
  finally
    Ops.Free;
  end;
end;


// Get full path
function GetFullPath(const FileName: string): string;
begin
  Result := TPath.GetFullPath(FileName);
end;


function CountChar(const S: string; C: Char): Integer;
var
  i: Integer;
begin
  Result := 0;
  for i := 1 to Length(S) do
    if S[i] = C then
      Inc(Result);
end;


function IsArticleHeader(const Text: string; ParaStyle: string; FontSize: OleVariant;
                          out Level: Integer; out Number: string): Boolean;
var
  CleanText: string;
  Match: TMatch;
begin
  CleanText := Trim(Text);
  Level := 0;
  Number := '';
  Result := False;

  // 1. Detekce podle Word stylů (priorita)
  if (Pos('Heading', ParaStyle) > 0) then
  begin
    if (Pos('1', ParaStyle) > 0) then Level := 1
    else if (Pos('2', ParaStyle) > 0) then Level := 2
    else if (Pos('3', ParaStyle) > 0) then Level := 3
    else Level := 1;

    Number := CleanText;
    Result := True;
    Exit;
  end;

  // 2. Detekce podle velikosti písma a tučného textu
  if (VarIsNumeric(FontSize) and (FontSize >= 14)) then
  begin
    Level := 2;
    Number := CleanText;
    Result := True;
    // Neukončujeme - může být kombinace s číselným formátem
  end;

  // 3. Detekce podle číselného formátu
  Match := TRegEx.Match(CleanText, '^(\d+(?:\.\d+)*)\s+');
  if Match.Success then
  begin
    Number := Match.Groups[1].Value;
    Level := 1 + CountChar(Number, '.');  // počet teček = hloubka
    Result := True;
  end;
end;


// Extract text from TXT or DOCX
// Tested on Microsoft 365 2507
function ExtractTextFile(const FileName: string): TArray<TLineInfo>;
var
  Ext: string;
  WordApp, Doc, Para: OleVariant;
  FullPath: string;
  i: Integer;
  CurrentArticle, CurrentArticleNumber: string;
  CurrentArticleLevel: Integer;

begin
  CurrentArticle := '';
  CurrentArticleNumber := '0';
  CurrentArticleLevel := 0;

  if not FileExists(FileName) then
    raise Exception.Create('File not found: ' + FileName);

  Ext := LowerCase(ExtractFileExt(FileName));
  FullPath := GetFullPath(FileName);

  try
    if (Ext = '.docx') or (Ext = '.doc') or (Ext = '.pdf') then
    begin
      WordApp := CreateOleObject('Word.Application');
      WordApp.Visible := False;
      WordApp.DisplayAlerts := False;

      Doc := WordApp.Documents.Open(FullPath, False, True);

      SetLength(Result, Integer(Doc.Paragraphs.Count));
      for i := 1 to Doc.Paragraphs.Count do
      begin
        Para := Doc.Paragraphs.Item(i).Range;
        var ParaText := Trim(string(Para.Text));

        if ParaText = '' then
          Continue;

        var IsArticle := False;
        var ArticleLevel := 0;
        var ArticleNumber := '';

        // Získání stylu a formátování
        var ParaStyle := string(Para.Style.NameLocal);
        var FontSize := Para.Font.Size;

        // Detekce článku s využitím všech metod
        if IsArticleHeader(ParaText, ParaStyle, FontSize, ArticleLevel, ArticleNumber) then
        begin
          CurrentArticle := ParaText;
          CurrentArticleNumber := ArticleNumber;
          CurrentArticleLevel := ArticleLevel;
          IsArticle := True;
        end;

        Result[i-1].Text := ParaText;
        Result[i-1].ParagraphIndex := i;
        Result[i-1].PageNumber := Doc.Range(0, Doc.Paragraphs.Item(i).Range.End).Information[3];
        Result[i-1].Article := CurrentArticle;
        Result[i-1].IsArticleHeader := IsArticle;
        Result[i-1].ArticleLevel := CurrentArticleLevel;
        Result[i-1].ArticleNumber := CurrentArticleNumber;
      end;

      Doc.Close(False);
      WordApp.Quit;
      VarClear(Doc);
      VarClear(WordApp);
    end
    else if Ext = '.txt' then
    begin
      var SL := TStringList.Create;
      try
        SL.LoadFromFile(FileName, TEncoding.UTF8);
        SetLength(Result, SL.Count);
        for i := 0 to SL.Count - 1 do
        begin
          var ParaText := SL[i];
          var IsArticle := False;
          var ArticleLevel := 0;
          var ArticleNumber := '';

          // Pro TXT používáme pouze číselnou detekci
          if IsArticleHeader(ParaText, '', 0, ArticleLevel, ArticleNumber) then
          begin
            CurrentArticle := ParaText;
            CurrentArticleNumber := ArticleNumber;
            CurrentArticleLevel := ArticleLevel;
            IsArticle := True;
          end;

          Result[i].Text := ParaText;
          Result[i].ParagraphIndex := i;
          Result[i].PageNumber := 0;
          Result[i].Article := CurrentArticle;
          Result[i].IsArticleHeader := IsArticle;
          Result[i].ArticleLevel := CurrentArticleLevel;
          Result[i].ArticleNumber := CurrentArticleNumber;
        end;
      finally
        SL.Free;
      end;
    end
    else
      raise Exception.Create('Unsupported file type: ' + Ext);
  finally
  end;
end;


// Generate side-by-side HTML diff
procedure GenerateHTMLReport(const Ops: TArray<TDiffOp>;
  const File1, File2, OutFile: string);
var
  SL: TStringList;
  Op: TDiffOp;

  function EncodeSafe(const S: string): string;
  begin
    Result := TNetEncoding.HTML.Encode(S);
  end;

  function FirstTextOrEmpty(const Arr: TArray<TLineInfo>): string;
  begin
    if Length(Arr) > 0 then
      Result := Arr[0].Text
    else
      Result := '';
  end;

begin
  SL := TStringList.Create;
  try
    SL.Add('<html><head><meta charset="utf-8"><title>Diff Report</title>');
    SL.Add('<style>');
    SL.Add('body { font-family: Arial, sans-serif; margin: 20px; }');
    SL.Add('.container { display: flex; }');
    SL.Add('.column { width: 50%; box-sizing: border-box; white-space: pre-wrap; padding: 5px; }');
    SL.Add('.added { background-color: #d4fcdc; }');
    SL.Add('.removed { background-color: #fddcdc; }');
    SL.Add('.changed { background-color: #fff7a8; }');
    SL.Add('.empty-line { height: 1.2em; display: block; }');
    SL.Add('.legend { margin-bottom: 20px; padding: 10px; background-color: #f5f5f5; border-radius: 5px; }');
    SL.Add('.legend-item { display: inline-block; margin-right: 20px; }');
    SL.Add('.legend-color { display: inline-block; width: 20px; height: 20px; margin-right: 5px; vertical-align: middle; border: 1px solid #ccc; }');
    SL.Add('.legend-added { background-color: #d4fcdc; }');
    SL.Add('.legend-removed { background-color: #fddcdc; }');
    SL.Add('.legend-changed { background-color: #fff7a8; }');
    SL.Add('</style></head><body>');

    SL.Add('<h2>Comparing files: ' + EncodeSafe(File1) + ' and ' + EncodeSafe(File2) + '</h2>');

    // Legend
    SL.Add('<div class="legend">');
    SL.Add('<div class="legend-item"><span class="legend-color legend-added"></span>Added</div>');
    SL.Add('<div class="legend-item"><span class="legend-color legend-changed"></span>Changed</div>');
    SL.Add('<div class="legend-item"><span class="legend-color legend-removed"></span>Deleted</div>');
    SL.Add('</div>');

    SL.Add('<div class="container">');

    // Left column (File1)
    SL.Add('<div class="column"><b>' + EncodeSafe(File1) + '</b><br>');
    for Op in Ops do
    begin
      case Op.Tag of
        dtEqual:   SL.Add('<div>' + EncodeSafe(FirstTextOrEmpty(Op.Lines1)) + '</div>');
        dtDeleted: SL.Add('<div class="removed">' + EncodeSafe(FirstTextOrEmpty(Op.Lines1)) + '</div>');
        dtAdded:   SL.Add('<div class="empty-line"></div>');
        dtChanged: SL.Add('<div class="changed">' + EncodeSafe(FirstTextOrEmpty(Op.Lines1)) + '</div>');
      end;
    end;
    SL.Add('</div>');

    // Right column (File2)
    SL.Add('<div class="column"><b>' + EncodeSafe(File2) + '</b><br>');
    for Op in Ops do
    begin
      case Op.Tag of
        dtEqual:   SL.Add('<div>' + EncodeSafe(FirstTextOrEmpty(Op.Lines2)) + '</div>');
        dtAdded:   SL.Add('<div class="added">' + EncodeSafe(FirstTextOrEmpty(Op.Lines2)) + '</div>');
        dtDeleted: SL.Add('<div class="empty-line"></div>');
        dtChanged: SL.Add('<div class="changed">' + EncodeSafe(FirstTextOrEmpty(Op.Lines2)) + '</div>');
      end;
    end;
    SL.Add('</div>');

    SL.Add('</div></body></html>');

    SL.SaveToFile(OutFile, TEncoding.UTF8);
  finally
    SL.Free;
  end;
end;


function AskChatGPT(const Prompt, ApiKey: string): string;
var
  Client: THttpClient;
  JsonReq, JsonResp: TStringStream;
  RootObj, MsgObj: TJSONObject;
  MsgArray: TJSONArray;
  ChoiceObj: TJSONObject;
  ResponseCode: Integer;
begin
  Result := '';
  Client := THttpClient.Create;
  JsonResp := TStringStream.Create('', TEncoding.UTF8);
  try
    RootObj := TJSONObject.Create;
    MsgArray := TJSONArray.Create;

    MsgObj := TJSONObject.Create;
    MsgObj.AddPair('role', 'user');
    MsgObj.AddPair('content', Prompt);
    MsgArray.AddElement(MsgObj);

    RootObj.AddPair('model', 'gpt-4o-mini');
    RootObj.AddPair('messages', MsgArray);

    JsonReq := TStringStream.Create(RootObj.ToString, TEncoding.UTF8);
    try
      Client.CustomHeaders['Authorization'] := 'Bearer ' + ApiKey;
      Client.ContentType := 'application/json';

      try
        ResponseCode := Client.Post('https://api.openai.com/v1/chat/completions', JsonReq, JsonResp).StatusCode;
        Writeln('HTTP Response Code: ', ResponseCode);
        Writeln('Response: ', JsonResp.DataString);
      except
        on E: Exception do
          Writeln('HTTP Error: ', E.Message);
      end;

    finally
      JsonReq.Free;
    end;

    RootObj.Free;

    // Parse response
    if JsonResp.Size > 0 then
    begin
      RootObj := TJSONObject.ParseJSONValue(JsonResp.DataString) as TJSONObject;
      if Assigned(RootObj) then
      try
        if RootObj.TryGetValue('error', MsgObj) then
        begin
          Writeln('API Error: ', MsgObj.GetValue('message').Value);
          Exit;
        end;

        MsgArray := RootObj.GetValue('choices') as TJSONArray;
        if Assigned(MsgArray) and (MsgArray.Count > 0) then
        begin
          ChoiceObj := MsgArray.Items[0] as TJSONObject;
          if Assigned(ChoiceObj) then
          begin
            MsgObj := ChoiceObj.GetValue('message') as TJSONObject;
            if Assigned(MsgObj) then
              Result := MsgObj.GetValue('content').Value;
          end;
        end;
      finally
        RootObj.Free;
      end;
    end;
  finally
    JsonResp.Free;
    Client.Free;
  end;
end;

function ReadApiKeyFromConfig(const ConfigFile: string): string;
var
  SL: TStringList;
  i: Integer;
  Line: string;
begin
  Result := '';
  if not FileExists(ConfigFile) then
    Exit;

  SL := TStringList.Create;
  try
    SL.LoadFromFile(ConfigFile);
    for i := 0 to SL.Count - 1 do
    begin
      Line := Trim(SL[i]);
      // Skip empty lines and comments
      if (Line = '') or (Line.StartsWith(';')) or (Line.StartsWith('#')) then
        Continue;

      if Line.StartsWith('APIKey=') then
      begin
        Result := Trim(Line.Substring(7));
        Break;
      end;
    end;
  finally
    SL.Free;
  end;
end;


function GenerateDiffSummary(const Ops: TArray<TDiffOp>): string;
var
  Op: TDiffOp;
  Article, Key: string;
  Page: Integer;
  Changes: TObjectDictionary<string, TStringList>;
  OrderedKeys: TList<string>;
  SL: TStringList;
begin
  Result := '';
  Changes := TObjectDictionary<string, TStringList>.Create([doOwnsValues]);
  OrderedKeys := TList<string>.Create;
  try
    for Op in Ops do
    begin
      if Op.Tag = dtEqual then
        Continue;

      // Zjisti název článku a stránku
      if Length(Op.Lines1) > 0 then
      begin
        Article := Op.Lines1[0].Article;
        Page := Op.Lines1[0].PageNumber;
      end
      else if Length(Op.Lines2) > 0 then
      begin
        Article := Op.Lines2[0].Article;
        Page := Op.Lines2[0].PageNumber;
      end
      else
      begin
        Article := 'Neznámý článek';
        Page := 0;
      end;

      Key := Article + ' (strana ' + Page.ToString + ')';

      // Pokud článek ještě není v seznamu, přidej ho
      if not Changes.TryGetValue(Key, SL) then
      begin
        SL := TStringList.Create;
        Changes.Add(Key, SL);
        OrderedKeys.Add(Key);  // zachová pořadí podle výskytu
      end;

      // Přidej změnu do článku
      SL := Changes[Key];
      case Op.Tag of
        dtAdded:
          SL.Add('Added: ' + String.Join(' ', LineInfosToStrings(Op.Lines2)));
        dtDeleted:
          SL.Add('Deleted: ' + String.Join(' ', LineInfosToStrings(Op.Lines1)));
        dtChanged:
          SL.Add('Changed: "' + String.Join(' ', LineInfosToStrings(Op.Lines1)) +
                 '" → "' + String.Join(' ', LineInfosToStrings(Op.Lines2)) + '"');
      end;
    end;

    // Výstup jen z článků se změnami, v pořadí podle výskytu
    for Key in OrderedKeys do
    begin
      SL := Changes[Key];
      if SL.Count > 0 then
      begin
        Result := Result + sLineBreak + 'Článek: ' + Key + sLineBreak;
        Result := Result + SL.Text + sLineBreak;
      end;
    end;

  finally
    OrderedKeys.Free;
    Changes.Free;
  end;
end;


procedure AnalyzeDiffWithChatGPT(const DiffSummary: string; const ConfigFile: string);
var
  ApiKey: string;
  GPTResponse: string;
  Prompt: string;
  OutFile: string;
  SL: TStringList;
begin
  // 1. Load API key
  ApiKey := ReadApiKeyFromConfig(ConfigFile);
  if ApiKey = '' then
  begin
    Writeln('Warning: config.ini file not found or APIKey not set');
    Writeln('Please create config.ini with:');
    Writeln('APIKey=your-openai-api-key-here');
    Writeln('Skipping ChatGPT analysis');
    Exit;
  end;

  // 2. Build analysis prompt
  Prompt :=
    'Analyzuj rozdíly mezi dvěma dokumenty. Níže je seznam změn, ' +
    'kde každý záznam je označen tagem: "Added" (nově přidaný text), ' +
    '"Deleted" (odstraněný text) nebo "Changed" (upravený text).' + sLineBreak +
    'Tvým úkolem je popsat rozdíly v češtině postupně po jednotlivých článcích ' +
    'nebo odstavcích. Drž se pořadí změn a u každého článku uveď, co se v něm změnilo.' + sLineBreak +
    'Používej jasnou strukturu: pro každý článek napiš jeho číslo/název (pokud je k dispozici) ' +
    'a pod něj seznam změn (přidané, odstraněné, upravené části).' + sLineBreak +
    'Nepřepisuj celý text, pouze shrň rozdíly.' + sLineBreak +
    sLineBreak +
    'Seznam změn:' + sLineBreak +
    DiffSummary;

  // 3. Call ChatGPT API
  Writeln('API key found in config.ini');
  Writeln('Requesting analysis from ChatGPT...');

  GPTResponse := AskChatGPT(Prompt, ApiKey);

  // 4. Output response
  Writeln;
  Writeln('ChatGPT response:');
  Writeln(GPTResponse);

  // 5. Save response to file
  OutFile := 'chatgpt_analysis.txt';
  SL := TStringList.Create;
  try
    SL.Text := GPTResponse;
    SL.SaveToFile(OutFile, TEncoding.UTF8);
    Writeln('Analysis saved to file: ', OutFile);
  finally
    SL.Free;
  end;
end;


// Main program
var
  File1, File2: string;
  Text1, Text2: TArray<TLineInfo>;
  Ops: TArray<TDiffOp>;
  DiffSummary: string;
begin
  CoInitialize(nil);
  SetConsoleOutputCP(65001);
  SetConsoleCP(65001);
  try
    try
      File1 := 'SQ3251.doc';
      File2 := 'SQ3251.1.doc';

      if not FileExists(File1) then
      begin
        Writeln('Error: File not found: ', File1);
        Exit;
      end;

      if not FileExists(File2) then
      begin
        Writeln('Error: File not found: ', File2);
        Exit;
      end;

      Writeln('Extracting text from files...');
      Text1 := ExtractTextFile(File1);
      Text2 := ExtractTextFile(File2);

      Writeln('Computing differences...');
      Ops := ComputeDiff(Text1, Text2);

      Writeln('Generating HTML report...');
      GenerateHTMLReport(Ops, File1, File2, 'diff_report.html');

      Writeln('Diff report generated: diff_report.html');

      // Only proceed if we have valid text extraction
      if (Length(Text1) > 0) and (Length(Text2) > 0) then
      begin
        // Generate diff summary
        DiffSummary := GenerateDiffSummary(Ops);

        AnalyzeDiffWithChatGPT(DiffSummary, 'config.ini');
      end;

    except
      on E: Exception do
        Writeln('Error: ', E.Message);
    end;

    Writeln;
    Writeln('Press Enter to exit...');
    Readln;
  finally
    CoUninitialize;
  end;
end.
