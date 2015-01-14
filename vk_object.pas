unit vk_object;


interface
  uses
    Windows,
    ActiveX,
    SysUtils,
    Classes,
    Messages,
    Variants,
    Controls,
    MSXML,
    StdCtrls,
    DB,
    IBDatabase,
    IBCustomDataSet,
    IBQuery,
    IniFiles,
    Dialogs,
    XMLUnit;


   type TMode = (
     m_rus_name,    //Русское имя (метода или свойства)
     m_eng_name,    //Англ. имя (метода или свойства)
     m_get_value,   //1С читает значение свойства
     m_set_value,   //1С изменяет значение свойства
     m_n_params,    //1С получает число параметров метода (функции)
     m_execute      //Выполнение метода (Функции)
   );

   const
     XMLExt = '.xml';                         // расширение XML-файлов
     TXTExt = '.txt';                         // расширение TXT-файлов

  type T_vk_object = class(TObject)
  public

    C_ID_length,                          // - C_ID
    CHECK_DATE_length,                    // - CHECK_DATE
    TABEL_length,                         // - TABEL
    COST_length,                          // - COST
    DISH_NAME_length,                     // - DISH_NAME
    WEIGHT_length,                        // - WEIGHT
    MA_ID_length,                         // - MA_ID
    MENU_DATE_length : integer;           // - MENU_DATE


    g_Value: OleVariant;                  //Значение свойства или возвращаемое значение функции
    g_NParams: Integer;                   //Количество параметров метода (функции)
    g_Params: PSafeArray;                 //Массив с параметрами функции
    g_Event, g_Event_Data: String;        //Генерация события

                                          // СВОЙСТВА
    g_fname: String;                      //  

    function prop1(mode: TMode): String;
    (*11*)

    function meth1(mode: TMode): String;
    function meth2(mode: TMode): String;
    function meth3(mode: TMode): String;
    function meth4(mode: TMode): String;
    function meth5(mode: TMode): String;
    function meth6(mode: TMode): String;
    function meth7(mode: TMode): String;
    function meth8(mode: TMode): String;
    function meth9(mode: TMode): String;
    (*12*)

    Constructor Create;
    Destructor Destroy; Override;

  protected
    g_IconType: Integer;
    g_Title: String;

    DB : TIBDatabase;
    DBt : TIBTransaction;
    Query : TIBQuery;


    function GetNParam(lIndex: Integer ): OleVariant;
    procedure PutNParam(lIndex: Integer; var varPut: OleVariant);
    function GetParamAsString(lIndex: Integer ): String;
    function GetParamAsInteger(lIndex: Integer ): Integer;
    function StrToLength(str: string; l: Integer) : string;
    procedure CheckXMLExport(C_ID : integer; FileName : string);
    procedure CheckTxtExport(C_ID : integer; FileName : string);
    procedure MenuXMLExport(MA_ID : integer; FileName : string);
    procedure MenuTxtExport(MA_ID : integer; FileName : string);

  end;



implementation

//////////////////////////////////////////////////////////////
//Конструктор класса
Constructor T_vk_object.Create;
begin
  inherited Create;
  g_Value:='';
  g_NParams:=0;
  g_Params:=nil;
  g_Event:='';
  g_Event_Data:='';
end;

//////////////////////////////////////////////////////////////
//Деструктор класса
destructor T_vk_object.Destroy;
begin
    Query.Free;
    DBt.Free;
    DB.Free;
     inherited Destroy;
end;

// формантирование строки пробелами до нужной длины
function T_vk_object.StrToLength(str: string; l: Integer) : string;
var
  ii, lng : integer;
  buff : AnsiString;
begin
  buff := '';
  lng := Length(str);
  for ii := 1 to l do if ii <= lng then buff := buff + str[ii] else buff := buff + ' ';
  Result := buff;
end;

// Экспорт чека в XML
procedure T_vk_object.CheckXMLExport(C_ID : integer; FileName : string);
const
  XMLHeader = '<?xml version="1.0" encoding="windows-1251" ?>'#13#10;
var
  i, l, k, CommaCount : integer;
  XMLStream : TFileStream;
  XMLFileName, bufstr, bufw, OrderCode : string;
begin
  XMLFileName := FileName;
  try
    XMLStream := TFileStream.Create(XMLFileName, fmOpenReadWrite);
    XMLStream.Position := XMLStream.Size - length(ObjClose);
  except
    XMLStream := TFileStream.Create(XMLFileName, fmCreate);
    XMLStream.Write(XMLHeader,length(XMLHeader));
    // открываем корень
    OpenObject(XMLStream,'CheckHistory');
  end;
  // открываем новый объект
  OpenObject(XMLStream,'Check');
  WriteAttribute(XMLStream,'C_ID',IntToStr(C_ID));
  // извлечение чека из БД
  with Query do begin
    Close;
    SQL.Clear;
    SQL.Add('select * from checks where c_id = ' + IntToStr(C_ID));
    Open;
  end;
  WriteAttribute(XMLStream,'CHECK_DATE', Query.FieldByName('CHECK_DATE').AsString);
  WriteAttribute(XMLStream,'TABEL',Query.FieldByName('TABEL').AsString);
  WriteAttribute(XMLStream,'SUM',Query.FieldByName('COST').AsString);
  OrderCode := Query.FieldByName('MENU').AsString;
  // разбор кода заказа
  CommaCount := 0;
  // подсчёт разделителей и звёздочек
  for i := 1 to length(OrderCode) do begin
    if OrderCode[i] = ',' then inc(CommaCount);
  end;
  // основной цикл
  l := 1;
  for i := 1 to CommaCount do begin
    bufstr := '';
    bufw := '';
    inc(l);
    repeat
      bufstr := bufstr + OrderCode[l];
      inc(l);
      // поиск звёздочки
      if OrderCode[l] = '*' then begin
        k := l + 1;
        repeat
          bufw := bufw + OrderCode[k];
          inc(k);
        until (OrderCode[k] = ',') or (k > length(OrderCode));
        l := k;
      end;
    until (OrderCode[l] = ',') or (l > length(OrderCode));
    OpenObject(XMLStream,'Dish');
    WriteAttribute(XMLStream,'ME_ID',bufstr);
    // извлечение данных о MENU_ENTRY
    with Query do begin
      Close;
      SQL.Clear;
      SQL.Add('select dishes.dish_name, menu_entries.weight, menu_entries.cost');
      SQL.Add('from dishes, menu_entries');
      SQL.Add('where (dishes.d_id = menu_entries.d_id) and (menu_entries.me_id = '
        + bufstr + ');');
      Open;
      FetchAll;
    end;
    WriteAttribute(XMLStream,'DISH_NAME',Query.FieldByName('DISH_NAME').AsString);
    WriteAttribute(XMLStream,'WEIGHT_NOMINAL',Query.FieldByName('WEIGHT').AsString);
    WriteAttribute(XMLStream,'COST_NOMINAL',Query.FieldByName('COST').AsString);
    if bufw <> '' then begin
      WriteAttribute(XMLStream,'WEIGHT_ORDERED',bufw);
      WriteAttribute(XMLStream,'COST_ORDERED',
        IntToStr(round(StrToInt(bufw) / Query.FieldByName('WEIGHT').AsInteger
          * Query.FieldByName('COST').AsInteger)));
    end
    else begin
      WriteAttribute(XMLStream,'WEIGHT_ORDERED',Query.FieldByName('WEIGHT').AsString);
      WriteAttribute(XMLStream,'COST_ORDERED',Query.FieldByName('COST').AsString);
    end;
    CloseObject(XMLStream);
  end; // for i
  // закрываем объект
  CloseObject(XMLStream);
  // закрываем корень
  CloseObject(XMLStream);
  XMLStream.Free;
end;

// Экспорт чека в текстовый файл
procedure T_vk_object.CheckTxtExport(C_ID : integer; FileName : string);
var
  i, l, k, CommaCount : integer;
  TXTStream : TFileStream;
  StrStream : TStringStream;
  TXTFileName, bufstr, bufw, OrderCode, ExportBase, ExportTail : string;
begin
  TXTFileName := FileName;
  try
    TXTStream := TFileStream.Create(FileName, fmOpenReadWrite);
    TXTStream.Position := TXTStream.Size;
  except
    TXTStream := TFileStream.Create(FileName, fmCreate);
  end;
  // извлечение чека из БД
  with Query do begin
    Close;
    SQL.Clear;
    SQL.Add('select * from checks where c_id = ' + IntToStr(C_ID));
    Open;
  end;
  // формирование начала записи
  ExportBase := StrToLength(Query.FieldByName('C_ID').AsString,C_ID_length)
      + StrToLength(FormatDateTime('dd/mm/yyyy',Query.FieldByName('CHECK_DATE').AsVariant),CHECK_DATE_length)
      + StrToLength(Query.FieldByName('TABEL').AsString,TABEL_length);
  OrderCode := Query.FieldByName('MENU').AsString;
  // разбор кода заказа
  CommaCount := 0;
  StrStream := TStringStream.Create('');
  // подсчёт разделителей и звёздочек
  for i := 1 to length(OrderCode) do begin
    if OrderCode[i] = ',' then inc(CommaCount);
  end;
  // основной цикл
  l := 1;
  for i := 1 to CommaCount do begin
    bufstr := '';
    bufw := '';
    inc(l);
    repeat
      bufstr := bufstr + OrderCode[l];
      inc(l);
      // поиск звёздочки
      if OrderCode[l] = '*' then begin
        k := l + 1;
        repeat
          bufw := bufw + OrderCode[k];
          inc(k);
        until (OrderCode[k] = ',') or (k > length(OrderCode));
        l := k;
      end;
    until (OrderCode[l] = ',') or (l > length(OrderCode));
    // извлечение данных о MENU_ENTRY
    with Query do begin
      Close;
      SQL.Clear;
      SQL.Add('select dishes.dish_name, menu_entries.weight, menu_entries.cost');
      SQL.Add('from dishes, menu_entries');
      SQL.Add('where (dishes.d_id = menu_entries.d_id) and (menu_entries.me_id = '
        + bufstr + ');');
      Open;
    end;
    // формирование хвоста записи
    if bufw = '' then begin
      ExportTail := StrToLength(Query.FieldByName('COST').AsString,COST_length)
        + StrToLength(Query.FieldByName('DISH_NAME').AsString,DISH_NAME_length)
        + StrToLength(Query.FieldByName('WEIGHT').AsString, WEIGHT_length) + #13#10;
    end else
    begin
      ExportTail := StrToLength(IntToStr(round(StrToInt(bufw) / Query.FieldByName('WEIGHT').AsInteger
        * Query.FieldByName('COST').AsInteger)),COST_length)
        + StrToLength(Query.FieldByName('DISH_NAME').AsString,DISH_NAME_length)
        + StrToLength(bufw,WEIGHT_length) + #13#10;
    end;
    // помещение полной записи в строковый поток
    StrStream.WriteString(ExportBase + ExportTail);
  end; // for i
  // копирование строкового потока в конец потока файла
  try
    StrStream.Position := 0;
    TXTStream.Position := TXTStream.Size;
    try
      TXTStream.CopyFrom(StrStream,StrStream.Size);
    finally
      TXTStream.Free;
    end;
  finally
    StrStream.Free;
  end;
end;

// Экспорт меню в XML
procedure T_vk_object.MenuXMLExport(MA_ID : integer; FileName : string);
const
  XMLHeader = '<?xml version="1.0" encoding="windows-1251" ?>'#13#10;
var
  i, l, CommaCount : integer;
  XMLStream : TFileStream;
  XMLFileName, bufstr, MenuCode : string;
begin
  XMLFileName := FileName;
  try
    XMLStream := TFileStream.Create(XMLFileName, fmOpenReadWrite);
    XMLStream.Position := XMLStream.Size - length(ObjClose);
  except
    XMLStream := TFileStream.Create(XMLFileName, fmCreate);
    XMLStream.Write(XMLHeader,length(XMLHeader));
    // открываем корень
    OpenObject(XMLStream,'MenuHistory');
  end;
  // открываем новый объект
  OpenObject(XMLStream,'Menu');
  WriteAttribute(XMLStream,'MA_ID',IntToStr(MA_ID));
  // извлечение чека из БД
  with Query do begin
    Close;
    SQL.Clear;
    SQL.Add('select * from menu_archives where ma_id = ' + IntToStr(MA_ID));
    Open;
  end;
  WriteAttribute(XMLStream,'MENU_DATE', Query.FieldByName('MENU_DATE').AsString);
  MenuCode := Query.FieldByName('MENU').AsString;
  // разбор кода меню
  CommaCount := 0;
  // подсчёт разделителей
  for i := 1 to length(MenuCode) do begin
    if MenuCode[i] = ',' then inc(CommaCount);
  end;
  // основной цикл
  l := 1;
  for i := 1 to CommaCount do begin
    bufstr := '';
    inc(l);
    repeat
      bufstr := bufstr + MenuCode[l];
      inc(l);
    until (MenuCode[l] = ',') or (l > length(MenuCode));
    OpenObject(XMLStream,'Dish');
    WriteAttribute(XMLStream,'ME_ID',bufstr);
    // извлечение данных о MENU_ENTRY
    with Query do begin
      Close;
      SQL.Clear;
      SQL.Add('select dishes.dish_name, menu_entries.weight, menu_entries.cost, menu_entries.on_weight');
      SQL.Add('from dishes, menu_entries');
      SQL.Add('where (dishes.d_id = menu_entries.d_id) and (menu_entries.me_id = '
        + bufstr + ');');
      Open;
      FetchAll;
    end;
    WriteAttribute(XMLStream,'DISH_NAME',Query.FieldByName('DISH_NAME').AsString);
    WriteAttribute(XMLStream,'WEIGHT',Query.FieldByName('WEIGHT').AsString);
    WriteAttribute(XMLStream,'COST',Query.FieldByName('COST').AsString);
    WriteAttribute(XMLStream,'ON_WEIGHT',Query.FieldByName('ON_WEIGHT').AsString);
    CloseObject(XMLStream);
  end; // for i
  // закрываем объект
  CloseObject(XMLStream);
  // закрываем корень
  CloseObject(XMLStream);
  XMLStream.Free;
end;

// Экспорт меню в текстовый файл
procedure T_vk_object.MenuTxtExport(MA_ID : integer; FileName : string);
var
  i, l, CommaCount : integer;
  TXTStream : TFileStream;
  StrStream : TStringStream;
  TXTFileName, bufstr, MenuCode, ExportBase, ExportTail : string;
begin
  TXTFileName := FileName;
  try
    TXTStream := TFileStream.Create(FileName, fmOpenReadWrite);
    TXTStream.Position := TXTStream.Size;
  except
    TXTStream := TFileStream.Create(FileName, fmCreate);
  end;
  // извлечение чека из БД
  with Query do begin
    Close;
    SQL.Clear;
    SQL.Add('select * from menu_archives where ma_id = ' + IntToStr(MA_ID));
    Open;
  end;
  // формирование начала записи
  ExportBase := StrToLength(Query.FieldByName('Ma_Id').AsString,MA_ID_length)
      + StrToLength(FormatDateTime('dd/mm/yyyy',Query.FieldByName('MENU_DATE').AsVariant),MENU_DATE_length);
  MenuCode := Query.FieldByName('MENU').AsString;
  // разбор кода заказа
  CommaCount := 0;
  StrStream := TStringStream.Create('');
  // подсчёт разделителей
  for i := 1 to length(MenuCode) do begin
    if MenuCode[i] = ',' then inc(CommaCount);
  end;
  // основной цикл
  l := 1;
  for i := 1 to CommaCount do begin
    bufstr := '';
    inc(l);
    repeat
      bufstr := bufstr + MenuCode[l];
      inc(l);
    until (MenuCode[l] = ',') or (l > length(MenuCode));
    // извлечение данных о MENU_ENTRY
    with Query do begin
      Close;
      SQL.Clear;
      SQL.Add('select dishes.dish_name, menu_entries.weight, menu_entries.cost, menu_entries.on_weight');
      SQL.Add('from dishes, menu_entries');
      SQL.Add('where (dishes.d_id = menu_entries.d_id) and (menu_entries.me_id = '
        + bufstr + ');');
      Open;
      FetchAll;
    end;
    // формирование хвоста записи
    ExportTail := StrToLength(Query.FieldByName('COST').AsString,COST_length)
      + StrToLength(Query.FieldByName('DISH_NAME').AsString,DISH_NAME_length)
      + StrToLength(Query.FieldByName('WEIGHT').AsString,WEIGHT_length) + #13#10;
    // помещение полной записи в строковый поток
    StrStream.WriteString(ExportBase + ExportTail);
  end; // for i
  // копирование строкового потока в конец потока файла
  try
    StrStream.Position := 0;
    TXTStream.Position := TXTStream.Size;
    try
      TXTStream.CopyFrom(StrStream,StrStream.Size);
    finally
      TXTStream.Free;
    end;
  finally
    StrStream.Free;
  end;
end;



  /////////////////////////////////////////////////////////////////////
  function T_vk_object.prop1(mode: TMode): String;
  begin
    case mode of
      m_rus_name: Result:='';
      m_eng_name: Result:='';
      m_get_value: g_Value:=g_fname;
      m_set_value: ;
    end;//case
  end;

  (*13*)

  /////////////////////////////////////////////////////////////////////
  // Коннект к базе
  function T_vk_object.meth1(mode: TMode): String;
  var
    s: String;
    ms: Integer;
    Config : TIniFile;
  begin
    case mode of
      m_rus_name: Result := 'Инициализация';
      m_eng_name: Result := 'Init';
      m_n_params: g_NParams := 2; //Количество параметров функции
      m_execute: begin
        DB := TIBDatabase.Create(nil);
        DBt := TIBTransaction.Create(nil);
        Query := TIBQuery.Create(nil);
        DB.DatabaseName := GetParamAsString(0);
        DB.LoginPrompt := false;
        DB.Params.Add('user_name=sysdba');
        DB.Params.Add('PASSWORD=masterkey');
        DB.Params.Add('lc_ctype=win1251');
        DB.DefaultTransaction := DBt;
        Query.Database := DB;
        Query.Transaction := DBt;
        try
          DB.Connected := true;
          DBt.Active := true;
        except
          raise Exception.Create('Невозможно соединиться с базой данных.');
        end;
        Config := TIniFile.Create(GetParamAsString(1));
        C_ID_length := Config.ReadInteger('TxtFileFormat','C_ID',15);
        CHECK_DATE_length := Config.ReadInteger('TxtFileFormat','CHECK_DATE',17);
        TABEL_length := Config.ReadInteger('TxtFileFormat','TABEL',7);
        COST_length := Config.ReadInteger('TxtFileFormat','COST',7);
        DISH_NAME_length := Config.ReadInteger('TxtFileFormat','DISH_NAME',51);
        WEIGHT_length := Config.ReadInteger('TxtFileFormat','WEIGHT',5);
        MA_ID_length := Config.ReadInteger('TxtFileFormat','MA_ID',15);
        MENU_DATE_length := Config.ReadInteger('TxtFileFormat','MENU_DATE',17);
        Config.Free;
      end;
    end;//case
  end;

  /////////////////////////////////////////////////////////////////////
  // Извлечение всех чеков в текстовый файл
  function T_vk_object.meth2(mode: TMode): String;
  var
    FileName : String;
    i : integer;
    tmpQuery : TIBQuery;
  begin
    case mode of
      m_rus_name: Result := 'ИзвлечьВсеЧекиВТекст';
      m_eng_name: Result := 'ImportAllChecksToTXT';
      m_n_params: g_NParams := 1; //Количество параметров функции
      m_execute: begin
        with tmpQuery do begin
          FileName := GetParamAsString(0);
          tmpQuery := TIBQuery.Create(nil);
          Database := DB;
          Transaction := DBt;
          Close;
          SQL.Clear;
          SQL.Add('select c_id from checks order by checks.check_date;');
          Open;
          FetchAll;
          First;
          for i := 1 to RecordCount - 1 do begin
            CheckTxtExport(FieldByName('C_ID').AsInteger,FileName);
            Next;
          end;
          CheckTxtExport(FieldByName('C_ID').AsInteger,FileName);
          Free;
        end;
      end;
    end;//case
  end;

  /////////////////////////////////////////////////////////////////////
  // Извлечение всех чеков в файл XML
  function T_vk_object.meth3(mode: TMode): String;
  var
    FileName: String;
    i : integer;
    tmpQuery : TIBQuery;
  begin
    case mode of
      m_rus_name: Result := 'ИзвлечьВсеЧекиВXML';
      m_eng_name: Result := 'ImportAllChecksToXML';
      m_n_params: g_NParams := 1; //Количество параметров функции
      m_execute: begin
        with tmpQuery do begin
          FileName := GetParamAsString(0);
          tmpQuery := TIBQuery.Create(nil);
          Database := DB;
          Transaction := DBt;
          Close;
          SQL.Clear;
          SQL.Add('select c_id from checks order by checks.check_date;');
          Open;
          FetchAll;
          First;
          for i := 1 to RecordCount - 1 do begin
            CheckXMLExport(FieldByName('C_ID').AsInteger,FileName);
            Next;
          end;
          CheckXMLExport(FieldByName('C_ID').AsInteger,FileName);
          Free;
        end;
      end;
    end;//case
  end;

  /////////////////////////////////////////////////////////////////////
  // Извлечение всего архива меню в текстовый файл
  function T_vk_object.meth4(mode: TMode): String;
  var
    FileName: String;
    i : integer;
    tmpQuery : TIBQuery;
  begin
    case mode of
      m_rus_name: Result := 'ИзвлечьВсеМенюВТекст';
      m_eng_name: Result := 'ImportAllMenuToTXT';
      m_n_params: g_NParams := 1; //Количество параметров функции
      m_execute: begin
        with tmpQuery do begin
          FileName := GetParamAsString(0);
          tmpQuery := TIBQuery.Create(nil);
          Database := DB;
          Transaction := DBt;
          Close;
          SQL.Clear;
          SQL.Add('select MA_ID from menu_archives order by menu_date;');
          Open;
          FetchAll;
          First;
          for i := 1 to RecordCount - 1 do begin
            MenuTxtExport(FieldByName('MA_ID').AsInteger,FileName);
            Next;
          end;
          MenuTxtExport(FieldByName('MA_ID').AsInteger,FileName);
          Free;
        end;
      end;
    end;//case
  end;

  /////////////////////////////////////////////////////////////////////
  // Извлечение всего архива меню в файл XML
  function T_vk_object.meth5(mode: TMode): String;
  var
    FileName: String;
    i : integer;
    tmpQuery : TIBQuery;
  begin
    case mode of
      m_rus_name: Result := 'ИзвлечьВсеМенюВXML';
      m_eng_name: Result := 'ImportAllMenuToXML';
      m_n_params: g_NParams := 1; //Количество параметров функции
      m_execute: begin
        with tmpQuery do begin
          FileName := GetParamAsString(0);
          tmpQuery := TIBQuery.Create(nil);
          Database := DB;
          Transaction := DBt;
          Close;
          SQL.Clear;
          SQL.Add('select MA_ID from menu_archives order by menu_date;');
          Open;
          FetchAll;
          First;
          for i := 1 to RecordCount - 1 do begin
            MenuXMLExport(FieldByName('MA_ID').AsInteger,FileName);
            Next;
          end;
          MenuXMLExport(FieldByName('MA_ID').AsInteger,FileName);
          Free;
        end;
      end;
    end;//case
  end;

  /////////////////////////////////////////////////////////////////////
  // Извлечение чеков за период времени в текстовый файл
  function T_vk_object.meth6(mode: TMode): String;
  var
    FileName : String;
    i : integer;
    tmpQuery : TIBQuery;
  begin
    case mode of
      m_rus_name: Result := 'ИзвлечьЧекиВТекст';
      m_eng_name: Result := 'ImportChecksToTXT';
      m_n_params: g_NParams := 3; //Количество параметров функции
      m_execute: begin
        with tmpQuery do begin
          FileName := GetParamAsString(0);
          tmpQuery := TIBQuery.Create(nil);
          Database := DB;
          Transaction := DBt;
          Close;
          SQL.Clear;
          SQL.Add('select c_id from checks where check_date between '''
            + GetParamAsString(1) + ''' and ''' + GetParamAsString(2)
            + ''' order by checks.check_date');
          Open;
          FetchAll;
          First;
          for i := 1 to RecordCount - 1 do begin
            CheckTxtExport(FieldByName('C_ID').AsInteger,FileName);
            Next;
          end;
          CheckTxtExport(FieldByName('C_ID').AsInteger,FileName);
          Free;
        end;
      end;
    end;//case
  end;

  /////////////////////////////////////////////////////////////////////
  // Извлечение чеков зв период времени в файл XML
  function T_vk_object.meth7(mode: TMode): String;
  var
    FileName : String;
    i : integer;
    tmpQuery : TIBQuery;
  begin
    case mode of
      m_rus_name: Result := 'ИзвлечьЧекиВXML';
      m_eng_name: Result := 'ImportChecksToXML';
      m_n_params: g_NParams := 3; //Количество параметров функции
      m_execute: begin
        with tmpQuery do begin
          FileName := GetParamAsString(0);
          tmpQuery := TIBQuery.Create(nil);
          Database := DB;
          Transaction := DBt;
          Close;
          SQL.Clear;
          SQL.Add('select c_id from checks where check_date between '''
            + GetParamAsString(1) + ''' and ''' + GetParamAsString(2)
            + ''' order by checks.check_date');
          Open;
          FetchAll;
          First;
          for i := 1 to RecordCount - 1 do begin
            CheckXMLExport(FieldByName('C_ID').AsInteger,FileName);
            Next;
          end;
          CheckXMLExport(FieldByName('C_ID').AsInteger,FileName);
          Free;
        end;
      end;
    end;//case
  end;

  /////////////////////////////////////////////////////////////////////
  // Извлечение меню из архива за период времени в текстовый файл
  function T_vk_object.meth8(mode: TMode): String;
  var
    FileName: String;
    i : integer;
    tmpQuery : TIBQuery;
  begin
    case mode of
      m_rus_name: Result := 'ИзвлечьМенюВТекст';
      m_eng_name: Result := 'ImportMenuToTXT';
      m_n_params: g_NParams := 3; //Количество параметров функции
      m_execute: begin
        with tmpQuery do begin
          FileName := GetParamAsString(0);
          tmpQuery := TIBQuery.Create(nil);
          Database := DB;
          Transaction := DBt;
          Close;
          SQL.Clear;
          SQL.Add('select MA_ID from menu_archives where menu_date between '''
            + GetParamAsString(1) + ''' and ''' + GetParamAsString(2)
            + ''' order by menu_date;');
          Open;
          FetchAll;
          First;
          for i := 1 to RecordCount - 1 do begin
            MenuTxtExport(FieldByName('MA_ID').AsInteger,FileName);
            Next;
          end;
          MenuTxtExport(FieldByName('MA_ID').AsInteger,FileName);
          Free;
        end;
      end;
    end;//case
  end;

  /////////////////////////////////////////////////////////////////////
  // Извлечение меню из архива за период времени в файл XML
  function T_vk_object.meth9(mode: TMode): String;
  var
    FileName: String;
    i : integer;
    tmpQuery : TIBQuery;
  begin
    case mode of
      m_rus_name: Result := 'ИзвлечьМенюВXML';
      m_eng_name: Result := 'ImportMenuToXML';
      m_n_params: g_NParams := 3; //Количество параметров функции
      m_execute: begin
        with tmpQuery do begin
          FileName := GetParamAsString(0);
          tmpQuery := TIBQuery.Create(nil);
          Database := DB;
          Transaction := DBt;
          Close;
          SQL.Clear;
          SQL.Add('select MA_ID from menu_archives where menu_date between '''
            + GetParamAsString(1) + ''' and ''' + GetParamAsString(2)
            + ''' order by menu_date;');
          Open;
          FetchAll;
          First;
          for i := 1 to RecordCount - 1 do begin
            MenuXMLExport(FieldByName('MA_ID').AsInteger,FileName);
            Next;
          end;
          MenuXMLExport(FieldByName('MA_ID').AsInteger,FileName);
          Free;
        end;
      end;
    end;//case
  end;

  (*14*)

////////////////////////////////////////////////////////////////////////
//Функция извлекает параметр из массива g_Params по его индексу
function T_vk_object.GetNParam(lIndex: Integer ): OleVariant;
var
  varGet : OleVariant;
begin
  SafeArrayGetElement(g_Params,lIndex,varGet);
  GetNParam := varGet;
end;

////////////////////////////////////////////////////////////////////////
//Функция извлекает параметр из массива g_Params по его индексу.
//Функция предполагает, что тип значения - строка
function T_vk_object.GetParamAsString(lIndex: Integer ): String;
var
  varGet : OleVariant;
begin
  SafeArrayGetElement(g_Params,lIndex,varGet);
  try
    Result := varGet;
  except
    Raise Exception.Create('Параметр номер '
      + IntToStr(lIndex+1) + ' не может быть преобразован в строку.');
  end;
end;

////////////////////////////////////////////////////////////////////////
//Функция извлекает параметр из массива g_Params по его индексу.
//Функция предполагает, что тип значения - целое число
function T_vk_object.GetParamAsInteger(lIndex: Integer ): Integer;
var
  varGet : OleVariant;
begin
  SafeArrayGetElement(g_Params,lIndex,varGet);
  try
    Result := varGet;
  except
    Raise Exception.Create('Параметр номер '
      + IntToStr(lIndex+1) + ' не может быть преобразован в целое число.');
  end;
end;

////////////////////////////////////////////////////////////////////////
//Функция помещает значение в массив g_Params по указанному индексу
procedure T_vk_object.PutNParam(lIndex: Integer; var varPut: OleVariant);
begin
  SafeArrayPutElement(g_Params,lIndex,varPut);
end;

////////////////////////////////////////////////////////////////////
// Инициализация глобальных переменных
initialization
  CoInitialize(nil);

////////////////////////////////////////////////////////////////////
// Уничтожение глобальных переменных
finalization
  CoUninitialize;
end.
