unit AddInObj;

interface

uses  { Какие библиотеки используем }
  ComServ, ComObj, ActiveX, SysUtils, Windows, AddInLib, vk_object;

     (*1*)
     const c_AddinName = 'CacheImp'; //Имя внешней компоненты

////////////////////////////////////////////////////////////////////////
     //Количество свойств
     const c_PropCount = 1;  (*2*)

     //Идентификаторы свойств
     type TProperties = (
       prop1
     );

////////////////////////////////////////////////////////////////////////
    //Количество методов
     const c_MethCount = 9;   (*3*)
    //Идентификаторы методов.
    type TMethods = (
       meth1,
       meth2,
       meth3,
       meth4,
       meth5,
       meth6,
       meth7,
       meth8,
       meth9
       );

////////////////////////////////////////////////////////////////////////
const
//Нажмите Ctrl-Shift-G, чтобы сгенерировать новый уникальный идентификатор CLSID
//внешней компоненты.           ['{86EB71DF-1610-4821-A35A-847C2321753D}']
     (*4*)
     CLSID_AddInObject : TGUID = '{820268D4-E5E1-45FB-8CB4-95B8F93B7C6D}';



////////////////////////////////////////////////////////////////////////
type

  AddInObject = class(TComObject, IDispatch, IInitDone, ILanguageExtender)

  public
    i1cv7: IDispatch;
    iStatus: IStatusLine;
    iExtWindows: IExtWndsSupport;
    iError: IErrorLog;
    iEvent : IAsyncEvent;
  protected


    //Переменная объекта внешней компоненты
    vk_object: T_vk_object;


    { IInitDone implementation }
    function Init(pConnection: IDispatch): HResult; stdcall;
    function Done: HResult; stdcall;
    function GetInfo(var pInfo: PSafeArray): HResult; stdcall;

    { ILanguageExtender implementation }
    function RegisterExtensionAs(var bstrExtensionName: WideString): HResult; stdcall;
    function GetNProps(var plProps: Integer): HResult; stdcall;
    function FindProp(const bstrPropName: WideString; var plPropNum: Integer): HResult; stdcall;
    function GetPropName(lPropNum, lPropAlias: Integer; var pbstrPropName: WideString): HResult; stdcall;
    function GetPropVal(lPropNum: Integer; var pvarPropVal: OleVariant): HResult; stdcall;
    function SetPropVal(lPropNum: Integer; var varPropVal: OleVariant): HResult; stdcall;
    function IsPropReadable(lPropNum: Integer; var pboolPropRead: Integer): HResult; stdcall;
    function IsPropWritable(lPropNum: Integer; var pboolPropWrite: Integer): HResult; stdcall;
    function GetNMethods(var plMethods: Integer): HResult; stdcall;
    function FindMethod(const bstrMethodName: WideString; var plMethodNum: Integer): HResult; stdcall;
    function GetMethodName(lMethodNum, lMethodAlias: Integer; var pbstrMethodName: WideString): HResult; stdcall;
    function GetNParams(lMethodNum: Integer; var plParams: Integer): HResult; stdcall;
    function GetParamDefValue(lMethodNum, lParamNum: Integer; var pvarParamDefValue: OleVariant): HResult; stdcall;
    function HasRetVal(lMethodNum: Integer; var pboolRetValue: Integer): HResult; stdcall;
    function CallAsProc(lMethodNum: Integer; var paParams: PSafeArray): HResult; stdcall;
    function CallAsFunc(lMethodNum: Integer; var pvarRetValue: OleVariant; var paParams: PSafeArray): HResult; stdcall;

    { IDispatch }
    function GetIDsOfNames(const IID: TGUID; Names: Pointer;
      NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; virtual; stdcall;
    function GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; virtual; stdcall;
    function GetTypeInfoCount(out Count: Integer): HResult; virtual; stdcall;
    function Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; virtual; stdcall;

    { IStatusLine }
    function SetStatusLine(const bstrSource: WideString): HResult; safecall;
    function ResetStatusLine(): HResult; safecall;

    procedure ShowErrorLog(fMessage:WideString);



  end;

////////////////////////////////////////////////////////////////////////
implementation

////////////////////////////////////////////////////////////////////////
function AddInObject.GetPropVal(lPropNum: Integer; var pvarPropVal: OleVariant): HResult; stdcall;
//Здесь 1С читает значения свойств
begin
     VarClear(vk_object.g_Value);
     try
       GetPropVal := S_OK;
       case TProperties(lPropNum) of

            prop1: vk_object.prop1(m_get_value);
            (*5*)
            else
              GetPropVal := S_FALSE;
       end;
       pvarPropVal := vk_object.g_Value;
     except
       on E: Exception do begin
         ShowErrorLog('Ошибка чтения свойства: '+E.Message);
         GetPropVal := S_FALSE;
       end;
     end;

end;

////////////////////////////////////////////////////////////////////////
//Здесь 1С устанавливает значения свойств
function AddInObject.SetPropVal(
  lPropNum: Integer; //Номер свойства
  var varPropVal:  OleVariant //Значение, которое 1С хочет установить
  ): HResult; stdcall;
begin
     try
       Result := S_OK;
       vk_object.g_Value:=varPropVal;
       case TProperties(lPropNum) of
            prop1: vk_object.prop1(m_set_value);
            (*6*)
            else
              Result := S_FALSE;
       end;
     except
       on E:Exception do begin
       ShowErrorLog('Ошибка изменения свойства: '+E.Message);
       SetPropVal := S_FALSE;
       end;
     end;
end;



////////////////////////////////////////////////////////////////////////
function AddInObject.CallAsFunc(lMethodNum: Integer; var pvarRetValue: OleVariant; var paParams: PSafeArray): HResult; stdcall;
{Здесь 1С выполняет код внешних функций}
begin
     try
     pvarRetValue:=0;
     vk_object.g_Value:=0;
     vk_object.g_Params:=paParams;
     vk_object.g_Event:='';
     vk_object.g_Event_Data:='';

     case TMethods(lMethodNum) of
          meth1: vk_object.meth1(m_execute);
          meth2: vk_object.meth2(m_execute);
          meth3: vk_object.meth3(m_execute);
          meth4: vk_object.meth4(m_execute);
          meth5: vk_object.meth5(m_execute);
          meth6: vk_object.meth6(m_execute);
          meth7: vk_object.meth7(m_execute);
          meth8: vk_object.meth8(m_execute);
          meth9: vk_object.meth9(m_execute);

          (*7*)


          else begin
               CallAsFunc := S_FALSE;
               Exit;
               end;
          end; //case
          pvarRetValue := vk_object.g_Value;

          if vk_object.g_Event<>'' then begin
            iEvent.ExternalEvent(c_AddinName, vk_object.g_Event, vk_object.g_Event_Data);
          end;

     except
        on E: Exception do begin
          ShowErrorLog(E.Message);
        end;
     end;


          CallAsFunc := S_OK;
end;

////////////////////////////////////////////////////////////////////////
function AddInObject.FindProp(const bstrPropName: WideString; var plPropNum: Integer): HResult; stdcall;
var
  i: Integer;
  var s: String;
begin
     s:= bstrPropName;
          if s= vk_object.prop1(m_rus_name) then i:=Integer(prop1)
     else if s= vk_object.prop1(m_eng_name) then i:=Integer(prop1)

     (*8*)


     else begin
         plPropNum:=-1;
         FindProp := S_FALSE;
         Exit;
     end;
     plPropNum:=i;
     FindProp := S_OK;
end;

////////////////////////////////////////////////////////////////////////
function AddInObject.FindMethod(const bstrMethodName: WideString; var plMethodNum: Integer): HResult; stdcall;
//1С получает номер метода по его имени
var NewMethodName: WideString;
var i:Integer;
var s:String;
begin
  NewMethodName := bstrMethodName;


     s:= bstrMethodName;
          if s= vk_object.meth1(m_rus_name) then i:=Integer(meth1)
     else if s= vk_object.meth1(m_eng_name) then i:=Integer(meth1)

     else if s= vk_object.meth2(m_rus_name) then i:=Integer(meth2)
     else if s= vk_object.meth2(m_eng_name) then i:=Integer(meth2)

     else if s= vk_object.meth3(m_rus_name) then i:=Integer(meth3)
     else if s= vk_object.meth3(m_eng_name) then i:=Integer(meth3)

     else if s= vk_object.meth4(m_rus_name) then i:=Integer(meth4)
     else if s= vk_object.meth4(m_eng_name) then i:=Integer(meth4)

     else if s= vk_object.meth5(m_rus_name) then i:=Integer(meth5)
     else if s= vk_object.meth5(m_eng_name) then i:=Integer(meth5)

     else if s= vk_object.meth6(m_rus_name) then i:=Integer(meth6)
     else if s= vk_object.meth6(m_eng_name) then i:=Integer(meth6)

     else if s= vk_object.meth7(m_rus_name) then i:=Integer(meth7)
     else if s= vk_object.meth7(m_eng_name) then i:=Integer(meth7)

     else if s= vk_object.meth8(m_rus_name) then i:=Integer(meth8)
     else if s= vk_object.meth8(m_eng_name) then i:=Integer(meth8)

     else if s= vk_object.meth9(m_rus_name) then i:=Integer(meth9)
     else if s= vk_object.meth9(m_eng_name) then i:=Integer(meth9)

     (*9*)

     else begin
         plMethodNum:=-1;
         FindMethod := S_FALSE;
         Exit;
     end;

     plMethodNum:=i;
     FindMethod := S_OK;
end;

////////////////////////////////////////////////////////////////////////
function AddInObject.GetNParams(lMethodNum: Integer; var plParams: Integer): HResult; stdcall;
//Здесь 1С узнает количество параметров у методов
begin
     vk_object.g_NParams:=0;
     case TMethods(lMethodNum) of
          meth1: vk_object.meth1(m_n_params);
          meth2: vk_object.meth2(m_n_params);
          meth3: vk_object.meth3(m_n_params);
          meth4: vk_object.meth4(m_n_params);
          meth5: vk_object.meth5(m_n_params);
          meth6: vk_object.meth6(m_n_params);
          meth7: vk_object.meth7(m_n_params);
          meth8: vk_object.meth8(m_n_params);
          meth9: vk_object.meth9(m_n_params);
          (*10*)

     end;
     plParams:=vk_object.g_NParams;
     GetNParams := S_OK;

end;


////////////////////////////////////////////////////////////////////////
function AddInObject.Init(pConnection: IDispatch): HResult; stdcall;
//1С вызывает эту функцию при инициализации (старте) компоненты
begin
  i1cv7:=pConnection;

  iError:=nil;
  pConnection.QueryInterface(IID_IErrorLog,iError);

  iStatus:=nil;
  pConnection.QueryInterface(IID_IStatusLine,iStatus);

  iEvent := nil;
  pConnection.QueryInterface(IID_IAsyncEvent,iEvent);

  iExtWindows:=nil;
  pConnection.QueryInterface(IID_IExtWndsSupport,iExtWindows);

  vk_object:=T_vk_object.Create();

  Init := S_OK;
end;

////////////////////////////////////////////////////////////////////////
function AddInObject.Done: HResult; stdcall;
//1С вызывает эту функцию при завершении работы компоненты
begin
  If ( iStatus <> nil ) then
    iStatus._Release();

  If ( iExtWindows <> nil ) then
    iExtWindows._Release();

  If ( iError <> nil ) then
    iError._Release();

  if (iEvent <> nil) then
    iEvent._Release();

  vk_object.Destroy();

  Done := S_OK;
end;

////////////////////////////////////////////////////////////////////////
function AddInObject.GetInfo(var pInfo: PSafeArray{(OleVariant)}): HResult; stdcall;
var  varInfo : OleVariant;
var i: Integer;
begin
  varInfo := '2000';
  SafeArrayPutElement(pInfo,i,varInfo);

  GetInfo := S_OK;
end;

////////////////////////////////////////////////////////////////////////
function AddInObject.SetStatusLine(const bstrSource: WideString): HResult; safecall;
//Функции для работы со строкой состояния
begin
  SetStatusLine:=S_OK;
end;

////////////////////////////////////////////////////////////////////////
function AddInObject.ResetStatusLine(): HResult; safecall;
begin
  Result := S_OK;
end;

////////////////////////////////////////////////////////////////////////
function AddInObject.RegisterExtensionAs(var bstrExtensionName: WideString): HResult; stdcall;
begin
  bstrExtensionName := c_AddinName;
  RegisterExtensionAs := S_OK;
end;

////////////////////////////////////////////////////////////////////////
function AddInObject.GetNProps(var plProps: Integer): HResult; stdcall;
begin
     plProps := Integer(c_PropCount);
     GetNProps := S_OK;
end;


////////////////////////////////////////////////////////////////////////
function AddInObject.GetPropName(lPropNum, lPropAlias: Integer; var pbstrPropName: WideString): HResult; stdcall;
begin
     pbstrPropName := '';
     GetPropName := S_OK;
end;


////////////////////////////////////////////////////////////////////////
function AddInObject.IsPropReadable(lPropNum: Integer; var pboolPropRead: Integer): HResult; stdcall;
begin
  IsPropReadable := S_OK; //Все свойства можно читать
end;

////////////////////////////////////////////////////////////////////////
function AddInObject.IsPropWritable(lPropNum: Integer; var pboolPropWrite: Integer): HResult; stdcall;
begin
     IsPropWritable := S_OK; //Все свойства можно читать
end;

////////////////////////////////////////////////////////////////////////
function AddInObject.GetNMethods(var plMethods: Integer): HResult; stdcall;
begin
     plMethods := c_MethCount;
     GetNMethods := S_OK;
end;


////////////////////////////////////////////////////////////////////////
function AddInObject.GetMethodName(lMethodNum, lMethodAlias: Integer; var pbstrMethodName: WideString): HResult; stdcall;
begin
     pbstrMethodName := '';
     GetMethodName := S_OK;
end;


////////////////////////////////////////////////////////////////////////
function AddInObject.GetParamDefValue(lMethodNum, lParamNum: Integer; var pvarParamDefValue: OleVariant): HResult; stdcall;
//Позволяет установить значения по умолчанию для параметров.
begin
  { Ther is no default value for any parameter }
  VarClear(pvarParamDefValue);
  GetParamDefValue := S_OK;
end;

////////////////////////////////////////////////////////////////////////
function AddInObject.HasRetVal(lMethodNum: Integer; var pboolRetValue: Integer): HResult; stdcall;
begin
     pboolRetValue := 1; //Все методы возвращают значение
     HasRetVal := S_OK;
end;

////////////////////////////////////////////////////////////////////////
function AddInObject.CallAsProc(lMethodNum: Integer; var paParams: PSafeArray{(OleVariant)}): HResult; stdcall;
begin
    CallAsProc := S_FALSE;
end;



////////////////////////////////////////////////////////////////////////
function AddInObject.GetIDsOfNames(const IID: TGUID; Names: Pointer;
  NameCount, LocaleID: Integer; DispIDs: Pointer): HResult;
begin
  Result := E_NOTIMPL;
end;

////////////////////////////////////////////////////////////////////////
function AddInObject.GetTypeInfo(Index, LocaleID: Integer;
  out TypeInfo): HResult;
begin
  Result := E_NOTIMPL;
end;

////////////////////////////////////////////////////////////////////////
function AddInObject.GetTypeInfoCount(out Count: Integer): HResult;
begin
  Result := E_NOTIMPL;
end;

////////////////////////////////////////////////////////////////////////
function AddInObject.Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
  Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult;
begin
  Result := E_NOTIMPL;
end;
////////////////////////////////////////////////////////////////////////
procedure AddInObject.ShowErrorLog(fMessage:WideString);
//Показ сообщения об ошибке.
var
  ErrInfo: PExcepInfo;
begin
  If Trim(fMessage) = '' then Exit;
  New(ErrInfo);
  ErrInfo^.bstrSource := c_AddinName;
  ErrInfo^.bstrDescription := fMessage;
  ErrInfo^.wCode:=1006;
  ErrInfo^.sCode:=E_FAIL;
  iError.AddError(nil, ErrInfo);
end;
////////////////////////////////////////////////////////////////////////

begin

  ComServer.SetServerName('AddIn');
  TComObjectFactory.Create(ComServer,AddInObject,CLSID_AddInObject,
  c_AddinName,'V7 AddIn 2.0',ciMultiInstance);


end.
