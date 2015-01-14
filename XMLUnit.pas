unit XMLUnit;

interface

uses
  Classes;

const
  ObjClose = '</Object>'#13#10;
  
procedure OpenObject(var Stream : TFileStream; ObjName : string);
procedure CloseObject(var Stream : TFileStream);
procedure WriteAttribute(var Stream : TFileStream; AttrName, AttrValue : string);

implementation

procedure OpenObject(var Stream : TFileStream; ObjName : string);
var
  Buf : string;
  Temp : TStringStream;
begin
  Buf := '<Object Name="' + ObjName + '">' + #13#10;
  Temp := TStringStream.Create(Buf);
  Stream.CopyFrom(Temp, Temp.Size);
  Temp.Free;
end;

procedure CloseObject(var Stream : TFileStream);
var
  Buf : string;
  Temp : TStringStream;
begin
  Buf := ObjClose;
  Temp := TStringStream.Create(Buf);
  Stream.CopyFrom(Temp, Temp.Size);
  Temp.Free;
end;

procedure WriteAttribute(var Stream : TFileStream; AttrName, AttrValue : string);
var
  Buf : string;
  Temp : TStringStream;
begin
  Buf := '<Attribute Name="' + AttrName + '" Value="' + AttrValue + '"/>' + #13#10;
  Temp := TStringStream.Create(Buf);
  Stream.CopyFrom(Temp, Temp.Size);
  Temp.Free;
end;


end.
