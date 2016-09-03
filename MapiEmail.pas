unit MapiEmail;

interface

uses
  SysUtils, Classes, JclMapi;

type
  TPrerequisites = class
  public
    function IsMapiAvailable: Boolean;
    function IsClientAvailable: Boolean;
  end;

  TMAPISendMail = class
  private
    FAJclEmail: TJclEmail;
    FShowDialog: Boolean;
    FResolveNames: Boolean;
    FPrerequisites: TPrerequisites;
    // proxy property getters
    function GetMailBody: string;
    function GetHTMLBody: Boolean;
    function GetMailSubject: string;
    // proxy property setters
    procedure SetMailBody(const Value: string);
    procedure SetHTMLBody(const Value: Boolean);
    procedure SetMailSubject(const Value: string);
  protected
    function DoSendMail: Boolean; virtual;
  public
    constructor Create;
    destructor Destroy; override;
    // properties of the wrapper class
    property MailBody: string read GetMailBody write SetMailBody;
    property HTMLBody: Boolean read GetHTMLBody write SetHTMLBody;
    property ShowDialog: Boolean read FShowDialog write FShowDialog;
    property MailSubject: string read GetMailSubject write SetMailSubject;
    property ResolveNames: Boolean read FResolveNames write FResolveNames;
    property Prerequisites: TPrerequisites read FPrerequisites;
    // procedure and functions of the wrapper class
    procedure AddRecipient(const Address: string; const Name: string = '');
    procedure AddAttachment(const FileName: string);
    function SendMail: Boolean;
  end;

implementation

{ TMAPISendMail }

constructor TMAPISendMail.Create;
begin
  FPrerequisites := TPrerequisites.Create;
  FAJclEmail := TJclEmail.Create;
  FShowDialog := True;
end;

destructor TMAPISendMail.Destroy;
begin
  FreeAndNil(FAJclEmail);
  FreeAndNil(FPrerequisites);

  inherited;
end;

function TMAPISendMail.DoSendMail: Boolean;
begin
  Result := FAJclEmail.Send(FShowDialog);
end;

function TMAPISendMail.SendMail: Boolean;
begin
  Result := DoSendMail;
end;

function TMAPISendMail.GetMailBody: string;
begin
  Result := FAJclEmail.Body;
end;

procedure TMAPISendMail.SetMailBody(const Value: string);
begin
  FAJclEmail.Body := Value;
end;

procedure TMAPISendMail.AddAttachment(const FileName: string);
begin
  FAJclEmail.Attachments.Add(FileName);
end;

procedure TMAPISendMail.AddRecipient(const Address, Name: string);
var
  LocalName: string;
  LocalAddress: string;
begin
  LocalAddress := Address;
  LocalName := Name;

  if FResolveNames then
    if not FAJclEmail.ResolveName(LocalName, LocalAddress) then
      raise Exception.Create('Could not resolve Recipient name and address!');

  FAJclEmail.Recipients.Add(LocalAddress, LocalName);
end;

function TMAPISendMail.GetMailSubject: string;
begin
  Result := FAJclEmail.Subject;
end;

procedure TMAPISendMail.SetMailSubject(const Value: string);
begin
  FAJclEmail.Subject := Value;
end;

function TMAPISendMail.GetHTMLBody: Boolean;
begin
  Result := FAJclEmail.HtmlBody;
end;

procedure TMAPISendMail.SetHTMLBody(const Value: Boolean);
begin
  FAJclEmail.HtmlBody := Value;
end;

{ TPrerequisites }

function TPrerequisites.IsClientAvailable: Boolean;
var
  SimpleMAPI: TJclSimpleMapi;
begin
  SimpleMAPI := TJclSimpleMapi.Create;
  try
    Result := SimpleMAPI.AnyClientInstalled;
  finally
    SimpleMAPI.Free;
  end;
end;

function TPrerequisites.IsMapiAvailable: Boolean;
var
  SimpleMAPI: TJclSimpleMapi;
begin
  SimpleMAPI := TJclSimpleMapi.Create;
  try
    Result := SimpleMAPI.SimpleMapiInstalled;
  finally
    SimpleMAPI.Free;
  end;
end;

end.

