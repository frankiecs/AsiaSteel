codeunit 50100 "ImportFunction"
{
    procedure ping(): Text
    begin
        exit('Pong');
    end;

    [ServiceEnabled]
    //procedure ImportInsurance(var ActionContext: WebServiceActionContext)
    procedure ImportInsurance()
    var
        Tbuffer: record TextBuffer;
    begin
        Tbuffer.Reset();
        Tbuffer.DeleteAll();
    end;
}