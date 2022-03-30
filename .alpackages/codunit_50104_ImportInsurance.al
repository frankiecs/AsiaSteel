codeunit 50104 "Import Insurance"
{

    procedure ping(): Text
    begin
        exit('Pong');
    end;

    procedure ImportInsur(): Text
    begin
        recTB.Reset();
        if recTB.FindFirst() then begin
            Message(format(recTB.Count));
        end;
    end;

    var
        recTB: Record TextBuffer;
}