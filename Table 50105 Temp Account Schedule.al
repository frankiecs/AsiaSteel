table 50105 "Temp Account Schedule"
{
    DataClassification = ToBeClassified;

    fields
    {
        field(1; "Line No."; Integer) { }
        field(2; "Shipment Method"; Code[20]) { }
        field(3; "PN No."; Code[20]) { }
        field(4; "G/L Account"; Code[20]) { }
        field(5; "Amount"; Decimal) { }


    }

    keys
    {
        key(Key1; "Line No.")
        {
            Clustered = true;
        }
    }


    procedure InitAsiaPNLData(StartDate: Date; Endate: Date)
    var
        l_recGLEntry: Record "G/L Entry";
        l_recTempAccountSchedule: Record "Temp Account Schedule";
    begin
        l_recTempAccountSchedule.DeleteAll();

        l_recGLEntry.Reset();
        l_recGLEntry.SetRange("Posting Date", StartDate, Endate);
        l_recGLEntry.SetFilter("Global Dimension 1 Code", '<>%1', '');
        if l_recGLEntry.FindFirst() then
            repeat
                l_recTempAccountSchedule.Reset();
                l_recTempAccountSchedule.SetFilter("G/L Account", l_recGLEntry."G/L Account No.");
                l_recTempAccountSchedule.SetFilter("PN No.", l_recGLEntry."Global Dimension 1 Code");
                if l_recTempAccountSchedule.FindFirst() then begin
                    l_recTempAccountSchedule.Amount += l_recGLEntry.Amount;
                    l_recTempAccountSchedule.Modify();
                end
                else begin
                    Rec.Init();
                    Rec."Line No." := Rec.Count + 1;
                    Rec."Shipment Method" := 'FOB';
                    Rec."G/L Account" := l_recGLEntry."G/L Account No.";
                    Rec."PN No." := l_recGLEntry."Global Dimension 1 Code";
                    Rec.Amount := l_recGLEntry.Amount;
                    Rec.Insert();
                end;

            until l_recGLEntry.Next() = 0;
    end;


}