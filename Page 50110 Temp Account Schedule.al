page 50110 "Temp Account Schedule"
{
    PageType = list;
    ApplicationArea = All;
    UsageCategory = Administration;
    SourceTable = 50105;

    layout
    {
        area(Content)
        {
            repeater(GroupName)
            {
                field("Line No."; "Line No.")
                {
                    ApplicationArea = All;

                }
                field("G/L Account"; "G/L Account") { ApplicationArea = All; }
                field("Shipment Method"; "Shipment Method") { ApplicationArea = All; }
                field("PN No."; "PN No.") { ApplicationArea = All; }
                field(Amount; Amount) { ApplicationArea = All; }
            }
        }
    }

    actions
    {
        area(Processing)
        {
            action(ActionName)
            {
                ApplicationArea = All;

                trigger OnAction()
                begin

                end;
            }
        }
    }

    var
        myInt: Integer;
}