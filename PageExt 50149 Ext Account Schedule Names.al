pageextension 50149 "Ext Account Schedule Names" extends "Account Schedule Names" //50103 is used by another extension
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {
        addafter(Print)
        {
            action("Print Asia Steel P&L")
            {
                ApplicationArea = Basic, Suite;
                Caption = 'Print Asia Steel P&L';
                Ellipsis = true;
                Image = Print;
                Promoted = true;
                PromotedCategory = Category4;
                PromotedIsBig = true;
                Scope = Repeater;


                trigger OnAction()
                begin
                    PrintAsiaSteelPNL();
                end;
            }
        }
    }


    procedure PrintAsiaSteelPNL()
    var
        l_reportAsiaSteelPNL: Report "Asia Steel P&L";
    begin
        l_reportAsiaSteelPNL.SetAccSchedName(Name);
        l_reportAsiaSteelPNL.SetColumnLayoutName("Default Column Layout");
        l_reportAsiaSteelPNL.Run;
    end;
}


