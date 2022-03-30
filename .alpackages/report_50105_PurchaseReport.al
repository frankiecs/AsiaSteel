report 50105 PurchaseReport
{
    Caption = 'Purchase Report';
    UsageCategory = Administration;
    ApplicationArea = All;
    ProcessingOnly = true;
    UseRequestPage = true;


    dataset
    {
        dataitem(Vendor1; Vendor)
        {
            trigger OnPreDataItem()
            begin
                ExcelBuffer.NewRow();
                ExcelBuffer.NewRow();
                ExcelBuffer.NewRow();
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn('Purchase ' + format(intYear), false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn('NAME OF SUPPLIER', false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.AddColumn('JAN-' + textYear2, false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.AddColumn('FEB-' + textYear2, false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.AddColumn('MAR-' + textYear2, false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.AddColumn('APR-' + textYear2, false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.AddColumn('MAY-' + textYear2, false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.AddColumn('JUN-' + textYear2, false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.AddColumn('JUL-' + textYear2, false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.AddColumn('AUG-' + textYear2, false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.AddColumn('SEP-' + textYear2, false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.AddColumn('OCT-' + textYear2, false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.AddColumn('NOV-' + textYear2, false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.AddColumn('DEC-' + textYear2, false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.AddColumn(format(intYear) + ' TOTAL PURCHASE', false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.AddColumn('%', false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
            end;

            trigger OnAfterGetRecord()
            begin
                Clear(OneVendPurchAmt);
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn(Vendor1.Name, false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.AddColumn(-VendPurchAmt(Vendor1."No.", JANFirstDate, JANLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-VendPurchAmt(Vendor1."No.", FEBFirstDate, FEBLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-VendPurchAmt(Vendor1."No.", MARFirstDate, MARLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);

                ExcelBuffer.AddColumn(-VendPurchAmt(Vendor1."No.", APRFirstDate, APRLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-VendPurchAmt(Vendor1."No.", MAYFirstDate, MAYLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-VendPurchAmt(Vendor1."No.", JUNFirstDate, JUNLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);

                ExcelBuffer.AddColumn(-VendPurchAmt(Vendor1."No.", JULFirstDate, JULLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-VendPurchAmt(Vendor1."No.", AUGFirstDate, AUGLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-VendPurchAmt(Vendor1."No.", SEPFirstDate, SEPLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);

                ExcelBuffer.AddColumn(-VendPurchAmt(Vendor1."No.", OCTFirstDate, OCTLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-VendPurchAmt(Vendor1."No.", NOVFirstDate, NOVLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-VendPurchAmt(Vendor1."No.", DECFirstDate, DECLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);

                ExcelBuffer.AddColumn(-OneVendPurchAmt, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                IF SumAllPurchAmt <> 0 then
                    ExcelBuffer.AddColumn(OneVendPurchAmt / SumAllPurchAmt * 100, false, '', false, false, false, '#0.00', ExcelBuffer."Cell Type"::Number);

            end;


            trigger OnPostDataItem()
            begin
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn('TOTAL', false, '', true, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);

                ExcelBuffer.AddColumn(-JANTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-FEBTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-MARTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-APRTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-MAYTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-JUNTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-JULTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-AUGTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-SEPTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-OCTTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-NOVTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-DECTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(-SumAllPurchAmt, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                IF SumAllPurchAmt <> 0 then
                    ExcelBuffer.AddColumn(-SumAllPurchAmt / -SumAllPurchAmt * 100, false, '', false, false, false, '#0.00', ExcelBuffer."Cell Type"::Number)
                else
                    ExcelBuffer.AddColumn(1, false, '', false, false, false, '#0.00', ExcelBuffer."Cell Type"::Number);

                //Verify Date Range
                /*
                ExcelBuffer.NewRow();
                ExcelBuffer.NewRow();
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn(format(JANFirstDate) + '..' + Format(JANLastDate), false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn(format(FEBFirstDate) + '..' + Format(FEBLastDate), false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn(format(MARFirstDate) + '..' + Format(MARLastDate), false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn(format(APRFirstDate) + '..' + Format(APRLastDate), false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn(format(MAYFirstDate) + '..' + Format(MAYLastDate), false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn(format(JUNFirstDate) + '..' + Format(JUNLastDate), false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn(format(JULFirstDate) + '..' + Format(JULLastDate), false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn(format(AUGFirstDate) + '..' + Format(AUGLastDate), false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn(format(SEPFirstDate) + '..' + Format(SEPLastDate), false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn(format(OCTFirstDate) + '..' + Format(OCTLastDate), false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn(format(NOVFirstDate) + '..' + Format(NOVLastDate), false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn(format(DECFirstDate) + '..' + Format(DECLastDate), false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.NewRow();
                */
            end;

        }
    }

    requestpage
    {
        layout
        {
            area(Content)
            {
                group(GroupName)
                {
                    field(SelectDate; SelectDate)
                    {
                        ApplicationArea = All;
                        Caption = 'Select a date within required year.';
                    }
                }
            }
        }
    }
    /*
        actions
        {
            area(processing)
            {
                action(ActionName)
                {
                    ApplicationArea = All;
                    
                }
            }
        }
    }
    */


    trigger OnInitReport()
    begin
        ExcelBuffer.reset;
        ExcelBuffer.DeleteAll();
        if recGLSetup.FindFirst() then;
    end;


    trigger OnPreReport()
    begin
        if SelectDate <> 0D then begin
            intYear := DATE2DMY(SelectDate, 3);
            textYear2 := copystr(Format(intYear), 3, 2);

            JANFirstDate := DMY2Date(1, 1, intYear);
            JANLastDate := CALCDATE('CM', JANFirstDate);

            FEBFirstDate := DMY2Date(1, 2, intYear);
            FEBLastDate := CALCDATE('CM', FEBFirstDate);

            MARFirstDate := DMY2Date(1, 3, intYear);
            MARLastDate := CALCDATE('CM', MARFirstDate);

            APRFirstDate := DMY2Date(1, 4, intYear);
            APRLastDate := CALCDATE('CM', APRFirstDate);

            MAYFirstDate := DMY2Date(1, 5, intYear);
            MAYLastDate := CALCDATE('CM', MAYFirstDate);

            JUNFirstDate := DMY2Date(1, 6, intYear);
            JUNLastDate := CALCDATE('CM', JUNFirstDate);

            JULFirstDate := DMY2Date(1, 7, intYear);
            JULLastDate := CALCDATE('CM', JULFirstDate);

            AUGFirstDate := DMY2Date(1, 8, intYear);
            AUGLastDate := CALCDATE('CM', AUGFirstDate);

            SEPFirstDate := DMY2Date(1, 9, intYear);
            SEPLastDate := CALCDATE('CM', SEPFirstDate);

            OCTFirstDate := DMY2Date(1, 10, intYear);
            OCTLastDate := CALCDATE('CM', OCTFirstDate);

            NOVFirstDate := DMY2Date(1, 11, intYear);
            NOVLastDate := CALCDATE('CM', NOVFirstDate);

            DECFirstDate := DMY2Date(1, 12, intYear);
            DECLastDate := CALCDATE('CM', DECFirstDate);

            SumAllPurchAmt := SumPurchAmt(JANFirstDate, DECLastDate)

        end else begin
            Error('A date must be select to define the week.')
        end;
    end;


    trigger OnPostReport()
    begin
        ExcelBuffer.CreateNewBook('Purchase Report');
        ExcelBuffer.WriteSheet('Purchase Report', CompanyName, UserId);
        ExcelBuffer.CloseBook();
        excelbuffer.OpenExcel();
    end;


    var
        ExcelBuffer: Record "Excel Buffer" temporary;
        recGLSetup: Record "General Ledger Setup";
        SelectDate: date;
        TempDate: Date;
        intYear: Integer;
        textYear2: text[2];
        JANFirstDate: Date;
        JANLastDate: Date;
        FEBFIrstDate: Date;
        FEBLastDate: Date;
        MARFirstDate: Date;
        MARLastDate: Date;
        APRFirstDate: Date;
        APRLastDate: Date;
        MAYFirstDate: Date;
        MAYLastDate: Date;
        JUNFirstDate: Date;
        JUNLastDate: Date;
        JULFirstDate: Date;
        JULLastDate: Date;
        AUGFirstDate: Date;
        AUGLastDate: Date;
        SEPFirstDate: Date;
        SEPLastDate: Date;
        OCTFirstDate: Date;
        OCTLastDate: Date;
        NOVFirstDate: Date;
        NOVLastDate: Date;
        DECFirstDate: Date;
        DECLastDate: Date;
        SumAllPurchAmt: Decimal;
        SumVendPurchAmt: Decimal;
        OneVendPurchAmt: Decimal;
        JANTotal: Decimal;
        FEBTotal: Decimal;
        MARTotal: Decimal;
        APRTotal: Decimal;
        MAYTotal: Decimal;
        JUNTotal: Decimal;
        JULTotal: Decimal;
        AUGTotal: Decimal;
        SEPTotal: Decimal;
        OCTTotal: Decimal;
        NOVTotal: Decimal;
        DECTotal: Decimal;

    local procedure VendPurchAmt(VendNo: Code[20]; FromDate: Date; ToDate: Date): Decimal
    var
        recVLE: Record "Vendor Ledger Entry";
        SumAmt: Decimal;
    begin
        recVLE.Reset();
        recVLE.SetRange("Vendor No.", VendNo);
        recVLE.SetFilter("Posting Date", '%1..%2', FromDate, ToDate);
        recVLE.SetFilter("Document Type", '%1|%2', recVLE."Document Type"::Invoice, recVLE."Document Type"::"Credit Memo");
        IF recVLE.FindSet() then begin
            repeat
                if recVLE.CalcFields("Amount (LCY)") then SumAmt += recVLE."Amount (LCY)";
            until recVLE.Next() = 0;
            OneVendPurchAmt += SumAmt;

            case Date2DMY(FromDate, 2) of
                1:
                    JANTotal += SumAmt;
                2:
                    FEBTotal += SumAmt;
                3:
                    MARTotal += SumAmt;
                4:
                    APRTotal += SumAmt;
                5:
                    MAYTotal += SumAmt;
                6:
                    JUNTotal += SumAmt;
                7:
                    JULTotal += SumAmt;
                8:
                    AUGTotal += SumAmt;
                9:
                    SEPTotal += SumAmt;
                10:
                    OCTTotal += SumAmt;
                11:
                    NOVTotal += SumAmt;
                12:
                    DECTotal += SumAmt;
            end;

            exit(SumAmt);
        end else begin
            exit(0);
        end;
    end;

    local procedure SumPurchAmt(FromDate: Date; ToDate: Date): Decimal
    var
        recVLE: Record "Vendor Ledger Entry";
        SumAmt: Decimal;
    begin
        recVLE.Reset();
        recVLE.SetFilter("Posting Date", '%1..%2', FromDate, ToDate);
        recVLE.SetFilter("Document Type", '%1|%2', recVLE."Document Type"::Invoice, recVLE."Document Type"::"Credit Memo");
        IF recVLE.FindSet() then begin
            repeat
                if recVLE.CalcFields("Amount (LCY)") then SumAmt += recVLE."Amount (LCY)";
            until recVLE.Next() = 0;
            exit(SumAmt);
        end else begin
            exit(0);
        end;
    end;


}