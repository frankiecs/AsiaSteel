report 50106 SalesReport
{
    Caption = 'Sales Report';
    UsageCategory = Administration;
    ApplicationArea = All;
    ProcessingOnly = true;
    UseRequestPage = true;


    dataset
    {
        dataitem(Customer1; Customer)
        {
            trigger OnPreDataItem()
            begin
                ExcelBuffer.NewRow();
                ExcelBuffer.NewRow();
                ExcelBuffer.NewRow();
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn('Sales ' + format(intYear), false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn('NAME OF CUSTOMER', false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
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
                ExcelBuffer.AddColumn(format(intYear) + ' TOTAL SALES', false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.AddColumn('%', false, '', true, false, false, '', ExcelBuffer."Cell Type"::Text);
            end;

            trigger OnAfterGetRecord()
            begin
                Clear(OneVendSalesAmt);
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn(Customer1.Name, false, '', false, false, false, '', ExcelBuffer."Cell Type"::Text);
                ExcelBuffer.AddColumn(CustSalesAmt(Customer1."No.", JANFirstDate, JANLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(CustSalesAmt(Customer1."No.", FEBFirstDate, FEBLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(CustSalesAmt(Customer1."No.", MARFirstDate, MARLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);

                ExcelBuffer.AddColumn(CustSalesAmt(Customer1."No.", APRFirstDate, APRLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(CustSalesAmt(Customer1."No.", MAYFirstDate, MAYLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(CustSalesAmt(Customer1."No.", JUNFirstDate, JUNLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);

                ExcelBuffer.AddColumn(CustSalesAmt(Customer1."No.", JULFirstDate, JULLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(CustSalesAmt(Customer1."No.", AUGFirstDate, AUGLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(CustSalesAmt(Customer1."No.", SEPFirstDate, SEPLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);

                ExcelBuffer.AddColumn(CustSalesAmt(Customer1."No.", OCTFirstDate, OCTLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(CustSalesAmt(Customer1."No.", NOVFirstDate, NOVLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(CustSalesAmt(Customer1."No.", DECFirstDate, DECLastDate), false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);

                ExcelBuffer.AddColumn(OneVendSalesAmt, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                IF SumAllSalesAmt <> 0 then
                    ExcelBuffer.AddColumn(OneVendSalesAmt / SumAllSalesAmt * 100, false, '', false, false, false, '#0.00', ExcelBuffer."Cell Type"::Number);

            end;


            trigger OnPostDataItem()
            begin
                ExcelBuffer.NewRow();
                ExcelBuffer.AddColumn('TOTAL', false, '', true, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);

                ExcelBuffer.AddColumn(JANTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(FEBTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(MARTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(APRTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(MAYTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(JUNTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(JULTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(AUGTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(SEPTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(OCTTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(NOVTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(DECTotal, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                ExcelBuffer.AddColumn(SumAllSalesAmt, false, '', false, false, false, '#,##0.00 ;(#,##0.00)', ExcelBuffer."Cell Type"::Number);
                IF SumAllSalesAmt <> 0 then
                    ExcelBuffer.AddColumn(SumAllSalesAmt / SumAllSalesAmt * 100, false, '', false, false, false, '#0.00', ExcelBuffer."Cell Type"::Number)
                else
                    ExcelBuffer.AddColumn(1, false, '', false, false, false, '#0.00', ExcelBuffer."Cell Type"::Number);

                //Verify Date Ranges
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

            SumAllSalesAmt := SumSalesAmt(JANFirstDate, DECLastDate)

        end else begin
            Error('A date must be select to define the week.')
        end;
    end;


    trigger OnPostReport()
    begin
        ExcelBuffer.CreateNewBook('Sales Report');
        ExcelBuffer.WriteSheet('Sales Report', CompanyName, UserId);
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
        SumAllSalesAmt: Decimal;
        SumVendSalesAmt: Decimal;
        OneVendSalesAmt: Decimal;
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

    local procedure CustSalesAmt(CustNo: Code[20]; FromDate: Date; ToDate: Date): Decimal
    var
        recCLE: Record "Cust. Ledger Entry";
        SumAmt: Decimal;
    begin
        recCLE.Reset();
        recCLE.SetRange("Customer No.", CustNo);
        recCLE.SetFilter("Posting Date", '%1..%2', FromDate, ToDate);
        recCLE.SetFilter("Document Type", '%1|%2', recCLE."Document Type"::Invoice, recCLE."Document Type"::"Credit Memo");
        IF recCLE.FindSet() then begin
            repeat
                if recCLE.CalcFields("Amount (LCY)") then SumAmt += recCLE."Amount (LCY)";
            until recCLE.Next() = 0;
            OneVendSalesAmt += SumAmt;

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

    local procedure SumSalesAmt(FromDate: Date; ToDate: Date): Decimal
    var
        recCLE: Record "Cust. Ledger Entry";
        SumAmt: Decimal;
    begin
        recCLE.Reset();
        recCLE.SetFilter("Posting Date", '%1..%2', FromDate, ToDate);
        recCLE.SetFilter("Document Type", '%1|%2', recCLE."Document Type"::Invoice, recCLE."Document Type"::"Credit Memo");
        IF recCLE.FindSet() then begin
            repeat
                if recCLE.CalcFields("Amount (LCY)") then SumAmt += recCLE."Amount (LCY)";
            until recCLE.Next() = 0;
            exit(SumAmt);
        end else begin
            exit(0);
        end;
    end;


}