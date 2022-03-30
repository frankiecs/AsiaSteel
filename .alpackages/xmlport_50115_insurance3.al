xmlport 50115 InsuranceImport3
{

    Caption = 'Import Insurance';
    UseRequestPage = false;
    DefaultFieldsValidation = true;
    UseDefaultNamespace = true;
    FormatEvaluate = Xml;
    Direction = Both;
    DefaultNamespace = 'urn:microsoft-dynamics/Customer';

    /*
        Direction = Import;
        TextEncoding = UTF8;
        Format = VariableText;
        FieldDelimiter = '"';
        FieldSeparator = ',';
    */
    schema
    {
        textelement(Root)
        {
            tableelement(TBuffer; TextBuffer)
            {
                //MaxOccurs = Once;
                //MinOccurs = Zero;

                fieldelement(LineNo; TBuffer."Line No.") { MinOccurs = Zero; }
                fieldelement(VendorName; TBuffer.F001) { MinOccurs = Zero; }
                fieldelement(VendorRefNo; TBuffer.F002) { MinOccurs = Zero; }
                fieldelement(IssueDate; TBuffer.F003) { MinOccurs = Zero; }
                fieldelement(Subject; TBuffer.F004) { MinOccurs = Zero; }
                fieldelement(TotalAmount; TBuffer.F005) { MinOccurs = Zero; }
                fieldelement(Currency; TBuffer.F006) { MinOccurs = Zero; }
                fieldelement(Remark; TBuffer.F007) { MinOccurs = Zero; }
                fieldelement(DueDate; TBuffer.F008) { MinOccurs = Zero; }
                /*
                                textelement(VendorName) { MinOccurs = Zero; }
                                textelement(VendorRefNo) { MinOccurs = Zero; }
                                textelement(IssueDate) { MinOccurs = Zero; }
                                textelement(Subject) { MinOccurs = Zero; }
                                textelement(TotalAmount) { MinOccurs = Zero; }
                                textelement(Currency) { MinOccurs = Zero; }
                                textelement(Remark) { MinOccurs = Zero; }
                                textelement(DueDate) { MinOccurs = Zero; }
                */
                trigger OnBeforeInsertRecord()
                begin

                end;
            }
        }
    }
    trigger OnPreXmlPort()
    begin
        booFirstline := true;
        if recPurchSetup.GET then
            codDocNo := NoSeriesMgt.GetNextNo(recPurchSetup."Invoice Nos.", Today, true);
        clear(recTextBuffer);
        recTextBuffer.DeleteAll();
    end;

    var
        recTextBuffer: Record TextBuffer;
        recPurchSetup: Record "Purchases & Payables Setup";
        booFirstline: Boolean;
        codDocNo: Code[20];
        NoSeriesMgt: Codeunit NoSeriesManagement;

    procedure GetlastCommentLineNo() PurCommLineNo: integer
    begin
        /*
        Clear(tblPurCommLine2);
        tblPurCommLine2.SetRange("Document Type", tblPurHdr."Document Type"::Invoice);
        tblPurCommLine2.SetRange("No.", tblPurHdr."No.");
        tblPurCommLine.SetCurrentKey("Line No.");
        if tblPurCommLine2.FindLast() then
            PurCommLineNo := tblPurCommLine2."Line No."
        else
            PurCommLineNo := 0;
        */
    end;
}