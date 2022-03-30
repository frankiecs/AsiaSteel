pageextension 50105 ImportJnl extends "General Journal"
{
    actions
    {
        addlast(processing)
        {
            /*
            action("NVT MoveUp")
            {
                ApplicationArea = All;
                Caption = 'Import Commercial Invoice';
                Image = MoveUp;
                Promoted = true;
                PromotedCategory = Process;
                PromotedOnly = true;

                trigger OnAction()
                begin
                    RunXMLPortImport;
                end;
            }

            action("NVT MoveDown")
            {
                ApplicationArea = All;
                Caption = 'Import Insurance Expense';
                Image = MoveDown;
                Promoted = true;
                PromotedCategory = Process;
                PromotedOnly = true;

                trigger OnAction()
                begin
                    RunXMLPortImport3;
                end;
            }

            action("NVT MoveFormat")
            {
                ApplicationArea = All;
                Caption = 'Import Cloud';
                Image = MoveNegativeLines;
                Promoted = true;
                PromotedCategory = Process;
                PromotedOnly = true;

                trigger OnAction()
                begin
                    RunXMLPortImportCloud;
                end;
            }

            */
        }
    }


    procedure RunXMLPortImport()

    var

        FileInstream: InStream;
        FileInstream2: InStream;
        FileName: Text;
        FromFolder: Text;
        FileInSteamRead: Text;
        xmlGenJnlImport2: XmlPort GenJournalImport2;
        xmlImportScanDoc: XmlPort PurInvImport;
        xmlImportInsurance: XmlPort InsuranceImport;
        xmlImportFreight: XmlPort FreightImport;
        xmlImportInspection: XmlPort InspectionImport;
        xmlImportSupplier: XmlPort SupplierImport;
        xmlAdmExp: XmlPort AdmExp;
        xmlComInv: XmlPort CommInvoiceImport;

    begin

        UploadIntoStream('', '', 'All Files (*.*)|*.*', FileName, FileInstream);

        //Check what kind of document importing using csv buffer


        //CSVBuffer.DeleteAll;
        //CSVBuffer.LoadDataFromStream(FileInstream, ',');
        //IF CSVBuffer.FindSet() then
        //    repeat
        //        Message(format(CSVBuffer."Field No.") + ': ' + format(CSVBuffer.Value));
        //    until CSVBuffer.Next() = 0;

        //CSVBuffer.DeleteAll();

        //FileInstream.ReadText(FileInSteamRead);

        //Go to the kind of document importing

        //xmlGenJnlImport2.SetJournalTempalteBatch(rec."Journal Template Name", rec."Journal Batch Name");
        //xmlGenJnlImport2.SetSource(FileInstream);
        //xmlGenJnlImport2.Import();

        xmlComInv.SetSource(FileInstream);
        xmlComInv.import();

        Message('File Import sucessfully.');

    end;


    procedure RunXMLPortImport3()

    var

        FileInstream: InStream;
        FileInstream2: InStream;
        FileName: Text;
        FromFolder: Text;
        FileInSteamRead: Text;

        xmlImportInsurance: XmlPort InsuranceImport3;

    begin

        UploadIntoStream('', '', 'All Files (*.*)|*.*', FileName, FileInstream);
        xmlImportInsurance.SetSource(FileInstream);
        xmlImportInsurance.import();

        Message('File Import sucessfully.');

    end;

    procedure RunXMLPortImportCloud()

    var

        //FileInstream: InStream;
        //FileInstream2: InStream;
        //FileName: Text;
        //FromFolder: Text;
        //FileInSteamRead: Text;

        pageHeaderBuffer: Page "Header Buffer";

    begin

        //pageHeaderBuffer.ImportInsurance2();
        Message('Invoice created.');

    end;

}









