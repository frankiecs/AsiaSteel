reportextension 50100 AgedAccRecPN extends "Aged Accounts Receivable"
{
    dataset
    {
        Add(TempCustLedgEntryLoop)
        {
            //column(PNNO; CustLedgEntryEndingDate."Global Dimension 1 Code") {}
        }
    }

    requestpage
    {
        // Add changes to the requestpage here
    }
}