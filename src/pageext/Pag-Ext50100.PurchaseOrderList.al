pageextension 50100 "Purchase Order List" extends "Purchase Order List"
{
    actions
    {
        addfirst(Processing)
        {
            action(ExportToExcel)
            {
                Caption = 'Export Template';
                ApplicationArea = All;
                Image = ExportToExcel;
                Promoted = true;
                PromotedIsBig = true;
                PromotedCategory = Process;
                PromotedOnly = true;
                trigger OnAction()
                var
                    Export: report "Purchase Order Export";
                    PurchaseOrder: Record "Purchase Header";
                begin
                    PurchaseOrder.SetRange("No.", rec."No.");
                    if PurchaseOrder.Find('-') then
                        Export.SetTableView(PurchaseOrder);
                    Export.Run();
                end;
            }
            action(ImportFromExcel)
            {
                Caption = 'Import Template';
                ApplicationArea = All;
                Image = ImportExcel;
                Promoted = true;
                PromotedIsBig = true;
                PromotedCategory = Process;
                PromotedOnly = true;
                trigger OnAction()
                begin
                    ReadExcelSheet();
                end;
            }
        }
    }
    local procedure ReadExcelSheet()
    var
        FileMgt: Codeunit "File Management";
        IStream: InStream;
        NameValueBufferOut: Record "Name/Value Buffer" temporary;
        FromFile: Text[100];
        MaxRow: Integer;
        ColNo: Integer;
        RowNo: Integer;
    begin
        SheetName := '';
        FileName := '';
        if UploadIntoStream(UploadExcelMsg, '', '', FromFile, IStream) then begin
            if FromFile <> '' then begin
                FileName := FileMgt.GetFileName(FromFile);
                TempExcelBuffer.GetSheetsNameListFromStream(IStream, NameValueBufferOut);
            end else
                Error(NoFileFoundMsg);
            TempExcelBuffer.Reset();
            TempExcelBuffer.DeleteAll();
            NameValueBufferOut.Reset();
            if NameValueBufferOut.FindSet() then begin
                repeat
                    Clear(SheetName);
                    SheetName := NameValueBufferOut.Value;
                    TempExcelBuffer.OpenBookStream(IStream, SheetName);
                    TempExcelBuffer.ReadSheet();

                    MaxRow := 0;
                    TempExcelBuffer.Reset();
                    if TempExcelBuffer.FindLast() then
                        MaxRow := TempExcelBuffer."Row No.";

                    ColNo := 0;
                    RowNo := 0;
                    IF SheetName = PurchaseOrderLineLbl THEN
                        ImportExcelDataLines();
                    IF SheetName = PurchaseOrderHeaderLbl THEN
                        ImportExcelDataHeader();
                    IF SheetName = ReservationLbl then
                        ImportReservationEntry();
                until NameValueBufferOut.Next() = 0;
            end;
        end else begin
            Message('File is not found');
        end;
    end;

    local procedure ReadExcelSheetHeader()
    var
        FileMgt: Codeunit "File Management";
        IStream: InStream;
        FromFile: Text[100];
    begin
        SheetName := '';
        FileName := '';
        UploadIntoStream(UploadExcelMsg, '', '', FromFile, IStream);
        if FromFile <> '' then begin
            FileName := FileMgt.GetFileName(FromFile);
            SheetName := TempExcelBuffer.SelectSheetsNameStream(IStream);
        end else
            Error(NoFileFoundMsg);
        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        TempExcelBuffer.OpenBookStream(IStream, SheetName);
        TempExcelBuffer.ReadSheet();
        SheetName := TempExcelBuffer.SelectSheetsNameStream(Istream);
        IF SheetName = PurchaseOrderHeaderLbl THEN
            ImportExcelDataHeader();
    end;

    local procedure ReadExcelSheetLine()
    var
        FileMgt: Codeunit "File Management";
        IStream: InStream;
        FromFile: Text[100];
    begin
        SheetName := '';
        FileName := '';
        UploadIntoStream(UploadExcelMsg, '', '', FromFile, IStream);
        if FromFile <> '' then begin
            FileName := FileMgt.GetFileName(FromFile);
            SheetName := TempExcelBuffer.SelectSheetsNameStream(IStream);
        end else
            Error(NoFileFoundMsg);
        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        TempExcelBuffer.OpenBookStream(IStream, SheetName);
        TempExcelBuffer.ReadSheet();
        SheetName := TempExcelBuffer.SelectSheetsNameStream(Istream);
        IF SheetName = PurchaseOrderLineLbl THEN
            ImportExcelDataLines();
    end;

    local procedure ImportExcelDataHeader()
    var
        SHImportBuffer: Record "Purchase Header";
        SLImportBuffer: Record "Purchase Line";
        RowNo: Integer;
        ColNo: Integer;
        LineNo: Integer;
        MaxRowNo: Integer;
        Progress: Dialog;
        RecordCounted: Integer;
        Counter: Integer;
        I: Integer;
        InvoiceNo: Code[24];
        UnitofMeasure: Code[20];
        ProgressMsg: Label 'Purchase Order Header Importing\Order No.: #1#####\Record Count: #2#####\Total Records Counted: #3#####';
    begin
        RowNo := 0;
        ColNo := 0;
        MaxRowNo := 0;
        LineNo := 0;
        SHImportBuffer.Reset();
        TempExcelBuffer.Reset();
        if TempExcelBuffer.FindLast() then begin
            MaxRowNo := TempExcelBuffer."Row No.";
        end;
        if TempExcelBuffer.FindSet() then
            RecordCounted := TempExcelBuffer.Count;
        Progress.OPEN(ProgressMsg, SHImportBuffer."No.", Counter, RecordCounted);
        for RowNo := 2 to MaxRowNo do begin
            Counter := Counter + 1;
            LineNo := LineNo + 10000;
            SHImportBuffer.Init();
            Evaluate(SHImportBuffer."Document Type", GetValueAtCell(RowNo, 1));
            Evaluate(SHImportBuffer."No.", GetValueAtCell(RowNo, 2));
            Evaluate(SHImportBuffer."Buy-from Vendor No.", GetValueAtCell(RowNo, 3));
            Evaluate(SHImportBuffer."Buy-from Vendor Name", GetValueAtCell(RowNo, 4));
            Evaluate(SHImportBuffer."Location Code", GetValueAtCell(RowNo, 5));
            Evaluate(SHImportBuffer."Document Date", GetValueAtCell(RowNo, 6));
            Evaluate(SHImportBuffer."Posting Date", GetValueAtCell(RowNo, 7));
            Evaluate(SHImportBuffer."Payment Terms Code", GetValueAtCell(RowNo, 8));
            SHImportBuffer.Insert(true);
            SHImportBuffer.Validate("Buy-from Vendor No.");
            SHImportBuffer.Modify(true);
            Progress.UPDATE();
            SLEEP(50);
        end;
        Progress.CLOSE();
        Message('Total Purchase Header Imported = %1', Counter);
    end;

    local procedure TestingConnection()
    begin
        Message('');
    end;

    local procedure ImportExcelDataLines()
    var
        SLImportBuffer: Record "Purchase Line";
        Location: Record Location;
        RowNo: Integer;
        ColNo: Integer;
        LineNo: Integer;
        MaxRowNo: Integer;
        LineType: Option "Item","G/L Account";
        No: Code[50];
        Description: Text[250];
        DocNo: Code[50];
        UnitPrice: Decimal;
        LineQty: Decimal;
        UOM: Code[24];
        Variant: Code[24];
        LineDiscountAmount: Decimal;
        ShortCutDim1: Code[50];
        ShortCutDim2: Code[50];
        UnitofMeasureCode: Code[20];
        DocumentType: Option "Invoice";
        LocationCode: Code[50];
        GeneralProductPostingGroup: Code[50];
        Progress: Dialog;
        RecordCounted: Integer;
        Counter: Integer;
        I: Integer;
        InvoiceNo: Code[24];
        ProgressMsg: Label 'Importing Purchase Lines\Order No.: #1#####\No.: #2#####\Record Count: #3#####';

    begin
        RowNo := 0;
        ColNo := 0;
        MaxRowNo := 0;
        LineNo := 0;
        SLImportBuffer.Reset();
        TempExcelBuffer.Reset();
        if TempExcelBuffer.FindLast() then begin
            MaxRowNo := TempExcelBuffer."Row No.";
        end;
        if TempExcelBuffer.FindSet() then
            RecordCounted := TempExcelBuffer.Count;
        Progress.OPEN(ProgressMsg, SLImportBuffer."Document No.", SLImportBuffer."No.", Counter);
        for RowNo := 2 to MaxRowNo do begin

            Counter := Counter + 1;
            DocNo := GetValueAtCell(RowNo, 3);
            SLImportBuffer.SetRange("Document No.", DocNo);
            Evaluate(LocationCode, GetValueAtCell(RowNo, 6));
            Evaluate(LineQty, GetValueAtCell(RowNo, 8));
            if SLImportBuffer.FindLast() then
                LineNo := SLImportBuffer."Line No." + 10000
            else
                LineNo := 10000;
            SLImportBuffer.Init();
            SLImportBuffer."Line No." := LineNo;
            Evaluate(SLImportBuffer."Type", GetValueAtCell(RowNo, 1));
            Evaluate(SLImportBuffer."Document Type", GetValueAtCell(RowNo, 2));
            Evaluate(SLImportBuffer."Document No.", DocNo, 3);
            Evaluate(SLImportBuffer."No.", GetValueAtCell(RowNo, 4));
            Evaluate(SLImportBuffer."Description", GetValueAtCell(RowNo, 5));
            Evaluate(SLImportBuffer."Location Code", GetValueAtCell(RowNo, 6));
            Evaluate(SLImportBuffer."Bin Code", GetValueAtCell(RowNo, 7));
            Evaluate(SLImportBuffer."Quantity", GetValueAtCell(RowNo, 8));
            Evaluate(SLImportBuffer."Unit of Measure Code", GetValueAtCell(RowNo, 9));
            Evaluate(SLImportBuffer."Promised Receipt Date", GetValueAtCell(RowNo, 10));
            Evaluate(SLImportBuffer."Planned Receipt Date", GetValueAtCell(RowNo, 11));
            Evaluate(SLImportBuffer."Expected Receipt Date", GetValueAtCell(RowNo, 12));
            SLImportBuffer.Insert();
            SLImportBuffer.Validate("No.", GetValueAtCell(RowNo, 4));
            SLImportBuffer.Validate(Description, GetValueAtCell(RowNo, 5));
            SLImportBuffer.validate("Location Code", GetValueAtCell(RowNo, 6));
            SLImportBuffer.validate(Quantity, LineQty);
            SLImportBuffer.validate("Unit of Measure Code", GetValueAtCell(RowNo, 9));
            SLImportBuffer.Modify(true);
            Progress.UPDATE();
            SLEEP(50);
        end;
        Progress.CLOSE();
        Message('Total Purchase Order Lines Imported = %1', Counter);
    end;

    local procedure ImportReservationEntry()
    var
        PurchaseLine: Record "Purchase Line";
        ReservationEntry: Record "Reservation Entry";
        RowNo: Integer;
        ColNo: Integer;
        ItemNo: code[20];
        Qty: Decimal;
        QtyBase: Decimal;
        QtyToInvoiceBase: Decimal;
        SerialNo: Code[50];
        QtyToHandleBase: Decimal;
        ApplFromItemEntry: Integer;
        LineNo: Integer;
        ExpirationDate: Date;
        LotNo: Code[50];
        MaxRowNo: Integer;
        Progress: Dialog;
        RecordCounted: Integer;
        Counter: Integer;
        LastEntryNo: Integer;
        DocNo: Code[20];
        I: Integer;
        LocationCodeRe: code[20];
        ProgressMsg: Label 'Importing Reservation Entry\No.: #1#####\Record Count: #2#####';
    begin
        RowNo := 0;
        ColNo := 0;
        MaxRowNo := 0;
        LineNo := 0;
        TempExcelBuffer.Reset();
        if TempExcelBuffer.FindLast() then begin
            MaxRowNo := TempExcelBuffer."Row No.";
        end;
        if TempExcelBuffer.FindSet() then
            RecordCounted := TempExcelBuffer.Count;
        Progress.OPEN(ProgressMsg, ReservationEntry."Lot No.", Counter, RecordCounted);
        for RowNo := 2 to MaxRowNo do begin
            LineNo := LineNo + 10000;
            Evaluate(DocNo, GetValueAtCell(RowNo, 1));
            Evaluate(LineNo, GetValueAtCell(RowNo, 2));
            Evaluate(ItemNo, GetValueAtCell(RowNo, 3));
            Evaluate(LotNo, GetValueAtCell(RowNo, 4));
            Evaluate(QtyBase, GetValueAtCell(RowNo, 5));
            Evaluate(ExpirationDate, GetValueAtCell(RowNo, 6));
            evaluate(LocationCodeRe, GetValueAtCell(RowNo, 7));

            Counter := Counter + 1;
            ReservationEntry.Reset();
            ReservationEntry.Init();
            ReservationEntry."Entry No." := NextEntryNo;
            ReservationEntry."Item No." := ItemNo;
            ReservationEntry."Reservation Status" := ReservationEntry."Reservation Status"::Surplus;
            ReservationEntry."Source Batch Name" := '';
            ReservationEntry."Source Type" := Database::"Purchase Line";
            ReservationEntry."Source Subtype" := ReservationEntry."Source Subtype"::"1";
            ReservationEntry."Source ID" := DocNo;
            ReservationEntry."Source Ref. No." := LineNo;
            ReservationEntry."Quantity (Base)" := QtyBase;
            ReservationEntry."Lot No." := LotNo;
            ReservationEntry."Expiration Date" := ExpirationDate;
            ReservationEntry."Location Code" := LocationCodeRe;
            ReservationEntry."Item Tracking" := ReservationEntry."Item Tracking"::"Lot No.";
            ReservationEntry."Created By" := UserId;
            ReservationEntry.Positive := true;
            ReservationEntry."Creation Date" := WorkDate();
            ReservationEntry.Insert(true);
            ReservationEntry."Source ID" := DocNo;
            ReservationEntry."Source Ref. No." := LineNo;
            ReservationEntry."Lot No." := LotNo;
            ReservationEntry."Quantity (Base)" := QtyBase;
            ReservationEntry."Expiration Date" := ExpirationDate;
            ReservationEntry.Validate("Item No.", ItemNo);
            ReservationEntry.Validate("Lot No.", LotNo);
            ReservationEntry.Validate("Quantity (Base)", QtyBase);
            ReservationEntry.validate("Location Code", LocationCodeRe);
            ReservationEntry.Modify(true);
            Progress.UPDATE();
        end;
        Progress.CLOSE();
        Message('Lots Imported = %1', Counter);
    end;

    local procedure NextEntryNo(): Integer
    var
        ReserveEntry: Record "Reservation Entry";
    begin
        ReserveEntry.Reset();
        if LastEntryNo = 0 then
            if ReserveEntry.FindLast() then
                LastEntryNo := ReserveEntry."Entry No.";
        LastEntryNo += 1;
        exit(LastEntryNo);
    end;

    local procedure GetValueAtCell(RowNo: Integer; ColNo: Integer): Text
    begin
        TempExcelBuffer.Reset();
        If TempExcelBuffer.Get(RowNo, ColNo) then
            exit(TempExcelBuffer."Cell Value as Text")
        else
            exit('');
    end;

    local procedure UpdateSellToCustomerNo(DocNo: Code[50]): Code[24]
    var
        PurchaseOrderHeader: Record "Purchase Line";
    begin
        PurchaseOrderHeader.Reset();
        PurchaseOrderHeader.SetRange("No.", DocNo);
        if PurchaseOrderHeader.Find('-') then
            exit(PurchaseOrderHeader."Buy-from Vendor No.");
    end;

    local procedure UpdateGenBusPostingGrp(CustomerNo: Code[50]): Code[50]
    var
        Cus: Record Customer;
    begin
        Cus.SetRange("No.", CustomerNo);
        if Cus.Find('-') then
            exit(Cus."Gen. Bus. Posting Group");
    end;

    local procedure VATBusPostingGrp(CustomerNo: Code[50]): Code[50]
    var
        Cus: Record Customer;
    begin
        Cus.SetRange("No.", CustomerNo);
        if Cus.Find('-') then
            exit(Cus."VAT Bus. Posting Group");
    end;

    var
        ImportVisible: Boolean;
        AccessControl: Record "Access Control";
        FileName: Text[100];
        SheetName: Text[100];
        TempExcelBuffer: Record "Excel Buffer" temporary;
        UploadExcelMsg: Label 'Please Choose the Excel file.';
        NoFileFoundMsg: Label 'No Excel file found!';
        BatchISBlankMsg: Label 'Batch name is blank';
        ExcelHeaderImportSucess: Label 'Purchase Order Header with document type Order is successfully imported.';
        ExcelLineImportSucess: Label 'Purchase Order Line with document type Order is successfully imported.';
        PurchaseOrderLineLbl: Label 'Purchase Order Line';
        PurchaseOrderHeaderLbl: Label 'Purchase Header';
        ReservationLbl: Label 'Lot Info';
        LastEntryNo: Integer;
}