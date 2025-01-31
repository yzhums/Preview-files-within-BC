pageextension 50110 DocumentAttachmentDetailsExt extends "Document Attachment Details"
{
    actions
    {
        addfirst(processing)
        {
            action(RunHyperLinkInBC)
            {
                Caption = 'Preview Document';
                ApplicationArea = All;
                Promoted = true;
                PromotedIsBig = true;
                PromotedCategory = Process;
                trigger OnAction()
                var
                    PreviewFiles: Page "Preview Files";
                    URL: Text;
                    SharePointFileID: Text;
                    DocAttach: Record "Document Attachment";
                    PreviewDocument: Codeunit PreviewDocument;
                begin
                    URL := '';
                    SharePointFileID := '';
                    DocAttach.Reset();
                    CurrPage.SetSelectionFilter(DocAttach);
                    if DocAttach.FindFirst() then begin
                        PreviewDocument.UploadFilesToSharePoint(DocAttach, URL, SharePointFileID);
                        PreviewFiles.SetURL(URL);
                        PreviewFiles.Run();
                    end;
                end;
            }
        }
    }
}
page 50111 "Preview Files"
{
    Extensible = false;
    Caption = 'Preview';
    Editable = false;
    PageType = Worksheet;

    layout
    {
        area(content)
        {
            usercontrol(WebPageViewer; WebPageViewer)
            {
                ApplicationArea = All;
                trigger ControlAddInReady(callbackUrl: Text)
                begin
                    CurrPage.WebPageViewer.Navigate(URL);
                end;

                trigger Callback(data: Text)
                begin
                    CurrPage.Close();
                end;
            }
        }
    }
    var
        URL: Text;

    procedure SetURL(NavigateToURL: Text)
    begin
        URL := NavigateToURL;
    end;
}

codeunit 50102 PreviewDocument
{
    procedure UploadFilesToSharePoint(DocAttach: Record "Document Attachment"; var WebURL: Text; var SharePointFileID: Text)
    var
        HttpClient: HttpClient;
        HttpRequestMessage: HttpRequestMessage;
        HttpResponseMessage: HttpResponseMessage;
        Headers: HttpHeaders;
        ContentHeader: HttpHeaders;
        RequestContent: HttpContent;
        JsonResponse: JsonObject;
        AuthToken: SecretText;
        SharePointFileUrl: Text;
        ResponseText: Text;
        JsonToken: JsonToken;
        TempBlob: Codeunit "Temp Blob";
        FileName: Text;
        TenantMedia: Record "Tenant Media";
        OutStream: OutStream;
        FileContent: InStream;
        MimeType: Text;
    begin
        // Get OAuth token
        AuthToken := GetOAuthToken();

        if AuthToken.IsEmpty() then
            Error('Failed to obtain access token.');

        if DocAttach."Document Reference ID".HasValue then begin
            TempBlob.CreateOutStream(OutStream);
            DocAttach."Document Reference ID".ExportStream(OutStream);
            TempBlob.CreateInStream(FileContent);
            FileName := DocAttach."File Name" + '.' + DocAttach."File Extension";
            if TenantMedia.Get(DocAttach."Document Reference ID".MediaId) then
                MimeType := TenantMedia."Mime Type";
        end;
        // Define the SharePoint folder URL

        // application permissions (replace with the actual site-id, drive-id, folder path and file name)
        SharePointFileUrl := 'https://graph.microsoft.com/v1.0/sites/5b3b7cec-cbfe-4893-a638-c18a34c6a394/drives/b!7Hw7W_7Lk0imOMGKNMajlK0n-8Wdev9FmPdhx03j5o95rz4xvtmtTIUW5qUH7Jww/root:/Business Central/' + FileName + ':/content';
        // Initialize the HTTP request
        HttpRequestMessage.SetRequestUri(SharePointFileUrl);
        HttpRequestMessage.Method := 'PUT';
        HttpRequestMessage.GetHeaders(Headers);
        Headers.Add('Authorization', SecretStrSubstNo('Bearer %1', AuthToken));
        RequestContent.GetHeaders(ContentHeader);
        ContentHeader.Clear();
        ContentHeader.Add('Content-Type', MimeType);
        HttpRequestMessage.Content.WriteFrom(FileContent);

        // Send the HTTP request
        if HttpClient.Send(HttpRequestMessage, HttpResponseMessage) then begin
            // Log the status code for debugging
            //Message('HTTP Status Code: %1', HttpResponseMessage.HttpStatusCode());

            if HttpResponseMessage.IsSuccessStatusCode() then begin
                HttpResponseMessage.Content.ReadAs(ResponseText);
                JsonResponse.ReadFrom(ResponseText);

                JsonResponse.Get('id', JsonToken);
                SharePointFileID := JsonToken.AsValue().AsText();

                JsonResponse.Get('webUrl', JsonToken);
                WebURL := JsonToken.AsValue().AsText();
            end else begin
                //Report errors!
                HttpResponseMessage.Content.ReadAs(ResponseText);
                Error('Failed to upload files to SharePoint: %1 %2', HttpResponseMessage.HttpStatusCode(), ResponseText);
            end;
        end else
            Error('Failed to send HTTP request to SharePoint');
    end;

    procedure GetOAuthToken() AuthToken: SecretText
    var
        ClientID: Text;
        ClientSecret: Text;
        TenantID: Text;
        AccessTokenURL: Text;
        OAuth2: Codeunit OAuth2;
        Scopes: List of [Text];
    begin
        ClientID := 'b4fe1687-f1ab-4bfa-b494-0e2236ed50bd';
        ClientSecret := 'huL8Q~edsQZ4pwyxka3f7.WUkoKNcPuqlOXv0bww';
        TenantID := '7e47da45-7f7d-448a-bd3d-1f4aa2ec8f62';
        AccessTokenURL := 'https://login.microsoftonline.com/' + TenantID + '/oauth2/v2.0/token';
        Scopes.Add('https://graph.microsoft.com/.default');
        if not OAuth2.AcquireTokenWithClientCredentials(ClientID, ClientSecret, AccessTokenURL, '', Scopes, AuthToken) then
            Error('Failed to get access token from response\%1', GetLastErrorText());
    end;
}
