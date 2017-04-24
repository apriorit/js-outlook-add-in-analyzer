//Ews request to get item info
function getItemDataRequest(itemId) {
    var soapToGetItemData = '<?xml version="1.0" encoding="utf-8"?>' +
                    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
                    '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
                    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
                    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
                    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
                    '  <soap:Header>' +
                    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
                    '  </soap:Header>' +
                    '  <soap:Body>' +
                    '    <GetItem' +
                    '                xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
                    '                xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
                    '      <ItemShape>' +
                    '        <t:BaseShape>IdOnly</t:BaseShape>' +
                    '        <t:AdditionalProperties>' +
                    '            <t:FieldURI FieldURI="item:Attachments" /> ' +
                    '        </t:AdditionalProperties> ' +
                    '      </ItemShape>' +
                    '      <ItemIds>' +
                    '        <t:ItemId Id="' + itemId + '"/>' +
                    '      </ItemIds>' +
                    '    </GetItem>' +
                    '  </soap:Body>' +
                    '</soap:Envelope>';

    return soapToGetItemData;
}

//Ews request to send the modified item
function getSendItemRequest(itemId, changeKey) {
    var soapSendItemRequest = '<?xml version="1.0" encoding="utf-8"?>' +
                            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
                            '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
                            '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
                            '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
                            '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
                            '  <soap:Header>' +
                            '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
                            '  </soap:Header>' +
                            '  <soap:Body> ' +
                            '    <SendItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
                            '              SaveItemToFolder="true"> ' +
                            '      <ItemIds> ' +
                            '        <t:ItemId Id="' + itemId + '" ChangeKey="' + changeKey + '" /> ' +
                            '      </ItemIds> ' +
                            '      <m:SavedItemFolderId>' +
                            '         <t:DistinguishedFolderId Id="sentitems" />' +
                            '      </m:SavedItemFolderId>' +
                            '    </SendItem> ' +
                            '  </soap:Body> ' +
                            '</soap:Envelope> ';
    return soapSendItemRequest;
}