var clickEvent;

Office.initialize = function () {
}

function sendEmail(event) {
    clickEvent = event;
    if (Office.context.mailbox.item.itemId === null || Office.context.mailbox.item.itemId == undefined) {
        Office.context.mailbox.item.saveAsync(saveItemCallBack);
    }
    else {
        var soapToGetItemData = getItemDataRequest(Office.context.mailbox.item.itemId);
        Office.context.mailbox.makeEwsRequestAsync(soapToGetItemData, itemDataCallback);
    }
}

function saveItemCallBack(result) {
    var soapToGetItemData = getItemDataRequest(result.value);
    Office.context.mailbox.makeEwsRequestAsync(soapToGetItemData, itemDataCallback);
}

function itemDataCallback(asyncResult) {
    if (asyncResult.error != null) {
        updateAndComplete("EWS Status: " + asyncResult.error.message);
        return;
    }

    var xmlDoc = getXMLDocParser(asyncResult.value);
    var result = $('ResponseCode', xmlDoc)[0].textContent;
    if (result != "NoError") {
        updateAndComplete("EWS Status", "The following error code was recieved: " + result);
        return;
    }

    var attachmentsInfo = buildAttachmentsInfo(xmlDoc);
    Office.context.mailbox.item.loadCustomPropertiesAsync(function (asyncResult) {
        var customProps = asyncResult.value;
        customProps.set("myProp", "value");
        customProps.saveAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                updateAndComplete(asyncResult.error.message);
                return;
            }

            modifyEmailAndSend(attachmentsInfo);
        });
    });
}

function modifyEmailAndSend(attachmentsInfo) {
    Office.context.mailbox.item.body.getAsync("html", { asyncContext: "This is passed to the callback" }, function (result) {
        var newText = result.value + "<br>" + attachmentsInfo;
        Office.context.mailbox.item.body.setAsync(newText, { coercionType: Office.CoercionType.Html }, function (asyncResult) {
            if (asyncResult.status != Office.AsyncResultStatus.Succeeded) {
                statusUpdate("Couldn't modify body");
                return;
            }
            Office.context.mailbox.item.saveAsync(function (result) {
                var soapToGetItemData = getItemDataRequest(result.value);
                Office.context.mailbox.makeEwsRequestAsync(soapToGetItemData, function (asyncResult) {
                    if (asyncResult.error != null) {
                        updateAndComplete("EWS Status: " + asyncResult.error.message);
                        return;
                    }

                    var xmlDoc = getXMLDocParser(asyncResult.value);
                    var changeKey = $('ItemId', xmlDoc)[0].getAttribute("ChangeKey");
                    var soapToSendItem = getSendItemRequest(result.value, changeKey);
                    Office.context.mailbox.makeEwsRequestAsync(soapToSendItem, function (asyncResult) {
                        if (asyncResult.error != null) {
                            statusUpdate("EWS Status: " + asyncResult.error.message);
                            return;
                        }

                        Office.context.mailbox.item.close();
                        clickEvent.completed();
                    });
                });
            });

        });
    });
}