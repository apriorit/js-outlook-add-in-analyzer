function buildAttachmentsInfo(xmlDoc) {
    var attachmentsInfo = "You have no any attachments.";
    if ($('HasAttachments', xmlDoc).length == 0) {
        return attachmentsInfo;
    }

    var attachSeparator = "--------------------------------------------- <br>";
    if ($('HasAttachments', xmlDoc)[0].textContent == "true") {
        attachmentsInfo = "";
        var childNodes = $('Attachments', xmlDoc)[0].childNodes;
        childNodes.forEach(function (fileAttachmentItem, fileAttachmentIndex) {
            fileAttachmentItem.childNodes.forEach(function (item, index) {
                if (item.tagName.includes("AttachmentId")) {
                    attachmentsInfo = attachmentsInfo.concat(item.tagName.replace("t:", "") + ': ' + item.getAttribute("Id") + "<br>");
                    return;
                }

                attachmentsInfo = attachmentsInfo.concat(item.tagName.replace("t:", "") + ': ' + item.textContent + "<br>");
            });

            attachmentsInfo = attachmentsInfo.concat(attachSeparator);
        });
    }

    return attachmentsInfo;
}

function updateAndComplete(text) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
        type: "informationalMessage",
        message: text,
        icon: "default_16",
        persistent: false
    });

    clickEvent.completed();
}

function getXMLDocParser(response)
{
    var xmlDoc;
    if (window.DOMParser) {
        var parser = new DOMParser();
        xmlDoc = parser.parseFromString(response, "text/xml");
    }
    else // Older Versions of Internet Explorer
    {
        xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
        xmlDoc.async = false;
        xmlDoc.loadXML(response);
    }
    return xmlDoc;
}