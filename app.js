// Copyright © 2025 Forcepoint LLC. All rights reserved.

let logEnable = false;
let urlDseRoot = 'https://localhost:55296/';

function sleep(delay) {
    const start = new Date().getTime();
    while (new Date().getTime() < start + delay);
}

Office.initialize = function () {}

function printLog(text) {
    console.log(text);
    if (logEnable && (typeof text === 'string' || text instanceof String)) {
        Office.context.mailbox.item.notificationMessages.replaceAsync("succeeded", {
            type: "progressIndicator",
            message: text.substring(0, Math.min(text.length, 250)),
        });
        sleep(1500);
    }
}

async function httpServerCheck(resolve, reject) {
    printLog("Checking the server");
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 30000);

    fetch(urlDseRoot + 'FirefoxExt/_1', {
        method: 'GET',
        mode: 'cors',
        cache: 'no-cache',
        credentials: 'same-origin',
        redirect: 'follow',
        referrerPolicy: 'no-referrer',
    }).then(response => {
        clearTimeout(timeout);
        if (!response.ok) {
            printLog("Server is down");
            reject(false);
        } else {
            printLog("Server is UP");
            resolve(true);
        }
    }).catch(e => {
        printLog("Request crashed");
        reject(false);
    });
}

async function sendToClasifier(url = '', data = {}, event) {
    printLog("Sending event to classifier");
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 35000);

    fetch(url, {
        signal: controller.signal,
        method: 'POST',
        mode: 'cors',
        cache: 'no-cache',
        credentials: 'same-origin',
        headers: { 'Content-Type': 'application/json' },
        redirect: 'follow',
        referrerPolicy: 'no-referrer',
        body: JSON.stringify(data)
    }).then(response => {
        if (!response.ok) {
            printLog("Engine returned error: " + response.status);
            handleError(response, event);
        }
        return response.json();
    }).then(response => {
        clearTimeout(timeout);
        handleResponse(response, event);
    }).catch(e => {
        printLog("Request crashed");
        printLog(e.name);
        handleError(e, event);
    });
}


function handleResponse(data, event) {
    printLog("Handling response from engine");
    let message = Office.context.mailbox.item;
    if (data["action"] === 1) {
        message.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Blocked by DLP engine' });
        printLog("DLP block");
        event.completed({ allowEvent: false });
    } else {
        printLog("DLP allow");
        event.completed({ allowEvent: true });
    }
}

async function tryPost(event, subject, from, to, cc, bcc, location, body, attachments) {
    printLog("Trying to post");
    let data = { subject, body, from, to, cc, bcc, location, attachments };
    if (attachments) printLog("Attachment list size: " + attachments.length);
    sendToClasifier(urlDseRoot + 'OutlookAddin', data, event);
}

async function postMessage(message, event, subject, from, to, cc, bcc, location, body, attachments){
    printLog("Posting message");
    if (attachments !== null) {
        await Promise.all(
            attachments.value.map(attachment => new Promise((resolve) => {
                message.getAttachmentContentAsync(attachment.id, data => {
                    let base64EncodedContent = data.value.content;
                    if (data.value.format !== "base64") {
                        base64EncodedContent = btoa(data.value.content);
                        printLog("Encoded attachment in base64");
                    }
                    resolve({
                        file_name: attachment.name,
                        data: base64EncodedContent,
                        content_type: attachment.contentType
                    });
                });
                setTimeout(() => resolve(null), 30000);
            }))
        ).then(result => {
            tryPost(event, subject, from, to, cc, bcc, location, body, result.filter(Boolean));
        });
    } else {
        tryPost(event, subject, from, to, cc, bcc, location, body, []);
    }
}

function getIfVal(result)
{
    return result.status === Office.AsyncResultStatus.Succeeded ? result.value : "";
}


async function validate(event) {
    message = Office.context.mailbox.item;

    if (message.itemType === "appointment") {
        printLog("Validating appointment")
        await Promise.all([
                new Promise((resolve,reject) =>{
                    httpServerCheck(resolve, reject)
                }),
                new Promise((resolve,reject) =>{
                    setTimeout(() => {resolve("")}, 3000)
                    message.subject.getAsync(result => {
                        resolve(getIfVal(result))
                    })
                }),
                new Promise((resolve,reject) =>{
                    setTimeout(() => {resolve("")}, 3000)
                    message.organizer.getAsync(result => {
                        resolve(getIfVal(result))
                    })
                }),
                new Promise((resolve,reject) =>{
                    setTimeout(() => {resolve("")}, 3000)
                    message.requiredAttendees.getAsync(result => {
                        resolve(getIfVal(result))
                    })
                }),
                new Promise((resolve,reject) =>{
                    setTimeout(() => {resolve("")}, 3000)
                    message.optionalAttendees.getAsync(result => {
                        resolve(getIfVal(result))
                    })
                }),
                new Promise((resolve,reject) =>{
                    setTimeout(() => {resolve("")}, 3000)
                    message.location.getAsync(result => {
                        resolve(getIfVal(result))
                    })
                }),
                new Promise((resolve,reject) =>{
                    setTimeout(() => {resolve("")}, 5000)
                    message.body.getAsync(Office.CoercionType.Html, { asyncContext: event }, result => {
					if (result.status === Office.AsyncResultStatus.Succeeded) {
						const htmlBody = result.value;

						// Normalize HTML to plain text
						const plainText = htmlBody
							.replace(/<[^>]+>/g, '');   // remove HTML tags
	
						// If needed, pass normalized body to detection logic
						resolve(plainText);
					} else {
						resolve("");
					}
					})
                }),
                new Promise((resolve,reject) =>{
                    setTimeout(() => {resolve(null)}, 5000)
                    message.getAttachmentsAsync(result => {
                        if (result.status === Office.AsyncResultStatus.Succeeded &&
                                result.value.length > 0) {
                            resolve(result)
                        }
                        resolve(null)
                    })
                })
        ]).then(([alive,subject, organizer, requiredAttendees, optionalAttendees, location, body, attachments]) => {	
            postMessage(message, event, subject, organizer, requiredAttendees, optionalAttendees, [],location, body, attachments);
        }).catch(err => {
            handleError("Server might be down", event);
        });
    } else if (message.itemType === "message") {
        printLog("Validating message")
        await Promise
            .all([
                new Promise((resolve,reject) =>{
                    httpServerCheck(resolve, reject)
                }),
                new Promise((resolve,reject) =>{
                    setTimeout(() => {resolve("")}, 3000)
                    message.subject.getAsync(result => {
                        resolve(getIfVal(result))
                    })
                }),
                new Promise((resolve,reject) =>{
                    setTimeout(() => {resolve("")}, 3000)
                    message.from.getAsync(result => {
                        resolve(getIfVal(result))
                    })
                }),
                new Promise((resolve,reject) =>{
                    setTimeout(() => {resolve("")}, 3000)
                    message.to.getAsync(result => {
                        resolve(getIfVal(result))
                    })
                }),
                new Promise((resolve,reject) =>{
                    setTimeout(() => {resolve("")}, 3000)
                    message.cc.getAsync(result => {
                        resolve(getIfVal(result))
                    })
                }),
                new Promise((resolve,reject) =>{
                    setTimeout(() => {resolve("")}, 3000)
                    message.bcc.getAsync(result => {
                        resolve(getIfVal(result))
                    })
                }),
                new Promise((resolve,reject) =>{
                    setTimeout(() => {resolve("")}, 5000)
                    message.body.getAsync(Office.CoercionType.Html, { asyncContext: event }, result => {
					if (result.status === Office.AsyncResultStatus.Succeeded) {
						const htmlBody = result.value;
	
						printLog("=== Raw HTML Body ===");
						printLog(htmlBody);
	
						// Normalize HTML to plain text
						const plainText = htmlBody
							.replace(/<[^>]+>/g, '');   // remove HTML tags
	
						printLog("=== Normalized Text ===");
						printLog(plainText);
	
						// If needed, pass normalized body to detection logic
						resolve(plainText);
					} else {
						resolve("");
					}
					})
                }),
                new Promise((resolve, reject) => {
                    setTimeout(() => {resolve(null)}, 5000)
                    message.getAttachmentsAsync(result => {
                        if (result.status === Office.AsyncResultStatus.Succeeded &&
                                result.value.length > 0) {
                            resolve(result)
                        }
                        resolve(null)
                    })
                })
                ])
            .then(([alive,subject, from, to, cc, bcc, body, attachments]) => {
                postMessage(message, event, subject, from, to, cc, bcc, "",  body, attachments);
            }) .catch(err => {
                printLog(err.message)
                handleError("Server might be down", event);
            });

    } else {
        printLog("message item type unknown");
        printLog(message.itemType);
        handleError("Unknown Message Type", event)
    }
}

function handleError(data, event) {
    printLog(data);
    printLog(event);
    printLog("Completing event ")
    event.completed({ allowEvent: true });
    printLog("Event Completed")
}

function operatingSytem() {
    var platform = Office.context.diagnostics.platform;
    if (platform === 'Mac') return 'MacOS';
    if (platform === 'OfficeOnline') return 'WindowsOS';
    return 'Other';
}

function onMessageSendHandler(event) {
    Office.onReady().then(function() {
        printLog("FP email validation started - [v1.2]")
        //Execute the add-in logic only if it is Outlook application running on MacOS
		
		var os = operatingSytem() 
        if(os == "MacOS"){
            printLog("MacOS detected")
			urlDseRoot = 'https://localhost:55296/';
            validate(event).catch(err => {handleError(err, event)});
        } else if(os == "WindowsOS"){
            printLog("WindowsOS detected")
		    urlDseRoot = 'https://localhost:55299/'
            validate(event).catch(err => {handleError(err, event)});
        } else{
            printLog("OS is not MacOS or WindowsOS")
            handleError("Not MacOS or WindowsOS", event);
        }
    });
}
