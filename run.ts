import {MsClient} from "./client";
import {LargeFileUploadSession, LargeFileUploadTask, Range, StreamUpload} from "@microsoft/microsoft-graph-client";
import * as fs from "fs";

require('isomorphic-fetch');

const client = new MsClient().getClient();

const message = {
    subject: 'This should have big PDF',
    body: {
        contentType: 'HTML',
        content: 'They were <b>awesome</b>!'
    },
    toRecipients: [
        {
            emailAddress: {
                address: 'iris.zhou@lever.co'
            }
        }
    ]
};


const fileUploadOptions = {
    rangeSize: 1024*1024,
    uploadEventHandlers: {
        // Called as each "slice" of the file is uploaded
        progress: (range?: Range, cb?: unknown) => {
            console.log(`Uploaded ${range?.minValue} to ${range?.maxValue}`);
        },
        extraCallbackParam: "any parameter needed by the callback implementation",
    }
};

const file = fs.createReadStream('./bigPDF.pdf');
const stats = fs.statSync("./bigPDF.pdf");
const totalSize = stats.size;

const payload = {
    AttachmentItem: {
        attachmentType: 'file',
        name: 'bigPDF.pdf',
        size: totalSize,
    }
}

async function uploadAndSend(): Promise<any> {
    //when using the token provided by our lever-microsoft, this will be a delegated auth-flow and we will be able to call /me/messages
    const draftMessage =  await client.api("/users/sadmin@levertest-native365.com/messages")
        .header('Prefer','IdType="ImmutableId"')
        .post(message);

    const messageId = draftMessage.id;
    const uploadSessionUrl = `/users/sadmin@levertest-native365.com/messages/${messageId}/attachments/createUploadSession`
    const uploadSession: LargeFileUploadSession = await LargeFileUploadTask.createUploadSession(client, uploadSessionUrl, payload);

    const fileObject = new StreamUpload(file, "bigPDF.pdf", totalSize);
    const task = new LargeFileUploadTask(client, fileObject, uploadSession, fileUploadOptions)
    task.upload().then(async (res) => {
        if (res) {
            console.log(res);
            return await client.api(`/users/sadmin@levertest-native365.com/messages/${messageId}/send`).post(message);
        }
    });
}

uploadAndSend()
    .then((sendResponse) => console.log('sent: ' + sendResponse))
    .catch((e) => console.log(e));


