const axios = require('axios');
const AWS = require('aws-sdk');
const qs = require('qs');

async function fetchUnreadEmails(args) {
    const {
        clientId, clientSecret, tenantId, email
    } = args;
    try {
        if (!clientId || !clientSecret || !tenantId) {
            throw new Error("clientId, clientSecret, tenantId are mandatory!!");
        }
        else {
            const token = await generateToken(args);
            const readEmailsUrl = `https://outlook.office365.com/api/v2.0/users/${email}/messages`
            const readEmailsObj = {
                method: 'get',
                maxBodyLength: Infinity,
                headers: {
                    'Authorization': `Bearer ${token?.access_token}`
                }
            }
            const emailMessages = await axios(readEmailsUrl, readEmailsObj);
            return emailMessages?.data?.value;
        }
    }
    catch (err) {
        throw err;
    }
}

async function fetchUnreadEmailsWithOnlyAttachments(args) {
    const {
        clientId, clientSecret, tenantId, email
    } = args;
    try {
        if (!clientId || !clientSecret || !tenantId || !email) {
            throw new Error("clientId, clientSecret, tenantId, email are mandatory!!");
        }
        else {
            const token = await generateToken(args);
            return await readAttachmentEmails(token?.access_token, email);
        }
    }
    catch (err) {
        throw err;
    }
}

async function readAttachmentEmails(token, email) {
    const readEmailsUrl = `https://outlook.office365.com/api/v2.0/users/${email}/messages?$filter=IsRead%20ne%20true%20and%20HasAttachments%20eq%20true`
    const readEmailsObj = {
        method: 'get',
        maxBodyLength: Infinity,
        headers: {
            'Authorization': `Bearer ${token}`
        }
    }
    try {
        const emailMessages = await axios(readEmailsUrl, readEmailsObj);
        return emailMessages?.data?.value;
    }
    catch (err) {
        throw err;
    }

}

async function generateToken({ clientId, clientSecret, tenantId }) {
    const tokenGenerationUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    try {
        const credentialsInfo = qs.stringify({
            'grant_type': 'client_credentials',
            'client_id': clientId,
            'scope': 'https://outlook.office365.com/.default',
            'client_secret': clientSecret
        });

        const bodyParams = {
            method: 'post',
            maxBodyLength: Infinity,
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            data: credentialsInfo
        }
        const tokenResponse = await axios(tokenGenerationUrl, bodyParams);
        return tokenResponse?.data;

    }
    catch (err) {
        throw err;
    }
}

async function readAttachmentUsingMessageId(args) {
    const {
        clientId, clientSecret, tenantId, email, messageId
    } = args;
    try {
        if (!clientId || !clientSecret || !tenantId || !messageId || !email) {
            throw new Error("clientId, clientSecret, tenantId, messageId, email are mandatory!!");
        }
        else {
            const token = await generateToken(args);
            const readAttachmentsUrl = `https://outlook.office365.com/api/v2.0/users/${email}/messages/${messageId}/attachments`
            const readAttachmentsObj = {
                method: 'get',
                maxBodyLength: Infinity,
                headers: {
                    'Authorization': `Bearer ${token?.access_token}`
                }
            }
            const emailMessages = await axios(readAttachmentsUrl, readAttachmentsObj);
            const attachments = [];
            for (let i = 0; i < emailMessages?.data?.value.length; i++) {
                attachments.push({
                    messageId,
                    fileName: emailMessages?.data?.value[i].Name,
                    attachmentBase64String: emailMessages?.data?.value[i].ContentBytes
                });
            }
            return attachments;
        }
    }
    catch (err) {
        throw err;
    }
}

async function readAndSaveAttachmentsToS3UsingMsgId(args) {
    const {
        clientId, clientSecret, tenantId, email, messageId, awsaccessKeyId, awssecretAccessKey, awsBucket
    } = args;
    try {
        if (!clientId || !clientSecret || !tenantId || !messageId || !email || !awsaccessKeyId || !awssecretAccessKey || !awsBucket) {
            throw new Error("clientId, clientSecret, tenantId, messageId, email, aws accesskey id, aws secret access key, aws bucket details are mandatory!!");
        }
        else {
            const token = await generateToken(args);
            return await getAttachmentsAndUploadAttachmentsToS3(email, messageId, awsaccessKeyId, awssecretAccessKey, awsBucket, token?.access_token);
        }
    }
    catch (err) {
        throw err;
    }
}

async function getAttachmentsAndUploadAttachmentsToS3(email, messageId, awsaccessKeyId, awssecretAccessKey, awsBucket, token) {
    const readAttachmentsUrl = `https://outlook.office365.com/api/v2.0/users/${email}/messages/${messageId}/attachments`
    const readAttachmentsObj = {
        method: 'get',
        maxBodyLength: Infinity,
        headers: {
            'Authorization': `Bearer ${token}`
        }
    }
    const emailMessages = await axios(readAttachmentsUrl, readAttachmentsObj);
    let uploadedAttachments = 0;
    for (let i = 0; i < emailMessages?.data?.value.length; i++) {
        const s3 = new AWS.S3({
            accessKeyId: awsaccessKeyId,
            secretAccessKey: awssecretAccessKey
        });

        const data = {
            Bucket: awsBucket,
            Key: emailMessages?.data?.value[i].Name,
            Body: Buffer.from(emailMessages?.data?.value[i].ContentBytes, 'base64'),
            ContentType: emailMessages?.data?.value[i].ContentType
        }

        await s3.upload(data).promise();

        uploadedAttachments++;
        if (i === emailMessages.data.value.length - 1) {
            const data = JSON.stringify({
                "IsRead": true
            });
            const patchRequestUrl = `https://outlook.office365.com/api/v2.0/users/${email}/messages/${messageId}`

            let config = {
                method: 'patch',
                maxBodyLength: Infinity,
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${token}`
                },
                data: data
            };

            await axios(patchRequestUrl, config);
        }
    }
    return uploadedAttachments;
}

async function readUnReadEmailsAttachmentsSaveToS3(args) {
    const {
        clientId, clientSecret, tenantId, email, awsaccessKeyId, awssecretAccessKey, awsBucket
    } = args;

    try {
        if (!clientId || !clientSecret || !tenantId || !email || !awsaccessKeyId || !awssecretAccessKey || !awsBucket) {
            throw new Error("clientId, clientSecret, tenantId, email, awsaccessKeyId, awssecretAccessKey, awsBucket are mandatory!!");
        }
        else {
            const token = await generateToken({ clientId, clientSecret, tenantId });
            const emails = await readAttachmentEmails(token?.access_token, email);
            let consolidatedAttachmentsUploaded = 0;
            for (let i = 0; i < emails.length; i++) {
                const attachmentsUploaded = await getAttachmentsAndUploadAttachmentsToS3(email, emails[i].Id, awsaccessKeyId, awssecretAccessKey, awsBucket, token?.access_token);
                consolidatedAttachmentsUploaded = consolidatedAttachmentsUploaded + attachmentsUploaded;
            }
            return `${consolidatedAttachmentsUploaded} uploaded to S3 bucket`;
        }
    }
    catch (err) {
        throw err;
    }
}

module.exports = {
    fetchUnreadEmails,
    fetchUnreadEmailsWithOnlyAttachments,
    readAttachmentUsingMessageId,
    readAndSaveAttachmentsToS3UsingMsgId,
    readUnReadEmailsAttachmentsSaveToS3
}
