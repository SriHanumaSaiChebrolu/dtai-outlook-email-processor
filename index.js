const axios = require('axios');
const AWS = require('aws-sdk');
const qs = require('qs');

// Fetches all unread emails from an outlook email
// List of args
// clientId - clientId captured from app registration in Azure with proper SMTP permissons
// clientSecret - clientSecret captured from app registration in Azure with proper SMTP permissons
// tenantId - tenantId of Azure account.
// email - email id from which unread emails has to be read
// response - Returns complete list of unread emails based on the above arguments
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
            const readEmailsUrl = `https://outlook.office365.com/api/v2.0/users/${email}/messages?$filter=IsRead%20ne%20true`
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

// Fetches all unread emails which has attachments from an outlook email
// List of args
// clientId - clientId captured from app registration in Azure with proper SMTP permissons
// clientSecret - clientSecret captured from app registration in Azure with proper SMTP permissons
// tenantId - tenantId of Azure account.
// email - email id from which unread emails has to be read
// response - Returns complete list of unread emails which has attachments based on the above arguments
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

// Fetches attachments based on messageId and returns base64 format of the attachment
// List of args
// clientId - clientId captured from app registration in Azure with proper SMTP permissons
// clientSecret - clientSecret captured from app registration in Azure with proper SMTP permissons
// tenantId - tenantId of Azure account.
// email - email id from which unread emails has to be read
// messageId = message id of the email for which attachments has to be read
// response - Returns attachment in base64 along with filename and messageId
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

// Fetches attachments based on messageId, uploads to S3 bucket and marks email as read
// List of args
// clientId - clientId captured from app registration in Azure with proper SMTP permissons
// clientSecret - clientSecret captured from app registration in Azure with proper SMTP permissons
// tenantId - tenantId of Azure account.
// email - email id from which unread emails has to be read
// messageId = message id of the email for which attachments has to be read
// awsaccessKeyId - aws access key of an user account which has permission to upload documents to S3.
// awssecretAccessKey - aws access secret key of an user account which has permission to upload documents to S3.
// awsBucket - aws bucket name to which documents has to be uploaded
// response - Returns number of files uploaded to S3 bucket.
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
            const countOfAttachmentsUploaded = await getAttachmentsAndUploadAttachmentsToS3(email, messageId, awsaccessKeyId, awssecretAccessKey, awsBucket, token?.access_token);

            return `${countOfAttachmentsUploaded} uploaded to provided S3 bucket`;
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

// Fetches all attachments of unread emails, saves them to S3 bucket and marks email as read
// List of args
// clientId - clientId captured from app registration in Azure with proper SMTP permissons
// clientSecret - clientSecret captured from app registration in Azure with proper SMTP permissons
// tenantId - tenantId of Azure account.
// email - email id from which unread emails has to be read
// awsaccessKeyId - aws access key of an user account which has permission to upload documents to S3.
// awssecretAccessKey - aws access secret key of an user account which has permission to upload documents to S3.
// awsBucket - aws bucket name to which documents has to be uploaded
// response - Returns number of files uploaded to S3 bucket.
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
