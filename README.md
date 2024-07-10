
# dtai-outlook-email-processor

This npm package allows you to handle unread emails and attachments from an Outlook email account. You can fetch unread emails, fetch emails with attachments, read attachments, and save attachments to an S3 bucket.

## Features

1. Fetch all unread emails from an Outlook email account.
2. Fetch unread emails with attachments from an Outlook email account.
3. Fetch attachments based on `messageId` and return them in base64 format.
4. Fetch attachments based on `messageId`, upload them to an S3 bucket, and mark the email as read.
5. Fetch all attachments of unread emails, save them to an S3 bucket, and mark the emails as read.

## Installation

```sh
npm install outlook-email-handler
```

## Usage

### 1. Fetch All Unread Emails

Fetch all unread emails from an Outlook email account.

```javascript
const { fetchUnreadEmails } = require('outlook-email-handler');

const clientId = 'your-client-id';
const clientSecret = 'your-client-secret';
const tenantId = 'your-tenant-id';
const email = 'your-email@example.com';

async function getUnreadEmails() {
  try {
    const emails = await fetchUnreadEmails({ clientId, clientSecret, tenantId, email });
    console.log(emails);
  } catch (error) {
    console.error(error);
  }
}

getUnreadEmails();
```

### 2. Fetch Unread Emails with Attachments

Fetch all unread emails which have attachments from an Outlook email account.

```javascript
const { fetchUnreadEmailsWithOnlyAttachments } = require('outlook-email-handler');

const clientId = 'your-client-id';
const clientSecret = 'your-client-secret';
const tenantId = 'your-tenant-id';
const email = 'your-email@example.com';

async function getUnreadEmailsWithAttachments() {
  try {
    const emails = await fetchUnreadEmailsWithOnlyAttachments({ clientId, clientSecret, tenantId, email });
    console.log(emails);
  } catch (error) {
    console.error(error);
  }
}

getUnreadEmailsWithAttachments();
```

### 3. Fetch Attachment by Message ID

Fetch an attachment based on `messageId` and return it in base64 format.

```javascript
const { readAttachmentUsingMessageId } = require('outlook-email-handler');

const clientId = 'your-client-id';
const clientSecret = 'your-client-secret';
const tenantId = 'your-tenant-id';
const email = 'your-email@example.com';
const messageId = 'your-message-id';

async function getAttachmentByMessageId() {
  try {
    const attachment = await readAttachmentUsingMessageId({ clientId, clientSecret, tenantId, email, messageId });
    console.log(attachment);
  } catch (error) {
    console.error(error);
  }
}

getAttachmentByMessageId();
```

### 4. Fetch Attachment by Message ID, Upload to S3, and Mark Email as Read

Fetch an attachment based on `messageId`, upload it to an S3 bucket, and mark the email as read.

```javascript
const { fetchAttachmentAndUploadToS3 } = require('outlook-email-handler');

const clientId = 'your-client-id';
const clientSecret = 'your-client-secret';
const tenantId = 'your-tenant-id';
const email = 'your-email@example.com';
const messageId = 'your-message-id';
const awsaccessKeyId = 'your-aws-access-key-id';
const awssecretAccessKey = 'your-aws-secret-access-key';
const awsBucket = 'your-s3-bucket-name';

async function fetchAndUploadAttachment() {
  try {
    const response = await fetchAttachmentAndUploadToS3({ clientId, clientSecret, tenantId, email, messageId, awsaccessKeyId, awssecretAccessKey, awsBucket });
    console.log(response);
  } catch (error) {
    console.error(error);
  }
}

fetchAndUploadAttachment();
```

### 5. Fetch All Attachments of Unread Emails, Save to S3, and Mark Emails as Read

Fetch all attachments of unread emails, save them to an S3 bucket, and mark the emails as read.

```javascript
const { fetchUnreadEmailsWithAttachmentsAndUploadToS3 } = require('outlook-email-handler');

const clientId = 'your-client-id';
const clientSecret = 'your-client-secret';
const tenantId = 'your-tenant-id';
const email = 'your-email@example.com';
const awsaccessKeyId = 'your-aws-access-key-id';
const awssecretAccessKey = 'your-aws-secret-access-key';
const awsBucket = 'your-s3-bucket-name';

async function fetchAndUploadAllAttachments() {
  try {
    const response = await fetchUnreadEmailsWithAttachmentsAndUploadToS3({ clientId, clientSecret, tenantId, email, awsaccessKeyId, awssecretAccessKey, awsBucket });
    console.log(response);
  } catch (error) {
    console.error(error);
  }
}

fetchAndUploadAllAttachments();
```

## License

This project is licensed under the MIT License.
