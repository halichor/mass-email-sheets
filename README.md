# Flexible Mail Merge with Google Sheets and Google Docs

This script sends batch emails using data from a Google Sheet and a Google Docs template. It supports dynamic placeholder replacement (e.g., `{{FirstName}}`), scheduling, CC/BCC, and file attachments.

<img width="1615" alt="image" src="https://github.com/user-attachments/assets/4b8d88eb-5e3d-43ae-9a55-dfe53edd5d1d" />

---

## Prerequisites

To use this script from scratch, you’ll need:

- A Google account
- A Google Sheet with a sheet named `Data`
- A Google Docs file for the email template
- Access to **Apps Script** via `Extensions > Apps Script` in Google Sheets
- **Drive API enabled** for the script project

### Enable APIs

1. In the Apps Script editor, click on the **"+" icon next to "Services"** in the left sidebar.
2. From the list of available services, **add "Gmail API"** by selecting it and clicking **"Add"**.
3. Repeat the process to add other services:
   - **Drive API** – for fetching the Google Docs as HTML
  
---

## Sheet Setup

### 1. Create a sheet named `Data`

Create a new Google Sheet (or rename one) to `Data`.

### 2. Add column headers in **row 2**

Add the following fixed headers in **row 2**. The **Status** column should be at the very end.

To | CC | Bcc | Subject | Schedule | AttachmentIDs | (Your Custom Columns) | Status
--|--|--|--|--|--|--|--|

Other columns can be anything you like (e.g., `FirstName`, `AmountDue`) — their values will be used as `{{placeholders}}` in your Google Docs email template.

> If you want the header to be on **row 1**, change this in your script:
>
> ```js
> const HEADER_ROW_INDEX = 0; // row 1 = index 0
> ```

By default, data input starts on **row 4**. You can change it with:

```js
const DATA_START_ROW_INDEX = HEADER_ROW_INDEX + 2;
```

### 3. Set the Google Doc Template ID

Put the Doc ID of your Google Docs template in **cell B2**. This comes after `https://docs.google.com/document/d/` in the URL.

If you move the location, update this line in the script:

```js
const docId = sheet.getRange("B2").getValue();
```

---

## How to Run

1. Open the Apps Script editor.
2. Add a function trigger or assign the `sendFlexibleMailMerge` function to an image or button on your sheet.
   - To assign: Insert image > right click > **Assign Script** > `sendFlexibleMailMerge`
3. Authorize the script when prompted.
4. Click on the image/button you made to run the script. A prompt will appear, asking if you want to proceed with mass sending. Click 'Yes' to continue.
5. Rows with empty "Schedule" columns will be sent immediately. Otherwise, they will be scheduled based on the time and date specified in `MM/DD/YYYY HH:MM` format.

## Using Placeholders

Placeholders in the template should be wrapped in double braces, e.g., `{{FirstName}}`. This can be used in the email body through the Google Docs template or in the subject line. You can add as many columns as you want, and the placeholder name will be based on the column's header name.
