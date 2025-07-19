# Email System Documentation

## Overview
This document explains how to insert emails into the ERP system's emailing tables. The system stores email information in database tables and a Python service processes these tables to send emails and mark them as sent.

## Process Flow
1. Build the email (get template, generate email ID)
2. Insert email details into ZGEM_EMAILING table
3. Add email body (either template-based or manual)
4. Set email variables for personalization
5. Add attachments (reference only or saved copy)
6. Send the email

## Detailed Steps

### 1. Building the Email

#### 1.1 Get Email Template
```sql
:EXEC_ENAME = 'ZGEM_IAUDITEMAILTEMP';
:EXEC_TYPE = 'P';
:EXTMSG_MESSAGE = (:IS_FIX1PV = 'Y' ? '<!--Fix1PV-->' : '<!--Fix2-->');
:RUN_BY = 'ZCLA_FIXACTSTAT/ZGEM_SENDIAUDITEMAIL';
#INCLUDE ZGEM_EMAILING/GETEMAILTEMPLATE
```
This step retrieves the template IDs by calling the GETEMAILTEMPLATE procedure.
- Output: `:TEMPLATE_EXEC` and `:TEMPLATE_NUM`

#### 1.2 Generate Email ID
```sql
:RUN_BY = 'ZCLA_FIXACTSTAT/ZGEM_SENDIAUDITEMAIL';
#INCLUDE ZGEM_EMAILING/GETEMAILID
```
This step generates a unique email ID.
- Output: `:EMAILID`

### 2. Inserting Email Details

```sql
INSERT INTO ZGEM_EMAILING (EMAILID
, CURDATE
, SENDER_NAME
, SENDER_EMAIL
, RECIPIENT_EMAIL
, SUBJECT
, TEMPLATE_EXEC
, TEMPLATE_NUM
, TYPE
, USE_MANUAL_EMAILBODY)
SELECT :EMAILID,
       SQL.DATE,
       'Clarkson Evans Audits',
       'priority@clarksonevans.co.uk',
       'Recipient@email.com',
       STRCAT(:VAR1, 'subject', :VAR2),
       :TEMPLATE_EXEC,
       :TEMPLATE_NUM,
       'Example Email',
       (:IS_MANUAL_EMAIL_BODY = 'Y' ? 'Y' : '')
FROM DUMMY;
```

This step inserts the core email data into the ZGEM_EMAILING table:
- `EMAILID`: Unique identifier for the email
- `CURDATE`: Current date
- `SENDER_NAME`: Display name of the sender
- `SENDER_EMAIL`: Email address of the sender
- `RECIPIENT_EMAIL`: Email address of the recipient
- `SUBJECT`: Subject line of the email (can be concatenated from variables)
- `TEMPLATE_EXEC` and `TEMPLATE_NUM`: Template identifiers
- `TYPE`: Email type category
- `USE_MANUAL_EMAILBODY`: Controls whether to use template or manual body:
  - 'Y': Use manual email body
  - '' (empty): Use template body

### 3. Adding Email Body

#### Option 1: Template-Based Body
If `USE_MANUAL_EMAILBODY` is not 'Y', the system will use the body from the email template.

#### Option 2: Manual Email Body
If `USE_MANUAL_EMAILBODY` is 'Y', you must insert the body manually using the INSERTLINE procedure:

```sql
/*Body*/
:EMAIL_TEXT_LINE = '<p>Dear <RECIPIENT_NAME></p>';
#INCLUDE ZGEM_EMAILBODY/INSERTLINE
:EMAIL_TEXT_LINE = '<p>There is an emergency at <SITE> <PLOT></p>';
#INCLUDE ZGEM_EMAILBODY/INSERTLINE
:EMAIL_TEXT_LINE = '<p>Please see below the emergencies.</p>';
#INCLUDE ZGEM_EMAILBODY/INSERTLINE
```

The INSERTLINE procedure automatically handles line numbering by:
- Finding the maximum existing line number for the email
- Inserting the new line with the next sequential line number
- Using the same value for both `TEXTLINE` and `TEXTORD`

**Required inputs for INSERTLINE:**
- `EMAILID`: The unique email identifier
- `EMAIL_TEXT_LINE`: The content to be added to the email body

Note: Variables (like `<RECIPIENT_NAME>`, `<SITE>`, `<PLOT>`) can be used in manual email bodies for personalization and will be replaced with actual values when the email is processed.

### 4. Setting Email Variables

```sql
:VARIABLENAME = '<TL_SNAME>';
:VALUE = :TL_SNAME;
:RUN_BY = 'ZCLA_FIXACTSTAT/ZGEM_SENDIAUDITEMAIL';
#INCLUDE ZGEM_EMAILVARIABLES/INSERTEMAILVARIABLE
```

This step adds variables for personalization:
- `VARIABLENAME`: The placeholder in the email template (e.g., `<TL_SNAME>`)
- `VALUE`: The actual value to replace the placeholder

### 5. Adding Attachments

#### Option 1: Reference Only
```sql
:EXTFILENAME = :LATEST_EL_EXTFILENAME;
:NAME = :LATEST_EL_EXTFILEDES;
:RUN_BY = 'ZCLA_FIXACTSTAT/ZGEM_SENDIAUDITEMAIL';
#INCLUDE ZGEM_EMAILATTCHMNTS/INSERTEMAILATTACH
```

This option adds a reference to the attachment. If the file changes after sending the email, it will affect the file attached to the email.

#### Option 2: Save a Copy of the Attachment
```sql
/* Link ZGEM_EMAILATTCHMNTS Table */
:EMAIL_ATTACHMNTS_LINK = '';
SELECT SQL.TMPFILE INTO :EMAIL_ATTACHMNTS_LINK FROM DUMMY;
LINK ZGEM_EMAILATTCHMNTS TO :EMAIL_ATTACHMNTS_LINK;
GOTO 13112410529 WHERE :RETVAL <= 0;
DELETE FROM ZGEM_EMAILATTCHMNTS;

/* Insert Attachments */
:EXTFILENAME = :LATEST_EL_EXTFILENAME;
:NAME = :LATEST_EL_EXTFILEDES;
:RUN_BY = 'ZCLA_FIXACTSTAT/ZGEM_SENDIAUDITEMAIL';
#INCLUDE ZGEM_EMAILATTCHMNTS/INSERTEMAILATTACH

/* Insert attachments from linked table into ORIG table */
#INCLUDE ZGEM_EMAILING/LINKTABELATTACHMENTS

/* Un-Link ZGEM_EMAILATTCHMNTS Table */
UNLINK AND REMOVE ZGEM_EMAILATTCHMNTS;
LABEL 13112410529;
```

This option makes a copy of the attachment at the time of email creation. If the original file changes later, it won't affect the copy attached to the email. This method preserves email history by keeping an exact copy of what was sent.

### 6. Sending the Email

```sql
:RUN_BY = 'ZCLA_FIXACTSTAT/ZGEM_SENDIAUDITEMAIL';
#INCLUDE ZGEM_EMAILING/SENDEMAIL
```

This step marks the email as ready to send by setting `READY_TO_SEND = 'Y'`. The Python service will pick up this email and send it.

## Summary of Options

### Email Body Options
1. **Template-Based Body**: Use a predefined template (when `USE_MANUAL_EMAILBODY` is not 'Y')
2. **Manual Email Body**: Insert custom content into ZGEM_EMAILBODY table (when `USE_MANUAL_EMAILBODY` is 'Y')

### Attachment Options
1. **Reference Only**: Links to the original file (changes to the file will affect the email)
2. **Save Copy**: Makes a copy of the file at the time of email creation (preserves email history)

## Tables Used
- `ZGEM_EMAILING`: Main email information
- `ZGEM_EMAILBODY`: Email body content (for manual body)
- `ZGEM_EMAILVARIABLES`: Variables for personalization
- `ZGEM_EMAILATTCHMNTS`: Email attachments
