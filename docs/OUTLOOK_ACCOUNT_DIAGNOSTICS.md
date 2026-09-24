# Outlook Account Diagnostics for julian.garrett@aliniant.com

## Problem
Getting instant "None of your email accounts could send to this recipient" error when trying to send emails.

## Diagnosis Steps

### 1. Check Account Settings in Outlook

1. Open Outlook
2. Click **File** → **Account Settings** → **Account Settings**
3. Look for `julian.garrett@aliniant.com`
4. What does it say under **Type**?
   - Exchange
   - IMAP
   - POP3
   - Microsoft 365
   - Other

### 2. Check if Account Can Send

**For Exchange/Microsoft 365 accounts:**
- These should work automatically for sending
- If not working, the account may be disconnected or credentials expired

**For IMAP/POP3 accounts:**
- These are often **receive-only** by default
- You MUST configure SMTP settings separately for sending
- Check if there's a "More Settings" button → "Outgoing Server" tab

### 3. Look for SMTP Server Settings

1. In Account Settings, select the account
2. Click **Change**
3. Click **More Settings** or **Advanced**
4. Look for:
   - **Outgoing mail server (SMTP)**: Should have a server address
   - **Port**: Usually 587 (TLS) or 465 (SSL)
   - **Authentication**: Should require login
   - **Username/Password**: Should be configured

### 4. Common Causes

**Cause A: Account is IMAP/POP3 without SMTP**
- You can receive but not send
- **Fix**: Add SMTP server settings

**Cause B: Credentials Expired**
- Outlook can't authenticate
- **Fix**: Re-enter password or re-add the account

**Cause C: No Default Account**
- Outlook doesn't know which account to use for sending
- **Fix**: Set a default sending account

**Cause D: Account Disconnected**
- Shows in account list but not actually connected
- **Fix**: Remove and re-add the account

## Quick Tests

### Test 1: Check if ANY Account Can Send
Try sending from a different email account in your Outlook to see if the issue is:
- Specific to julian.garrett@aliniant.com
- OR affects all accounts

### Test 2: Manual Account Selection
When composing an email:
1. Click the **From** dropdown
2. Does `julian.garrett@aliniant.com` appear?
3. Try selecting it explicitly
4. Try to send

### Test 3: Send to Internal Address
Try sending to another @aliniant.com email address instead of external
- If this works, external sending might be blocked
- If this fails too, account is completely non-functional for sending

## Solutions

### Solution 1: Re-add the Account (Easiest)
1. File → Account Settings → Account Settings
2. Select the account → **Remove**
3. Click **New** → Add the account again with full credentials
4. Outlook should auto-configure settings

### Solution 2: Configure SMTP Manually
If re-adding doesn't work, you need manual SMTP settings from your IT department:
- SMTP Server address
- Port number
- Authentication type
- Username/password

### Solution 3: Use Webmail Temporarily
If the account setup is complex, use the web interface:
- Go to webmail.aliniant.com (or whatever your company uses)
- Send from there while troubleshooting Outlook

## For Aliniant IT Support

If this is a company account, you may need to contact IT support with these details:
- Account: julian.garrett@aliniant.com
- Issue: Cannot send emails from Outlook - getting "None of your email accounts could send"
- Classification tags ARE being added (MSCAN working)
- Need SMTP configuration or account re-provisioning

## Next Steps

1. Check account type (Step 1 above)
2. Report back what type of account it is
3. I can provide specific instructions based on the account type

