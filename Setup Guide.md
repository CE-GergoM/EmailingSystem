# Email System Setup Guide (Superusers)

## Overview
This guide is for system administrators and superusers who need to configure the email system before developers can use it to send automated emails. The email system requires proper setup and configuration to function correctly.

## Prerequisites
- Superuser access to the ERP system
- Access to System Management and Office Management modules
- Understanding of your organization's email infrastructure

## Setup Steps

### 1. Enable the Emailing System

Navigate to: **System Management > System Maintenance > Companies > Companies**

1. Open the Companies form
2. Locate the **"Emailing System On"** tick-box
3. Check this box to enable the emailing system
4. Save the changes

**Important:** The emailing system will not function until this setting is enabled.

### 2. Configure Email Settings

Navigate to: **Office Management > Mail > Emailing > Email Settings**

Configure the basic email system settings:

#### 2.1 Recipient Whitelist Setting

Toggle the **RECIPIENT_WHITELIST** setting to control email filtering:

**Values:**
- **On**: Emails will only be sent to whitelisted recipients
- **Off**: Emails will be sent to all specified recipients

**When to use "On":**
- Testing environments with live data
- Development systems
- When you want to prevent accidental emails to real customers/users
- During system testing and validation

**When to use "Off":**
- Production environments
- When all recipients should receive emails normally
- After testing is complete

### 3. Manage Recipient Whitelist

Navigate to: **Office Management > Mail > Emailing > Recipient Whitelist**

#### 3.1 Adding Recipients to Whitelist

When the recipient whitelist is enabled ("On"), only email addresses in this list will receive emails.

To add recipients:
1. Open the Recipient Whitelist form
2. Add email addresses that should receive emails during testing
3. Include:
   - Test email addresses
   - Developer email addresses
   - Key stakeholder email addresses for testing
   - Any other addresses that should receive emails when filtering is active

#### 3.2 Managing the Whitelist

**Best Practices:**
- Keep the whitelist minimal during testing
- Include only necessary recipients to avoid spam
- Document why each address is whitelisted
- Regularly review and clean up the whitelist
- Remove test addresses when moving to production

### 4. Email Server Configuration

Ensure your email server settings are properly configured:

1. SMTP server settings
2. Authentication credentials
3. Port and security settings
4. Sender authentication (SPF, DKIM if applicable)

**Note:** Specific email server configuration steps depend on your email infrastructure and are typically handled by IT administrators.

### 5. Testing the Email System

#### 5.1 Enable Recipient Filtering for Testing
1. Set RECIPIENT_WHITELIST to "On"
2. Add your test email addresses to the whitelist
3. Have developers test email functionality
4. Verify that only whitelisted recipients receive emails

#### 5.2 Test Email Delivery
1. Create test emails through the system
2. Verify emails are sent to whitelisted recipients only
3. Check that non-whitelisted recipients do not receive emails
4. Test different recipient types (To, Cc, Bcc)

#### 5.3 Monitor Email Logs
- Check system logs for email processing
- Verify Python service is running and processing emails
- Monitor for any delivery failures or errors

### 6. Production Deployment

When ready to deploy to production:

1. **Disable Recipient Filtering:**
   - Set RECIPIENT_WHITELIST to "Off"
   - This allows emails to be sent to all specified recipients

2. **Clear Test Data:**
   - Remove test emails from the system
   - Clean up any test recipients from the whitelist (optional, as they won't be used when filtering is off)

3. **Final Verification:**
   - Test with a small subset of real recipients
   - Monitor email delivery and system performance
   - Ensure all email templates and variables are correctly configured

## Troubleshooting

### Common Issues

**Emails Not Being Sent:**
- Check if "Emailing System On" is enabled in Companies form
- Verify Python email service is running
- Check email server connectivity and credentials

**Emails Going to Wrong Recipients:**
- Verify RECIPIENT_WHITELIST setting
- Check whitelist configuration if filtering is enabled
- Review recipient entries in ZGEM_EMAIL_RECIPIENT table

**Recipients Not Receiving Emails:**
- If whitelist is "On", ensure recipients are in the whitelist
- Check email server logs for delivery issues
- Verify recipient email addresses are correct

### Monitoring and Maintenance

**Regular Tasks:**
- Monitor email delivery success rates
- Review and update recipient whitelist as needed
- Check system logs for errors or warnings
- Verify email templates are up to date

**Security Considerations:**
- Keep whitelist settings appropriate for environment
- Monitor for unauthorized email sending
- Review email content for sensitive information
- Ensure proper access controls on email configuration

## Configuration Summary

| Setting | Location | Purpose |
|---------|----------|---------|
| Emailing System On | System Management > Companies | Master switch for email system |
| RECIPIENT_WHITELIST | Office Management > Email Settings | Controls recipient filtering |
| Recipient Whitelist | Office Management > Recipient Whitelist | List of allowed recipients when filtering is on |

## Support

For technical issues with email system configuration:
1. Check system logs and error messages
2. Verify all configuration steps have been completed
3. Contact your IT administrator for email server issues
4. Document any unusual behavior for developer troubleshooting

Remember: Proper configuration of these settings is essential for the email system to function correctly and safely in both testing and production environments.
