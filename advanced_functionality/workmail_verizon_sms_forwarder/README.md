# Amazon WorkMail -> Verizon SMS Forwarder (Lambda)

This example shows the **most appropriate AWS endpoint** for your use case:

- Receive inbound mail in **Amazon WorkMail**.
- Trigger a **WorkMail message flow rule**.
- Invoke **AWS Lambda** to read the message and forward a short summary as SMS via Verizon's email gateway (`number@vtext.com`).

## Why this approach

For WorkMail-hosted inboxes, a WorkMail message flow rule is the native integration point. It avoids running your own HTTP endpoint and gives direct, authenticated access to inbound messages.

## Files

- `lambda_function.py`: Lambda handler that:
  1. Reads the WorkMail raw message from `workmailmessageflow:GetRawMessageContent`
  2. Extracts sender/subject/body
  3. Sends SMS-formatted email to Verizon via Amazon SES (`ses:SendEmail`)

## Prerequisites

1. Amazon WorkMail organization and mailbox/routing configured.
2. SES verified identity for `SOURCE_EMAIL` in the same region as Lambda.
3. Lambda execution role with:
   - `workmailmessageflow:GetRawMessageContent`
   - `ses:SendEmail`
   - CloudWatch Logs permissions

## Lambda environment variables

- `TARGET_PHONE` = Verizon number (10 digits, optional `+1` prefix)
- `SOURCE_EMAIL` = verified SES sender (for example, `alerts@example.com`)

## Create the WorkMail rule

1. Open **WorkMail Console** -> your organization.
2. Go to **Organization settings** -> **Message flow**.
3. Create an **inbound** rule with your conditions (all mail or specific recipients).
4. Set action to invoke this Lambda function.

## Behavior notes

- The function sends to `TARGET_PHONE@vtext.com`.
- Message body is compressed and truncated to ~160 chars for SMS friendliness.
- If you need MMS/longer payloads, switch gateway domain to `vzwpix.com`.

## Test locally (syntax check)

```bash
python -m py_compile lambda_function.py
```

## Example WorkMail event shape

```json
{
  "summaryVersion": "2018-10-10",
  "messageId": "EXAMPLE-MESSAGE-ID",
  "envelope": {
    "mailFrom": {
      "address": "sender@example.com"
    },
    "recipients": [
      {
        "address": "inbox@example.com"
      }
    ]
  }
}
```
