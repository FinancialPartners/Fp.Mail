# Fp.Mail

Send e-mail from .NET through **Microsoft Graph**: text or HTML, cc/bcc/reply-to, importance and attachments up to 150 MB. Works with dependency injection and includes an in-memory test double.

```
dotnet add package Fp.Mail
```

## Configure

`appsettings.json`:

```json
{
  "FpMail": {
    "TenantId": "00000000-0000-0000-0000-000000000000",
    "ClientId": "00000000-0000-0000-0000-000000000000",
    "Secret": "<client secret, or leave empty to use managed identity / az login>",
    "Sender": "noreply@contoso.com",
    "SenderName": "Contoso",
    "SaveToSentItems": false,
    "RedirectAllTo": [],
    "BccAll": []
  }
}
```

Keep the secret out of source control (user secrets, environment variable, Key Vault). Options are validated on startup.

```csharp
builder.Services.AddFpMail();                       // binds the "FpMail" section
// or
builder.Services.AddFpMail(o => { o.TenantId = "..."; o.ClientId = "..."; o.Sender = "noreply@contoso.com"; });
```

## Send

```csharp
public class InvoiceNotifier(IMailService mail)
{
    public async Task Notify(byte[] pdf, CancellationToken ct)
    {
        var message = new MailMessage { Subject = "Your invoice", Body = "<p>See attached.</p>", IsHtml = true };
        message.To.Add(new MailAddress("customer@example.com", "Customer"));
        message.Attachments.Add(new MailAttachment("invoice.pdf", pdf, "application/pdf"));

        await mail.SendAsync(message, ct);
    }
}
```

- `From` defaults to `Sender`. Set `message.From` to send from another mailbox the app may use.
- Attachments up to 3 MB in total are sent in one request. Larger messages are sent through a draft with upload sessions (this always saves a copy in Sent Items).
- Failures throw `MailException`; invalid input throws `ArgumentException`.

## Test and acceptance environments

- `RedirectAllTo`: when filled, every mail goes only to these addresses instead of the real recipients.
- `BccAll`: these addresses get a blind copy of every mail.
- In unit tests use `services.AddInMemoryMail()` and read `InMemoryMailService.SentMessages`.

## Azure setup

1. Create an app registration and a client secret (or use a managed identity).
2. Add the **Application** permission `Mail.Send` for Microsoft Graph and grant admin consent.
3. Restrict the app to specific mailboxes. Without this the app can send as **any** mailbox in your tenant. Use an *Application Access Policy* (Exchange Online) or RSC for Exchange Online.

You can check the setup with the `Fp.Mail.TestSender` console project in the repository: it asks for tenant, client, sender and recipient and sends a test mail.
