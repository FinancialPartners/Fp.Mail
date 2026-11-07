using Fp.Mail;

Console.WriteLine("Fp.Mail test sender");
Console.WriteLine("Sends a test mail through Microsoft Graph to check your app registration and mailbox permissions.");
Console.WriteLine();

var options = new GraphMailOptions
{
    TenantId = Prompt("Tenant id"),
    ClientId = Prompt("Client id (empty = use your own login: Azure CLI / Visual Studio)", required: false),
};
if (!string.IsNullOrWhiteSpace(options.ClientId))
{
    options.Secret = PromptSecret("Client secret");
}
else
{
    options.ClientId = "-";
}

options.Sender = Prompt("Sender mailbox (e.g. noreply@contoso.com)");
var recipient = Prompt("Recipient");
var attachmentPath = Prompt("Attachment file (optional, empty = none)", required: false);

var message = new MailMessage
{
    Subject = "Fp.Mail test",
    Body = "<p>This is a test mail sent with <b>Fp.Mail</b>.</p>",
    IsHtml = true
};
message.To.Add(recipient);
if (!string.IsNullOrWhiteSpace(attachmentPath))
{
    message.Attachments.Add(new MailAttachment(Path.GetFileName(attachmentPath), await File.ReadAllBytesAsync(attachmentPath)));
}

try
{
    var service = GraphMailService.Create(options);
    await service.SendAsync(message);
    Console.WriteLine("Mail sent.");
    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Failed: {ex.Message}");
    return 1;
}

static string Prompt(string label, bool required = true)
{
    while (true)
    {
        Console.Write($"{label}: ");
        var value = Console.ReadLine()?.Trim() ?? string.Empty;
        if (value.Length > 0 || !required)
        {
            return value;
        }
    }
}

static string PromptSecret(string label)
{
    Console.Write($"{label}: ");
    var secret = new System.Text.StringBuilder();
    while (Console.ReadKey(intercept: true) is var key && key.Key != ConsoleKey.Enter)
    {
        if (key.Key == ConsoleKey.Backspace)
        {
            if (secret.Length > 0)
            {
                secret.Length--;
            }
        }
        else if (!char.IsControl(key.KeyChar))
        {
            secret.Append(key.KeyChar);
        }
    }

    Console.WriteLine();
    return secret.ToString();
}
