namespace Fp.Mail;

/// <summary>
/// An e-mail address with an optional display name. A plain string converts implicitly.
/// </summary>
public sealed record MailAddress(string Address, string? Name = null)
{
    public static implicit operator MailAddress(string address) => new(address);
}

/// <summary>
/// A file attached to a message. Total message size is limited to 150 MB by Microsoft Graph.
/// </summary>
public sealed record MailAttachment(string FileName, byte[] Content, string ContentType = "application/octet-stream");

public enum MailImportance
{
    Low,
    Normal,
    High
}

public sealed class MailMessage
{
    /// <summary>
    /// Mailbox to send from. Defaults to <see cref="GraphMailOptions.Sender"/>. The app registration needs permission on this mailbox.
    /// </summary>
    public MailAddress? From { get; set; }

    public List<MailAddress> To { get; } = [];
    public List<MailAddress> Cc { get; } = [];
    public List<MailAddress> Bcc { get; } = [];
    public List<MailAddress> ReplyTo { get; } = [];
    public string Subject { get; set; } = string.Empty;
    public string Body { get; set; } = string.Empty;
    public bool IsHtml { get; set; }
    public MailImportance Importance { get; set; } = MailImportance.Normal;
    public List<MailAttachment> Attachments { get; } = [];
}

public class MailException(string message, Exception? innerException = null) : Exception(message, innerException);
