using System.ComponentModel.DataAnnotations;

namespace Fp.Mail;

public class GraphMailOptions
{
    public const string DefaultSectionName = "FpMail";

    [Required]
    public string TenantId { get; set; } = string.Empty;

    [Required]
    public string ClientId { get; set; } = string.Empty;

    /// <summary>
    /// Client secret of the app registration. When empty, <c>DefaultAzureCredential</c> is used (e.g. managed identity).
    /// </summary>
    public string Secret { get; set; } = string.Empty;

    /// <summary>
    /// Default mailbox to send from, e.g. <c>noreply@contoso.com</c>.
    /// </summary>
    [Required]
    [EmailAddress]
    public string Sender { get; set; } = string.Empty;

    /// <summary>
    /// Display name shown for <see cref="Sender"/>.
    /// </summary>
    public string? SenderName { get; set; }

    /// <summary>
    /// Keep a copy of sent messages in the sender's Sent Items. Messages with large attachments are always saved.
    /// </summary>
    public bool SaveToSentItems { get; set; }

    /// <summary>
    /// For test and acceptance environments: when filled, every message goes only to these addresses instead of its real recipients.
    /// </summary>
    public List<string> RedirectAllTo { get; set; } = [];

    /// <summary>
    /// Addresses that receive a blind copy of every message.
    /// </summary>
    public List<string> BccAll { get; set; } = [];
}
