using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Graph.Users.Item.Messages.Item.Attachments.CreateUploadSession;
using Microsoft.Graph.Users.Item.SendMail;

namespace Fp.Mail;

public class GraphMailService(
    GraphServiceClient graphClient,
    IOptions<GraphMailOptions> mailOptions,
    ILogger<GraphMailService> logger) : IMailService
{
    /// <summary>Graph accepts attachments up to ~3 MB inline; the whole sendMail request is limited to about 4 MB.</summary>
    internal const int InlineLimitBytes = 3 * 1024 * 1024;
    internal const long MaxMessageBytes = 150L * 1024 * 1024;
    private const int ChunkSize = 4 * 320 * 1024;

    private readonly GraphMailOptions options = mailOptions.Value;

    /// <summary>
    /// Creates a service without dependency injection, e.g. for a console tool.
    /// </summary>
    public static GraphMailService Create(GraphMailOptions options, ILogger<GraphMailService>? logger = null) =>
        Create(Options.Create(options), logger);

    internal static GraphMailService Create(IOptions<GraphMailOptions> options, ILogger<GraphMailService>? logger) =>
        new(new GraphServiceClient(ServiceExtensions.CreateCredential(options.Value), [".default"]), options,
            logger ?? NullLogger<GraphMailService>.Instance);

    public async Task SendAsync(MailMessage message, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(message);

        var to = message.To.ToList();
        var cc = message.Cc.ToList();
        var bcc = message.Bcc.ToList();
        ApplyRecipientRules(to, cc, bcc);
        Validate(message, to, cc, bcc);

        var mailbox = message.From?.Address ?? options.Sender;
        var graphMessage = ToGraphMessage(message, mailbox, to, cc, bcc);
        var totalBytes = message.Attachments.Sum(a => (long)a.Content.Length);

        try
        {
            if (totalBytes <= InlineLimitBytes)
            {
                graphMessage.Attachments = message.Attachments.Select(ToGraphAttachment).ToList();
                await graphClient.Users[mailbox].SendMail.PostAsync(
                    new SendMailPostRequestBody { Message = graphMessage, SaveToSentItems = options.SaveToSentItems },
                    cancellationToken: cancellationToken);
            }
            else
            {
                await SendViaDraft(mailbox, graphMessage, message.Attachments, cancellationToken);
            }
        }
        catch (ODataError error)
        {
            var detail = error.Error?.Message ?? error.Message;
            throw new MailException($"Sending mail via Microsoft Graph failed: {detail}", error);
        }

        logger.LogInformation("Sent mail '{Subject}' from {Mailbox} to {Count} recipient(s)", message.Subject, mailbox, to.Count + cc.Count + bcc.Count);
    }

    private async Task SendViaDraft(string mailbox, Message graphMessage, List<MailAttachment> attachments, CancellationToken cancellationToken)
    {
        var user = graphClient.Users[mailbox];
        var draft = await user.Messages.PostAsync(graphMessage, cancellationToken: cancellationToken)
            ?? throw new MailException("Microsoft Graph did not return the draft message.");
        var draftId = draft.Id ?? throw new MailException("Microsoft Graph returned a draft message without id.");

        try
        {
            foreach (var attachment in attachments)
            {
                if (attachment.Content.Length <= InlineLimitBytes)
                {
                    await user.Messages[draftId].Attachments.PostAsync(ToGraphAttachment(attachment), cancellationToken: cancellationToken);
                }
                else
                {
                    await UploadLargeAttachment(mailbox, draftId, attachment, cancellationToken);
                }
            }

            await user.Messages[draftId].Send.PostAsync(cancellationToken: cancellationToken);
        }
        catch
        {
            await DeleteDraft(mailbox, draftId);
            throw;
        }
    }

    private async Task UploadLargeAttachment(string mailbox, string draftId, MailAttachment attachment, CancellationToken cancellationToken)
    {
        var body = new CreateUploadSessionPostRequestBody
        {
            AttachmentItem = new AttachmentItem
            {
                AttachmentType = AttachmentType.File,
                Name = attachment.FileName,
                Size = attachment.Content.Length,
                ContentType = attachment.ContentType
            }
        };

        var session = await graphClient.Users[mailbox].Messages[draftId].Attachments.CreateUploadSession
            .PostAsync(body, cancellationToken: cancellationToken);

        using var stream = new MemoryStream(attachment.Content, writable: false);
        var task = new LargeFileUploadTask<FileAttachment>(session, stream, ChunkSize, graphClient.RequestAdapter);
        var result = await task.UploadAsync(cancellationToken: cancellationToken);
        if (!result.UploadSucceeded)
        {
            throw new MailException($"Uploading attachment '{attachment.FileName}' did not succeed.");
        }
    }

    private async Task DeleteDraft(string mailbox, string draftId)
    {
        try
        {
            await graphClient.Users[mailbox].Messages[draftId].DeleteAsync();
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "Could not delete draft message {DraftId} of {Mailbox}", draftId, mailbox);
        }
    }

    private void ApplyRecipientRules(List<MailAddress> to, List<MailAddress> cc, List<MailAddress> bcc)
    {
        var redirect = options.RedirectAllTo.Where(a => !string.IsNullOrWhiteSpace(a)).ToList();
        if (redirect.Count > 0)
        {
            to.Clear();
            cc.Clear();
            bcc.Clear();
            to.AddRange(redirect.Select(a => new MailAddress(a)));
        }

        bcc.AddRange(options.BccAll.Where(a => !string.IsNullOrWhiteSpace(a)).Select(a => new MailAddress(a)));
    }

    private static void Validate(MailMessage message, List<MailAddress> to, List<MailAddress> cc, List<MailAddress> bcc)
    {
        if (to.Count + cc.Count + bcc.Count == 0)
        {
            throw new ArgumentException("A message needs at least one recipient.", nameof(message));
        }

        if (to.Concat(cc).Concat(bcc).Concat(message.ReplyTo).Any(a => string.IsNullOrWhiteSpace(a.Address)))
        {
            throw new ArgumentException("A recipient has an empty address.", nameof(message));
        }

        if (message.Attachments.Any(a => string.IsNullOrWhiteSpace(a.FileName)))
        {
            throw new ArgumentException("An attachment has no file name.", nameof(message));
        }

        if (message.Attachments.Sum(a => (long)a.Content.Length) > MaxMessageBytes)
        {
            throw new ArgumentException("The attachments exceed the 150 MB limit of Microsoft Graph.", nameof(message));
        }
    }

    private Message ToGraphMessage(MailMessage message, string mailbox, List<MailAddress> to, List<MailAddress> cc, List<MailAddress> bcc)
    {
        var sender = message.From ?? new MailAddress(options.Sender, options.SenderName);
        return new Message
        {
            Subject = message.Subject,
            Body = new ItemBody { ContentType = message.IsHtml ? BodyType.Html : BodyType.Text, Content = message.Body },
            From = ToRecipient(sender with { Address = mailbox }),
            ToRecipients = to.Select(ToRecipient).ToList(),
            CcRecipients = cc.Select(ToRecipient).ToList(),
            BccRecipients = bcc.Select(ToRecipient).ToList(),
            ReplyTo = message.ReplyTo.Select(ToRecipient).ToList(),
            Importance = message.Importance switch
            {
                MailImportance.Low => Importance.Low,
                MailImportance.High => Importance.High,
                _ => Importance.Normal
            }
        };
    }

    private static Recipient ToRecipient(MailAddress address) => new()
    {
        EmailAddress = new EmailAddress { Address = address.Address, Name = address.Name }
    };

    private static Attachment ToGraphAttachment(MailAttachment attachment) => new FileAttachment
    {
        Name = attachment.FileName,
        ContentType = attachment.ContentType,
        ContentBytes = attachment.Content
    };
}
