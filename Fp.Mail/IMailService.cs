namespace Fp.Mail;

public interface IMailService
{
    /// <summary>
    /// Sends the message. Throws <see cref="ArgumentException"/> for an invalid message and <see cref="MailException"/> when sending fails.
    /// </summary>
    Task SendAsync(MailMessage message, CancellationToken cancellationToken = default);
}
