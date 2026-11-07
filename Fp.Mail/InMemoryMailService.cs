namespace Fp.Mail;

/// <summary>
/// Test double that keeps sent messages in memory instead of sending them.
/// </summary>
public sealed class InMemoryMailService : IMailService
{
    private readonly List<MailMessage> sent = [];
    private readonly object gate = new();

    public IReadOnlyList<MailMessage> SentMessages
    {
        get
        {
            lock (gate)
            {
                return sent.ToList();
            }
        }
    }

    public Task SendAsync(MailMessage message, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(message);
        lock (gate)
        {
            sent.Add(message);
        }

        return Task.CompletedTask;
    }

    public void Clear()
    {
        lock (gate)
        {
            sent.Clear();
        }
    }
}
