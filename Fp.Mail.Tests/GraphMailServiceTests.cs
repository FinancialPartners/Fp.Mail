using System.Net;
using System.Text;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using Microsoft.Graph;

namespace Fp.Mail.Tests;

public sealed class GraphMailServiceTests
{
    private sealed class FakeHandler(Func<HttpRequestMessage, HttpResponseMessage> respond) : HttpMessageHandler
    {
        public List<(HttpMethod Method, Uri Uri, string Body)> Requests { get; } = [];

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            var body = request.Content is null ? string.Empty : await request.Content.ReadAsStringAsync(cancellationToken);
            Requests.Add((request.Method, request.RequestUri!, body));
            return respond(request);
        }
    }

    private static HttpResponseMessage Json(string json, HttpStatusCode status = HttpStatusCode.OK) =>
        new(status) { Content = new StringContent(json, Encoding.UTF8, "application/json") };

    private static string Path(Uri uri) => Uri.UnescapeDataString(uri.AbsolutePath);

    private static HttpResponseMessage Accepted() => new(HttpStatusCode.Accepted);

    private static GraphMailService CreateService(FakeHandler handler, Action<GraphMailOptions>? configure = null)
    {
        var options = new GraphMailOptions { Sender = "noreply@contoso.com", SenderName = "Contoso" };
        configure?.Invoke(options);
        return new GraphMailService(new GraphServiceClient(new HttpClient(handler)), Options.Create(options), NullLogger<GraphMailService>.Instance);
    }

    private static MailMessage SimpleMessage()
    {
        var message = new MailMessage { Subject = "Hello", Body = "<p>Hi</p>", IsHtml = true };
        message.To.Add(new MailAddress("a@contoso.com", "Anna"));
        return message;
    }

    [Fact]
    public async Task SendAsync_PostsSendMailForDefaultSender()
    {
        var handler = new FakeHandler(_ => Accepted());

        await CreateService(handler).SendAsync(SimpleMessage());

        var request = Assert.Single(handler.Requests);
        Assert.Equal(HttpMethod.Post, request.Method);
        Assert.EndsWith("/users/noreply@contoso.com/sendMail", Path(request.Uri));
        Assert.Contains("\"subject\":\"Hello\"", request.Body);
        Assert.Contains("a@contoso.com", request.Body);
        Assert.Contains("\"contentType\":\"html\"", request.Body);
        Assert.Contains("\"saveToSentItems\":false", request.Body, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Contoso", request.Body);
    }

    [Fact]
    public async Task SendAsync_UsesFromMailboxWhenGiven()
    {
        var handler = new FakeHandler(_ => Accepted());
        var message = SimpleMessage();
        message.From = new MailAddress("sales@contoso.com", "Sales");

        await CreateService(handler).SendAsync(message);

        var request = Assert.Single(handler.Requests);
        Assert.EndsWith("/users/sales@contoso.com/sendMail", Path(request.Uri));
        Assert.Contains("Sales", request.Body);
    }

    [Fact]
    public async Task SendAsync_SmallAttachment_IsSentInline()
    {
        var handler = new FakeHandler(_ => Accepted());
        var message = SimpleMessage();
        message.Attachments.Add(new MailAttachment("a.txt", "hello"u8.ToArray(), "text/plain"));

        await CreateService(handler).SendAsync(message);

        var request = Assert.Single(handler.Requests);
        Assert.Contains("fileAttachment", request.Body);
        Assert.Contains("a.txt", request.Body);
        Assert.Contains(Convert.ToBase64String("hello"u8.ToArray()), request.Body);
    }

    [Fact]
    public async Task SendAsync_RedirectAllTo_ReplacesRecipients()
    {
        var handler = new FakeHandler(_ => Accepted());
        var message = SimpleMessage();
        message.Cc.Add("cc@contoso.com");
        message.Bcc.Add("bcc@contoso.com");

        await CreateService(handler, o => o.RedirectAllTo = ["test@contoso.com"]).SendAsync(message);

        var body = Assert.Single(handler.Requests).Body;
        Assert.Contains("test@contoso.com", body);
        Assert.DoesNotContain("a@contoso.com", body);
        Assert.DoesNotContain("cc@contoso.com", body);
        Assert.DoesNotContain("bcc@contoso.com", body);
    }

    [Fact]
    public async Task SendAsync_BccAll_IsAddedToBcc()
    {
        var handler = new FakeHandler(_ => Accepted());
        var message = SimpleMessage();
        message.Bcc.Add("bcc@contoso.com");

        await CreateService(handler, o => o.BccAll = ["archive@contoso.com"]).SendAsync(message);

        var body = Assert.Single(handler.Requests).Body;
        Assert.Contains("archive@contoso.com", body);
        Assert.Contains("bcc@contoso.com", body);
        Assert.Contains("a@contoso.com", body);
    }

    [Fact]
    public async Task SendAsync_DoesNotMutateTheMessage()
    {
        var message = SimpleMessage();

        await CreateService(new FakeHandler(_ => Accepted()), o => o.RedirectAllTo = ["test@contoso.com"]).SendAsync(message);

        Assert.Equal("a@contoso.com", Assert.Single(message.To).Address);
        Assert.Empty(message.Bcc);
    }

    [Fact]
    public async Task SendAsync_RejectsMessageWithoutRecipients()
    {
        var service = CreateService(new FakeHandler(_ => Accepted()));

        await Assert.ThrowsAsync<ArgumentException>(() => service.SendAsync(new MailMessage { Subject = "x" }));
    }

    [Fact]
    public async Task SendAsync_RejectsEmptyAddress()
    {
        var message = SimpleMessage();
        message.Cc.Add(" ");

        await Assert.ThrowsAsync<ArgumentException>(() => CreateService(new FakeHandler(_ => Accepted())).SendAsync(message));
    }

    [Fact]
    public async Task SendAsync_RejectsNullMessage()
    {
        await Assert.ThrowsAsync<ArgumentNullException>(() => CreateService(new FakeHandler(_ => Accepted())).SendAsync(null!));
    }

    [Fact]
    public async Task SendAsync_RejectsAttachmentWithoutName()
    {
        var message = SimpleMessage();
        message.Attachments.Add(new MailAttachment("", [1]));

        await Assert.ThrowsAsync<ArgumentException>(() => CreateService(new FakeHandler(_ => Accepted())).SendAsync(message));
    }

    [Fact]
    public async Task SendAsync_GraphError_ThrowsMailException()
    {
        var handler = new FakeHandler(_ => Json("""{"error":{"code":"ErrorAccessDenied","message":"Access is denied."}}""", HttpStatusCode.Forbidden));

        var ex = await Assert.ThrowsAsync<MailException>(() => CreateService(handler).SendAsync(SimpleMessage()));

        Assert.Contains("Access is denied.", ex.Message);
    }

    [Fact]
    public async Task SendAsync_LargeTotal_UsesDraftAndSend()
    {
        var handler = new FakeHandler(request => request.RequestUri!.AbsolutePath.EndsWith("/send")
            ? Accepted()
            : request.RequestUri.AbsolutePath.EndsWith("/messages")
                ? Json("""{"id":"draft1"}""", HttpStatusCode.Created)
                : Json("""{"id":"att1"}""", HttpStatusCode.Created));
        var message = SimpleMessage();
        var twoMb = new byte[2 * 1024 * 1024];
        message.Attachments.Add(new MailAttachment("one.bin", twoMb));
        message.Attachments.Add(new MailAttachment("two.bin", twoMb));

        await CreateService(handler).SendAsync(message);

        Assert.Equal(4, handler.Requests.Count);
        Assert.EndsWith("/users/noreply@contoso.com/messages", Path(handler.Requests[0].Uri));
        Assert.EndsWith("/messages/draft1/attachments", Path(handler.Requests[1].Uri));
        Assert.EndsWith("/messages/draft1/attachments", Path(handler.Requests[2].Uri));
        Assert.EndsWith("/messages/draft1/send", Path(handler.Requests[3].Uri));
        Assert.DoesNotContain("sendMail", Path(handler.Requests[0].Uri));
    }

    [Fact]
    public async Task SendAsync_LargeAttachment_CreatesUploadSession_AndDeletesDraftOnFailure()
    {
        var handler = new FakeHandler(request => request.Method == HttpMethod.Delete
            ? new HttpResponseMessage(HttpStatusCode.NoContent)
            : request.RequestUri!.AbsolutePath.EndsWith("/messages")
                ? Json("""{"id":"draft1"}""", HttpStatusCode.Created)
                : Json("""{"error":{"code":"stop","message":"stop"}}""", HttpStatusCode.BadRequest));
        var message = SimpleMessage();
        message.Attachments.Add(new MailAttachment("big.bin", new byte[4 * 1024 * 1024]));

        await Assert.ThrowsAsync<MailException>(() => CreateService(handler).SendAsync(message));

        Assert.Equal(3, handler.Requests.Count);
        Assert.EndsWith("/messages/draft1/attachments/createUploadSession", Path(handler.Requests[1].Uri));
        Assert.Contains("big.bin", handler.Requests[1].Body);
        Assert.Equal(HttpMethod.Delete, handler.Requests[2].Method);
        Assert.EndsWith("/messages/draft1", Path(handler.Requests[2].Uri));
    }

    [Fact]
    public async Task SendAsync_RejectsAttachmentsOverGraphLimit()
    {
        var message = SimpleMessage();
        message.Attachments.Add(new MailAttachment("a.bin", new byte[100 * 1024 * 1024]));
        message.Attachments.Add(new MailAttachment("b.bin", new byte[60 * 1024 * 1024]));

        await Assert.ThrowsAsync<ArgumentException>(() => CreateService(new FakeHandler(_ => Accepted())).SendAsync(message));
    }
}
