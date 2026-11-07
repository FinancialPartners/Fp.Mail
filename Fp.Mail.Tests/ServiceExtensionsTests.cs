using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;

namespace Fp.Mail.Tests;

public sealed class ServiceExtensionsTests
{
    private static ServiceProvider Build(Action<IServiceCollection> register)
    {
        var services = new ServiceCollection();
        services.AddLogging();
        register(services);
        return services.BuildServiceProvider();
    }

    [Fact]
    public void AddFpMail_BindsConfigurationSection()
    {
        var config = new ConfigurationBuilder().AddInMemoryCollection(new Dictionary<string, string?>
        {
            ["FpMail:TenantId"] = "tenant",
            ["FpMail:ClientId"] = "client",
            ["FpMail:Secret"] = "secret",
            ["FpMail:Sender"] = "noreply@contoso.com",
            ["FpMail:RedirectAllTo:0"] = "test@contoso.com",
        }).Build();

        using var provider = Build(s => s.AddSingleton<IConfiguration>(config).AddFpMail());

        var options = provider.GetRequiredService<IOptions<GraphMailOptions>>().Value;
        Assert.Equal("noreply@contoso.com", options.Sender);
        Assert.Equal(["test@contoso.com"], options.RedirectAllTo);
        Assert.IsType<GraphMailService>(provider.GetRequiredService<IMailService>());
    }

    [Fact]
    public void AddFpMail_Action_ConfiguresOptions()
    {
        using var provider = Build(s => s.AddFpMail(o =>
        {
            o.TenantId = "tenant";
            o.ClientId = "client";
            o.Secret = "secret";
            o.Sender = "noreply@contoso.com";
        }));

        Assert.IsType<GraphMailService>(provider.GetRequiredService<IMailService>());
    }

    [Fact]
    public void AddFpMail_InvalidOptions_FailOnAccess()
    {
        using var provider = Build(s => s.AddFpMail(o => o.TenantId = "tenant"));

        Assert.Throws<OptionsValidationException>(() => provider.GetRequiredService<IOptions<GraphMailOptions>>().Value);
    }

    [Fact]
    public void AddFpMail_InvalidSenderAddress_Fails()
    {
        using var provider = Build(s => s.AddFpMail(o =>
        {
            o.TenantId = "t";
            o.ClientId = "c";
            o.Sender = "not-an-address";
        }));

        Assert.Throws<OptionsValidationException>(() => provider.GetRequiredService<IOptions<GraphMailOptions>>().Value);
    }

    [Fact]
    public void AddFpMail_DoesNotOverrideExistingRegistration()
    {
        var custom = new InMemoryMailService();

        using var provider = Build(s => s.AddSingleton<IMailService>(custom).AddFpMail(o =>
        {
            o.TenantId = "t";
            o.ClientId = "c";
            o.Sender = "noreply@contoso.com";
        }));

        Assert.Same(custom, provider.GetRequiredService<IMailService>());
    }

    [Fact]
    public async Task AddInMemoryMail_ReplacesServiceAndRecordsMessages()
    {
        using var provider = Build(s => s.AddFpMail(o =>
        {
            o.TenantId = "t";
            o.ClientId = "c";
            o.Sender = "noreply@contoso.com";
        }).AddInMemoryMail());

        var mail = provider.GetRequiredService<IMailService>();
        var message = new MailMessage { Subject = "x" };
        await mail.SendAsync(message);

        var memory = provider.GetRequiredService<InMemoryMailService>();
        Assert.Same(memory, mail);
        Assert.Same(message, Assert.Single(memory.SentMessages));
        memory.Clear();
        Assert.Empty(memory.SentMessages);
    }

    [Fact]
    public void MailAddress_ConvertsImplicitlyFromString()
    {
        MailAddress address = "a@contoso.com";

        Assert.Equal("a@contoso.com", address.Address);
        Assert.Null(address.Name);
    }
}
