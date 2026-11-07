using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using Microsoft.Graph;

namespace Fp.Mail;

public static class ServiceExtensions
{
    /// <summary>
    /// Registers <see cref="IMailService"/> (Microsoft Graph), binding <see cref="GraphMailOptions"/> from the given configuration section.
    /// </summary>
    public static IServiceCollection AddFpMail(this IServiceCollection services, string sectionName = GraphMailOptions.DefaultSectionName)
    {
        services.AddOptions<GraphMailOptions>()
            .BindConfiguration(sectionName)
            .ValidateDataAnnotations()
            .ValidateOnStart();

        return services.AddGraphMailCore();
    }

    /// <summary>
    /// Registers <see cref="IMailService"/> (Microsoft Graph), configuring <see cref="GraphMailOptions"/> in code.
    /// </summary>
    public static IServiceCollection AddFpMail(this IServiceCollection services, Action<GraphMailOptions> configure)
    {
        services.AddOptions<GraphMailOptions>()
            .Configure(configure)
            .ValidateDataAnnotations()
            .ValidateOnStart();

        return services.AddGraphMailCore();
    }

    /// <summary>
    /// Registers <see cref="InMemoryMailService"/> as <see cref="IMailService"/>, for tests.
    /// </summary>
    public static IServiceCollection AddInMemoryMail(this IServiceCollection services)
    {
        services.RemoveAll<IMailService>();
        services.RemoveAll<InMemoryMailService>();
        services.AddSingleton<InMemoryMailService>();
        services.AddSingleton<IMailService>(sp => sp.GetRequiredService<InMemoryMailService>());
        return services;
    }

    internal static TokenCredential CreateCredential(GraphMailOptions options) =>
        string.IsNullOrWhiteSpace(options.Secret)
            ? new DefaultAzureCredential(new DefaultAzureCredentialOptions { TenantId = options.TenantId })
            : new ClientSecretCredential(options.TenantId, options.ClientId, options.Secret);

    private static IServiceCollection AddGraphMailCore(this IServiceCollection services)
    {
        // The Graph client is created here and not registered, so it cannot clash with other libraries using their own app registration.
        services.TryAddSingleton<IMailService>(sp => GraphMailService.Create(
            sp.GetRequiredService<IOptions<GraphMailOptions>>(),
            sp.GetService<ILogger<GraphMailService>>()));

        return services;
    }
}
