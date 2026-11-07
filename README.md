# Fp.Mail

[![Build](https://github.com/FinancialPartners/Fp.Mail/actions/workflows/build.yml/badge.svg)](https://github.com/FinancialPartners/Fp.Mail/actions/workflows/build.yml)
[![NuGet](https://img.shields.io/nuget/v/Fp.Mail.svg)](https://www.nuget.org/packages/Fp.Mail)

Send e-mail from .NET through Microsoft Graph. See [Fp.Mail/README.md](Fp.Mail/README.md) for usage.

## Repository layout

| Project | Purpose |
|---|---|
| `Fp.Mail` | The NuGet package |
| `Fp.Mail.Tests` | Unit tests (xunit, fake `HttpMessageHandler` under `GraphServiceClient`) |
| `Fp.Mail.TestSender` | Console tool that sends a test mail to verify an app registration and mailbox permissions |

## Build and test

```
dotnet test
```

## Releasing

Versions come from git tags (MinVer). Push a tag like `v0.1.0` and the release workflow tests, packs and publishes to nuget.org using trusted publishing (repository variable `NUGET_USER`, environment `production`), then creates a GitHub release.

## License

MIT
