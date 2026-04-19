# CI tests

This GitHub Actions workflow restores, builds, and tests the solution on `windows-2022`.

Current scope:

- Build `dev/CaseInfoSystem.slnx` via MSBuild
- Skip manifest signing in CI so certificate-store-dependent packaging does not block compile verification
- Run `dev/CaseInfoSystem.Tests/CaseInfoSystem.Tests.csproj`

Out of scope:

- Deploy package validation
- Runtime synchronization to `Addins/`
- Office / VSTO behavior on a real machine
