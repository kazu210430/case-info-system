# CI tests

This GitHub Actions workflow restores, compiles, and tests the solution on `windows-2022`.

Standard commands:

- `.\build.ps1 -Mode Compile`
  - CI-equivalent compile-only validation.
  - Internally this runs `dotnet build .\dev\CaseInfoSystem.slnx -c Release /p:AllowCoreBuildWithoutVstoPackaging=true /p:SignManifests=false /p:ManifestCertificateThumbprint=`.
  - It intentionally skips VSTO packaging and runtime `Addins/` reflection.
- `.\build.ps1 -Mode Test`
  - Standard local test command.
- `.\build.ps1 -Mode Test -Configuration Release -NoBuild -NoRestore`
  - CI test execution after the compile step has already finished.

Why CI does not use raw `dotnet build .\dev\CaseInfoSystem.slnx`:

- On MSBuild Core, raw `dotnet build` disables VSTO packaging.
- The Excel/Word Add-in projects then fail on purpose via the VSTO packaging guard so compile-only output is not mistaken for runtime deployment output.
- CI therefore uses the explicit compile-only entrypoint instead of bypassing the guard ad hoc.

Out of scope:

- Deploy package validation
- Runtime synchronization to `Addins/`
- Office / VSTO behavior on a real machine
