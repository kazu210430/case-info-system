using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal sealed class AddInRuntimeExecutionDiagnosticsService
    {
        private readonly Logger _logger;

        internal AddInRuntimeExecutionDiagnosticsService(Logger logger)
        {
            _logger = logger;
        }

        internal void Trace(string reason)
        {
            if (_logger == null)
            {
                return;
            }

            try
            {
                Assembly executingAssembly = Assembly.GetExecutingAssembly();
                string assemblyLocation = executingAssembly.Location ?? string.Empty;
                string baseDirectory = AppDomain.CurrentDomain.BaseDirectory ?? string.Empty;
                string primaryLogPath = ExcelAddInTraceLogWriter.GetPrimaryTraceLogPath();
                string fallbackLogPath = Path.Combine(Path.GetTempPath(), "CaseInfoSystem.ExcelAddIn", "CaseInfoSystem.ExcelAddIn_trace.log");
                string processId = Process.GetCurrentProcess().Id.ToString(CultureInfo.InvariantCulture);
                string assemblyLastWriteTimeUtc = SafeGetAssemblyLastWriteTimeUtc(assemblyLocation);
                string assemblyFileSize = SafeGetAssemblyFileSize(assemblyLocation);
                string assemblySha256 = SafeComputeAssemblySha256(assemblyLocation);
                string assemblyFileVersion = SafeGetAssemblyFileVersion(assemblyLocation);
                string assemblyVersion = SafeGetAssemblyVersion(executingAssembly);
                string assemblyInformationalVersion = SafeGetAssemblyInformationalVersion(executingAssembly);
                string assemblyBuildMarker = SafeGetAssemblyBuildMarker(executingAssembly);

                _logger.Info(
                    "Runtime execution observed. reason=" + (reason ?? string.Empty)
                    + ", assemblyLocation=" + assemblyLocation
                    + ", assemblyLastWriteTimeUtc=" + assemblyLastWriteTimeUtc
                    + ", assemblyFileSize=" + assemblyFileSize
                    + ", assemblySha256=" + assemblySha256
                    + ", assemblyFileVersion=" + assemblyFileVersion
                    + ", assemblyVersion=" + assemblyVersion
                    + ", assemblyInformationalVersion=" + assemblyInformationalVersion
                    + ", assemblyBuildMarker=" + assemblyBuildMarker
                    + ", appDomainBaseDirectory=" + baseDirectory
                    + ", primaryLogPath=" + primaryLogPath
                    + ", fallbackLogPath=" + fallbackLogPath
                    + ", pid=" + processId);
            }
            catch (Exception ex)
            {
                _logger.Error("Runtime execution observation failed. reason=" + (reason ?? string.Empty), ex);
            }
        }

        private static string SafeGetAssemblyLastWriteTimeUtc(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    return string.Empty;
                }

                return new FileInfo(path).LastWriteTimeUtc.ToString("O", CultureInfo.InvariantCulture);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeGetAssemblyFileSize(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    return string.Empty;
                }

                return new FileInfo(path).Length.ToString(CultureInfo.InvariantCulture);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeComputeAssemblySha256(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    return string.Empty;
                }

                using (FileStream stream = File.OpenRead(path))
                using (SHA256 sha256 = SHA256.Create())
                {
                    byte[] hash = sha256.ComputeHash(stream);
                    return BitConverter.ToString(hash).Replace("-", string.Empty);
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeGetAssemblyFileVersion(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    return string.Empty;
                }

                return FileVersionInfo.GetVersionInfo(path).FileVersion ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeGetAssemblyVersion(Assembly assembly)
        {
            try
            {
                return assembly == null
                    ? string.Empty
                    : (assembly.GetName().Version == null ? string.Empty : assembly.GetName().Version.ToString());
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeGetAssemblyInformationalVersion(Assembly assembly)
        {
            try
            {
                if (assembly == null)
                {
                    return string.Empty;
                }

                object[] attributes = assembly.GetCustomAttributes(typeof(AssemblyInformationalVersionAttribute), false);
                if (attributes == null || attributes.Length == 0)
                {
                    return string.Empty;
                }

                AssemblyInformationalVersionAttribute attribute = attributes[0] as AssemblyInformationalVersionAttribute;
                return attribute == null ? string.Empty : attribute.InformationalVersion ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeGetAssemblyBuildMarker(Assembly assembly)
        {
            try
            {
                return assembly == null ? string.Empty : assembly.ManifestModule.ModuleVersionId.ToString("D");
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
