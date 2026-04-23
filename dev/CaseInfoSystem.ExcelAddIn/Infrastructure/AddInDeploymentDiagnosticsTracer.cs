using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Xml;
using Microsoft.Win32;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal static class AddInDeploymentDiagnosticsTracer
    {
        private const string ExcelAddInRegistrySubKey = @"Software\Microsoft\Office\Excel\Addins\CaseInfoSystem.ExcelAddIn";
        private const string WordAddInRegistrySubKey = @"Software\Microsoft\Office\Word\Addins\CaseInfoSystem.WordAddIn";
        private const string AddInAssemblyFileName = "CaseInfoSystem.ExcelAddIn.dll";
        private const string ApplicationManifestFileName = "CaseInfoSystem.ExcelAddIn.dll.manifest";
        private const string DeploymentManifestFileName = "CaseInfoSystem.ExcelAddIn.vsto";

        internal static void Trace(Logger logger)
        {
            if (logger == null)
            {
                return;
            }

            try
            {
                Assembly assembly = typeof(AddInDeploymentDiagnosticsTracer).Assembly;
                string assemblyLocation = SafeGetAssemblyLocation(assembly);
                string runtimeAddInDirectory = SafeGetDirectoryName(assemblyLocation);
                string repositoryRoot = TryResolveRepositoryRoot(runtimeAddInDirectory);
                string excelRegistryManifest = TraceAddInRegistration(logger, "Excel", ExcelAddInRegistrySubKey);
                string wordRegistryManifest = TraceAddInRegistration(logger, "Word", WordAddInRegistrySubKey);

                logger.Info(
                    "Add-in deployment diagnostics startup. "
                    + "timestampUtc=" + DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
                    + ", processId=" + SafeCurrentProcessId()
                    + ", processName=" + SafeCurrentProcessName()
                    + ", assemblyFullName=" + SafeGetAssemblyFullName(assembly)
                    + ", assemblyLocation=" + assemblyLocation
                    + ", assemblyCodeBase=" + SafeGetAssemblyCodeBase(assembly)
                    + ", assemblyVersion=" + SafeGetAssemblyVersion(assembly)
                    + ", assemblyFileVersion=" + SafeGetAssemblyFileVersion(assemblyLocation)
                    + ", appDomainBaseDirectory=" + SafeGetAppDomainBaseDirectory()
                    + ", runtimeAddInDirectory=" + runtimeAddInDirectory
                    + ", repositoryRoot=" + repositoryRoot
                    + ", primaryTraceLogPath=" + ExcelAddInTraceLogWriter.GetPrimaryTraceLogPath());

                logger.Info(
                    "Add-in deployment path candidates. "
                    + "loadedAssemblyPath=" + assemblyLocation
                    + ", runtimeAssemblyPath=" + CombineIfPossible(runtimeAddInDirectory, AddInAssemblyFileName)
                    + ", runtimeApplicationManifestPath=" + CombineIfPossible(runtimeAddInDirectory, ApplicationManifestFileName)
                    + ", runtimeDeploymentManifestPath=" + CombineIfPossible(runtimeAddInDirectory, DeploymentManifestFileName)
                    + ", debugBuildAssemblyPath=" + CombineIfPossible(repositoryRoot, @"開発物\dev\CaseInfoSystem.ExcelAddIn\bin\Debug\" + AddInAssemblyFileName)
                    + ", debugPackageAssemblyPath=" + CombineIfPossible(repositoryRoot, @"開発物\dev\Deploy\DebugPackage\CaseInfoSystem.ExcelAddIn\" + AddInAssemblyFileName)
                    + ", releasePackageAssemblyPath=" + CombineIfPossible(repositoryRoot, @"開発物\dev\Deploy\Package\CaseInfoSystem.ExcelAddIn\" + AddInAssemblyFileName)
                    + ", excelRegistryManifestPath=" + excelRegistryManifest
                    + ", wordRegistryManifestPath=" + wordRegistryManifest);

                TraceDeploymentManifest(logger, excelRegistryManifest, runtimeAddInDirectory);
                TraceApplicationManifest(logger, CombineIfPossible(runtimeAddInDirectory, ApplicationManifestFileName));
                TraceArtifactComparisons(logger, assemblyLocation, runtimeAddInDirectory, repositoryRoot, excelRegistryManifest);
            }
            catch (Exception ex)
            {
                logger.Error("Add-in deployment diagnostics trace failed.", ex);
            }
        }

        private static string TraceAddInRegistration(Logger logger, string hostName, string registrySubKey)
        {
            string manifestValue = string.Empty;

            try
            {
                using (RegistryKey currentUser = Registry.CurrentUser)
                using (RegistryKey key = currentUser == null ? null : currentUser.OpenSubKey(registrySubKey))
                {
                    if (key == null)
                    {
                        logger.Info("Add-in registration was not found. host=" + hostName + ", registrySubKey=" + registrySubKey);
                        return string.Empty;
                    }

                    manifestValue = Convert.ToString(key.GetValue("Manifest"), CultureInfo.InvariantCulture) ?? string.Empty;
                    int loadBehavior = SafeConvertToInt32(key.GetValue("LoadBehavior"));
                    string friendlyName = Convert.ToString(key.GetValue("FriendlyName"), CultureInfo.InvariantCulture) ?? string.Empty;
                    string description = Convert.ToString(key.GetValue("Description"), CultureInfo.InvariantCulture) ?? string.Empty;
                    string manifestPath = TryConvertManifestValueToLocalPath(manifestValue);

                    logger.Info(
                        "Add-in registration observed. "
                        + "host=" + hostName
                        + ", registrySubKey=" + registrySubKey
                        + ", loadBehavior=" + loadBehavior.ToString(CultureInfo.InvariantCulture)
                        + ", friendlyName=" + friendlyName
                        + ", description=" + description
                        + ", manifest=" + manifestValue
                        + ", manifestPath=" + manifestPath);
                }
            }
            catch (Exception ex)
            {
                logger.Error("Add-in registration trace failed. host=" + hostName + ", registrySubKey=" + registrySubKey, ex);
            }

            return TryConvertManifestValueToLocalPath(manifestValue);
        }

        private static void TraceDeploymentManifest(Logger logger, string deploymentManifestPath, string runtimeAddInDirectory)
        {
            string manifestPath = deploymentManifestPath;
            if (string.IsNullOrWhiteSpace(manifestPath))
            {
                manifestPath = CombineIfPossible(runtimeAddInDirectory, DeploymentManifestFileName);
            }

            if (string.IsNullOrWhiteSpace(manifestPath))
            {
                logger.Info("Deployment manifest trace skipped because no manifest path was available.");
                return;
            }

            try
            {
                logger.Info("Deployment manifest file read starting. path=" + manifestPath);
                var document = new XmlDocument();
                document.Load(manifestPath);

                XmlNode deploymentIdentity = document.SelectSingleNode("//*[local-name()='assemblyIdentity'][1]");
                XmlNode dependentAssembly = document.SelectSingleNode("//*[local-name()='dependency']/*[local-name()='dependentAssembly']");
                XmlNode applicationIdentity = document.SelectSingleNode("//*[local-name()='dependency']/*[local-name()='dependentAssembly']/*[local-name()='assemblyIdentity']");
                string applicationManifestCodebase = GetAttributeValue(dependentAssembly, "codebase");
                string applicationManifestPath = string.IsNullOrWhiteSpace(applicationManifestCodebase)
                    ? string.Empty
                    : Path.Combine(SafeGetDirectoryName(manifestPath), applicationManifestCodebase);

                logger.Info(
                    "Deployment manifest observed. "
                    + "path=" + manifestPath
                    + ", exists=" + File.Exists(manifestPath).ToString()
                    + ", deploymentIdentityName=" + GetAttributeValue(deploymentIdentity, "name")
                    + ", deploymentIdentityVersion=" + GetAttributeValue(deploymentIdentity, "version")
                    + ", applicationManifestCodebase=" + applicationManifestCodebase
                    + ", applicationManifestPath=" + applicationManifestPath
                    + ", applicationIdentityName=" + GetAttributeValue(applicationIdentity, "name")
                    + ", applicationIdentityVersion=" + GetAttributeValue(applicationIdentity, "version"));
            }
            catch (Exception ex)
            {
                logger.Error("Deployment manifest trace failed. path=" + manifestPath, ex);
            }
        }

        private static void TraceApplicationManifest(Logger logger, string applicationManifestPath)
        {
            if (string.IsNullOrWhiteSpace(applicationManifestPath))
            {
                return;
            }

            try
            {
                logger.Info("Application manifest file read starting. path=" + applicationManifestPath);
                var document = new XmlDocument();
                document.Load(applicationManifestPath);

                XmlNode assemblyIdentity = document.SelectSingleNode("//*[local-name()='assemblyIdentity'][1]");
                XmlNode entryPoint = document.SelectSingleNode("//*[local-name()='entryPoint']");

                logger.Info(
                    "Application manifest observed. "
                    + "path=" + applicationManifestPath
                    + ", exists=" + File.Exists(applicationManifestPath).ToString()
                    + ", identityName=" + GetAttributeValue(assemblyIdentity, "name")
                    + ", identityVersion=" + GetAttributeValue(assemblyIdentity, "version")
                    + ", entryPointClass=" + GetAttributeValue(entryPoint, "class"));
            }
            catch (Exception ex)
            {
                logger.Error("Application manifest trace failed. path=" + applicationManifestPath, ex);
            }
        }

        private static void TraceArtifactComparisons(Logger logger, string loadedAssemblyPath, string runtimeAddInDirectory, string repositoryRoot, string excelRegistryManifestPath)
        {
            TraceArtifactComparison(logger, "LoadedAssembly", loadedAssemblyPath, CombineIfPossible(runtimeAddInDirectory, AddInAssemblyFileName));
            TraceArtifactComparison(logger, "LoadedAssembly", loadedAssemblyPath, CombineIfPossible(repositoryRoot, @"開発物\dev\CaseInfoSystem.ExcelAddIn\bin\Debug\" + AddInAssemblyFileName));
            TraceArtifactComparison(logger, "LoadedAssembly", loadedAssemblyPath, CombineIfPossible(repositoryRoot, @"開発物\dev\Deploy\DebugPackage\CaseInfoSystem.ExcelAddIn\" + AddInAssemblyFileName));
            TraceArtifactComparison(logger, "LoadedAssembly", loadedAssemblyPath, CombineIfPossible(repositoryRoot, @"開発物\dev\Deploy\Package\CaseInfoSystem.ExcelAddIn\" + AddInAssemblyFileName));
            TraceArtifactComparison(logger, "RuntimeDeploymentManifest", CombineIfPossible(runtimeAddInDirectory, DeploymentManifestFileName), excelRegistryManifestPath);
            TraceArtifactComparison(logger, "RuntimeApplicationManifest", CombineIfPossible(runtimeAddInDirectory, ApplicationManifestFileName), CombineIfPossible(repositoryRoot, @"開発物\dev\Deploy\DebugPackage\CaseInfoSystem.ExcelAddIn\" + ApplicationManifestFileName));
        }

        private static void TraceArtifactComparison(Logger logger, string baselineLabel, string baselinePath, string candidatePath)
        {
            try
            {
                FileObservation baseline = ObserveFile(baselinePath);
                FileObservation candidate = ObserveFile(candidatePath);

                logger.Info(
                    "Deployment artifact comparison. "
                    + "baselineLabel=" + baselineLabel
                    + ", baselinePath=" + baseline.Path
                    + ", baselineExists=" + baseline.Exists.ToString()
                    + ", baselineLength=" + baseline.Length
                    + ", baselineLastWriteUtc=" + baseline.LastWriteUtc
                    + ", baselineFileVersion=" + baseline.FileVersion
                    + ", baselineSha256=" + baseline.Sha256
                    + ", candidatePath=" + candidate.Path
                    + ", candidateExists=" + candidate.Exists.ToString()
                    + ", candidateLength=" + candidate.Length
                    + ", candidateLastWriteUtc=" + candidate.LastWriteUtc
                    + ", candidateFileVersion=" + candidate.FileVersion
                    + ", candidateSha256=" + candidate.Sha256
                    + ", samePath=" + string.Equals(baseline.Path, candidate.Path, StringComparison.OrdinalIgnoreCase).ToString()
                    + ", sameSha256=" + (!string.IsNullOrWhiteSpace(baseline.Sha256) && string.Equals(baseline.Sha256, candidate.Sha256, StringComparison.OrdinalIgnoreCase)).ToString());
            }
            catch (Exception ex)
            {
                logger.Error(
                    "Deployment artifact comparison failed. baselinePath=" + (baselinePath ?? string.Empty) + ", candidatePath=" + (candidatePath ?? string.Empty),
                    ex);
            }
        }

        private static FileObservation ObserveFile(string path)
        {
            var observation = new FileObservation
            {
                Path = path ?? string.Empty,
                Exists = false,
                Length = string.Empty,
                LastWriteUtc = string.Empty,
                FileVersion = string.Empty,
                Sha256 = string.Empty
            };

            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                return observation;
            }

            observation.Exists = true;

            var fileInfo = new FileInfo(path);
            observation.Length = fileInfo.Length.ToString(CultureInfo.InvariantCulture);
            observation.LastWriteUtc = fileInfo.LastWriteTimeUtc.ToString("O", CultureInfo.InvariantCulture);
            observation.FileVersion = SafeGetAssemblyFileVersion(path);
            observation.Sha256 = ComputeSha256(path);
            return observation;
        }

        private static string ComputeSha256(string path)
        {
            try
            {
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

        private static string TryResolveRepositoryRoot(string runtimeAddInDirectory)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(runtimeAddInDirectory))
                {
                    return string.Empty;
                }

                DirectoryInfo addInDirectory = new DirectoryInfo(runtimeAddInDirectory);
                DirectoryInfo addinsDirectory = addInDirectory.Parent;
                DirectoryInfo repositoryRoot = addinsDirectory == null ? null : addinsDirectory.Parent;
                if (addinsDirectory == null || repositoryRoot == null)
                {
                    return string.Empty;
                }

                if (!string.Equals(addinsDirectory.Name, "Addins", StringComparison.OrdinalIgnoreCase))
                {
                    return string.Empty;
                }

                return repositoryRoot.FullName;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string TryConvertManifestValueToLocalPath(string manifestValue)
        {
            try
            {
                string normalized = manifestValue ?? string.Empty;
                int vstolocalIndex = normalized.IndexOf("|vstolocal", StringComparison.OrdinalIgnoreCase);
                if (vstolocalIndex >= 0)
                {
                    normalized = normalized.Substring(0, vstolocalIndex);
                }

                if (normalized.Length == 0)
                {
                    return string.Empty;
                }

                if (normalized.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
                {
                    return new Uri(normalized).LocalPath;
                }

                return normalized;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string CombineIfPossible(string basePath, string relativeOrFileName)
        {
            if (string.IsNullOrWhiteSpace(basePath))
            {
                return string.Empty;
            }

            return Path.Combine(basePath, relativeOrFileName ?? string.Empty);
        }

        private static string GetAttributeValue(XmlNode node, string attributeName)
        {
            if (node == null || node.Attributes == null || node.Attributes[attributeName] == null)
            {
                return string.Empty;
            }

            return node.Attributes[attributeName].Value ?? string.Empty;
        }

        private static string SafeGetAssemblyLocation(Assembly assembly)
        {
            try
            {
                return assembly == null ? string.Empty : assembly.Location ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeGetAssemblyCodeBase(Assembly assembly)
        {
            try
            {
                return assembly == null ? string.Empty : assembly.CodeBase ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeGetAssemblyFullName(Assembly assembly)
        {
            try
            {
                return assembly == null ? string.Empty : assembly.FullName ?? string.Empty;
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
                Version version = assembly == null ? null : assembly.GetName().Version;
                return version == null ? string.Empty : version.ToString();
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

        private static string SafeGetAppDomainBaseDirectory()
        {
            try
            {
                return AppDomain.CurrentDomain.BaseDirectory ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeGetDirectoryName(string path)
        {
            try
            {
                return string.IsNullOrWhiteSpace(path) ? string.Empty : (Path.GetDirectoryName(path) ?? string.Empty);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeCurrentProcessId()
        {
            try
            {
                return Process.GetCurrentProcess().Id.ToString(CultureInfo.InvariantCulture);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeCurrentProcessName()
        {
            try
            {
                return Process.GetCurrentProcess().ProcessName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static int SafeConvertToInt32(object value)
        {
            try
            {
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                return 0;
            }
        }

        private sealed class FileObservation
        {
            internal string Path { get; set; }

            internal bool Exists { get; set; }

            internal string Length { get; set; }

            internal string LastWriteUtc { get; set; }

            internal string FileVersion { get; set; }

            internal string Sha256 { get; set; }
        }
    }
}
