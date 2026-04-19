using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal sealed class PathCompatibilityService : IPathCompatibilityService
    {
        private const int MoveRetryCountPrimary = 20;
        private const int MoveRetryCountFallback = 20;
        private readonly Logger _logger;

        internal PathCompatibilityService(Logger logger = null)
        {
            _logger = logger;
        }

        public string NormalizePath(string path)
        {
            string normalized = (path ?? string.Empty).Trim();
            if (normalized.Length == 0)
            {
                return string.Empty;
            }

            if (normalized.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
            {
                normalized = normalized.Substring(8);
                normalized = normalized.Replace("/", "\\");
                return CollapseBackslashesSafe(normalized);
            }

            if (IsHttpUrl(normalized))
            {
                normalized = ConvertOneDriveUrlToLocalBestEffort(normalized);
            }

            normalized = normalized.Replace("/", "\\");
            return CollapseBackslashesSafe(normalized);
        }

        internal string CombinePath(string left, string right)
        {
            string normalizedLeft = NormalizePath(left);
            string normalizedRight = NormalizePath(right);
            if (normalizedLeft.Length == 0)
            {
                return normalizedRight;
            }

            if (normalizedRight.Length == 0)
            {
                return normalizedLeft;
            }

            if (!normalizedLeft.EndsWith("\\", StringComparison.Ordinal))
            {
                normalizedLeft += "\\";
            }

            if (normalizedRight.StartsWith("\\", StringComparison.Ordinal))
            {
                normalizedRight = normalizedRight.Substring(1);
            }

            return normalizedLeft + normalizedRight;
        }

        internal string GetParentFolderPath(string fullPath)
        {
            string normalized = NormalizePath(fullPath);
            if (normalized.Length == 0)
            {
                return string.Empty;
            }

            int separatorIndex = normalized.LastIndexOf('\\');
            return separatorIndex <= 0 ? string.Empty : normalized.Substring(0, separatorIndex);
        }

        internal string GetFileNameFromPath(string fullPath)
        {
            string normalized = NormalizePath(fullPath);
            if (normalized.Length == 0)
            {
                return string.Empty;
            }

            int separatorIndex = normalized.LastIndexOf('\\');
            return separatorIndex < 0 ? normalized : normalized.Substring(separatorIndex + 1);
        }

        internal string ResolveToExistingLocalPath(string path)
        {
            string trimmed = (path ?? string.Empty).Trim();
            if (trimmed.Length == 0)
            {
                return string.Empty;
            }

            if (trimmed.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
            {
                string localFilePath = NormalizePath(trimmed);
                return PathExistsLocal(localFilePath) ? localFilePath : string.Empty;
            }

            if (!IsHttpUrl(trimmed))
            {
                string localPath = NormalizePath(trimmed);
                return PathExistsLocal(localPath) ? localPath : string.Empty;
            }

            LogHttpObservation("ResolveToExistingLocalPath", "Inspecting HTTP path. input=" + trimmed);
            string relativePath = ExtractRelativePathFromCloudUrl(trimmed);
            if (relativePath.Length == 0)
            {
                _logger?.Warn("PathCompatibilityService.ResolveToExistingLocalPath could not extract relative path from HTTP input. input=" + trimmed);
                return string.Empty;
            }

            foreach (string root in GetSyncRootCandidates())
            {
                string candidate = NormalizePath(CombinePath(root, relativePath));
                if (PathExistsLocal(candidate))
                {
                    _logger?.Info("PathCompatibilityService.ResolveToExistingLocalPath resolved HTTP path. input=" + trimmed + " relative=" + relativePath + " resolved=" + candidate);
                    return candidate;
                }
            }

            _logger?.Warn("PathCompatibilityService.ResolveToExistingLocalPath did not find a local candidate for HTTP path. input=" + trimmed + " relative=" + relativePath);
            return string.Empty;
        }

        internal string BuildSafeSavePath(string rawFullPath)
        {
            string normalizedPath = NormalizePath(rawFullPath);
            if (normalizedPath.Length == 0)
            {
                return string.Empty;
            }

            string folderPath = GetParentFolderPath(normalizedPath);
            string fileName = GetFileNameFromPath(normalizedPath);
            if (folderPath.Length == 0 || fileName.Length == 0)
            {
                return string.Empty;
            }

            folderPath = ResolveToExistingLocalPath(folderPath);
            if (folderPath.Length == 0)
            {
                if (IsHttpUrl(rawFullPath))
                {
                    _logger?.Warn("PathCompatibilityService.BuildSafeSavePath could not resolve HTTP folder to an existing local path. input=" + (rawFullPath ?? string.Empty));
                }
                return string.Empty;
            }

            string safePath = NormalizePath(CombinePath(folderPath, fileName));
            if (IsHttpUrl(rawFullPath))
            {
                _logger?.Info("PathCompatibilityService.BuildSafeSavePath resolved HTTP save path. input=" + (rawFullPath ?? string.Empty) + " resolved=" + safePath);
            }

            return safePath;
        }

        internal bool EnsureFolderSafe(string folderPath)
        {
            string normalizedPath = NormalizePath(folderPath);
            if (normalizedPath.Length == 0)
            {
                return false;
            }

            if (normalizedPath.EndsWith("\\", StringComparison.Ordinal))
            {
                normalizedPath = normalizedPath.Substring(0, normalizedPath.Length - 1);
            }

            try
            {
                Directory.CreateDirectory(normalizedPath);
                return true;
            }
            catch
            {
                return false;
            }
        }

        internal bool FileExistsSafe(string path)
        {
            string normalized = NormalizePath(path);
            try
            {
                return File.Exists(normalized);
            }
            catch
            {
                return false;
            }
        }

        internal bool DirectoryExistsSafe(string path)
        {
            string normalized = NormalizePath(path);
            try
            {
                return Directory.Exists(normalized);
            }
            catch
            {
                return false;
            }
        }

        internal bool PathExistsSafe(string path)
        {
            return FileExistsSafe(path) || DirectoryExistsSafe(path);
        }

        internal string BuildUniquePath(string outFolder, string baseName, string extension)
        {
            string folder = NormalizePath(outFolder);
            string safeBaseName = SanitizeFileName((baseName ?? string.Empty).Trim());
            if (safeBaseName.Length == 0)
            {
                safeBaseName = "\u6587\u66F8";
            }

            string normalizedExtension = (extension ?? string.Empty).Trim();
            if (normalizedExtension.Length == 0)
            {
                normalizedExtension = ".docx";
            }

            if (!normalizedExtension.StartsWith(".", StringComparison.Ordinal))
            {
                normalizedExtension = "." + normalizedExtension;
            }

            string firstCandidate = CombinePath(folder, safeBaseName + normalizedExtension);
            if (!FileExistsSafe(firstCandidate))
            {
                return firstCandidate;
            }

            for (int index = 2; index <= 9999; index++)
            {
                string candidate = CombinePath(folder, safeBaseName + "_" + index.ToString("00") + normalizedExtension);
                if (!FileExistsSafe(candidate))
                {
                    return candidate;
                }
            }

            return CombinePath(folder, safeBaseName + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + normalizedExtension);
        }

        internal string EnsureUniquePathStandard(string fullPath)
        {
            string normalizedPath = NormalizePath(fullPath);
            if (!FileExistsSafe(normalizedPath))
            {
                return normalizedPath;
            }

            int dotPosition = normalizedPath.LastIndexOf('.');
            string basePath = dotPosition > 0 ? normalizedPath.Substring(0, dotPosition) : normalizedPath;
            string extension = dotPosition > 0 ? normalizedPath.Substring(dotPosition) : string.Empty;

            for (int sequence = 2; sequence <= 99; sequence++)
            {
                string candidate = basePath + "_" + sequence.ToString("00") + extension;
                if (!FileExistsSafe(candidate))
                {
                    return candidate;
                }
            }

            return basePath + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + extension;
        }

        internal string EnsureUniqueDirectoryPathStandard(string directoryPath)
        {
            string normalizedPath = NormalizePath(directoryPath);
            if (!PathExistsSafe(normalizedPath))
            {
                return normalizedPath;
            }

            for (int sequence = 2; sequence <= 99; sequence++)
            {
                string candidate = normalizedPath + "_" + sequence.ToString("00");
                if (!PathExistsSafe(candidate))
                {
                    return candidate;
                }
            }

            return normalizedPath + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
        }

        internal bool MoveFileSafe(string sourcePath, string destinationPath)
        {
            string normalizedSource = NormalizePath(sourcePath);
            string normalizedDestination = NormalizePath(destinationPath);
            if (normalizedSource.Length == 0 || normalizedDestination.Length == 0)
            {
                return false;
            }

            if (!FileExistsSafe(normalizedSource))
            {
                return false;
            }

            string destinationFolder = GetParentFolderPath(normalizedDestination);
            if (destinationFolder.Length > 0 && !EnsureFolderSafe(destinationFolder))
            {
                return false;
            }

            for (int retry = 1; retry <= MoveRetryCountPrimary; retry++)
            {
                try
                {
                    PromoteFileToDestinationSafely(normalizedSource, normalizedDestination);
                    return true;
                }
                catch
                {
                    WaitRetryTickMs(100);
                }
            }

            for (int retry = 1; retry <= MoveRetryCountFallback; retry++)
            {
                try
                {
                    PromoteFileToDestinationSafely(normalizedSource, normalizedDestination);
                    return true;
                }
                catch
                {
                    WaitRetryTickMs(150);
                }
            }

            return false;
        }

        private void PromoteFileToDestinationSafely(string normalizedSource, string normalizedDestination)
        {
            string replacementTempPath = BuildAdjacentTempFilePath(normalizedDestination);
            string backupTempPath = BuildAdjacentBackupFilePath(normalizedDestination);

            try
            {
                File.Copy(normalizedSource, replacementTempPath, overwrite: false);

                if (File.Exists(normalizedDestination))
                {
                    if (!TryReplaceFile(replacementTempPath, normalizedDestination, backupTempPath))
                    {
                        ReplaceFileViaBackupSwap(replacementTempPath, normalizedDestination, backupTempPath);
                    }
                }
                else
                {
                    File.Move(replacementTempPath, normalizedDestination);
                }

                TryDeleteFileQuietly(normalizedSource);
            }
            catch
            {
                TryDeleteFileQuietly(replacementTempPath);
                RestoreDestinationFromBackupIfNeeded(normalizedDestination, backupTempPath);
                TryDeleteFileQuietly(backupTempPath);
                throw;
            }
        }

        private static bool TryReplaceFile(string replacementTempPath, string normalizedDestination, string backupTempPath)
        {
            try
            {
                File.Replace(replacementTempPath, normalizedDestination, backupTempPath);
                TryDeleteFileQuietly(backupTempPath);
                return true;
            }
            catch (IOException)
            {
                TryDeleteFileQuietly(backupTempPath);
                return false;
            }
            catch (UnauthorizedAccessException)
            {
                TryDeleteFileQuietly(backupTempPath);
                return false;
            }
            catch (PlatformNotSupportedException)
            {
                TryDeleteFileQuietly(backupTempPath);
                return false;
            }
            catch (NotSupportedException)
            {
                TryDeleteFileQuietly(backupTempPath);
                return false;
            }
            catch (ArgumentException)
            {
                TryDeleteFileQuietly(backupTempPath);
                return false;
            }
        }

        private static void ReplaceFileViaBackupSwap(string replacementTempPath, string normalizedDestination, string backupTempPath)
        {
            File.Move(normalizedDestination, backupTempPath);

            try
            {
                File.Move(replacementTempPath, normalizedDestination);
                TryDeleteFileQuietly(backupTempPath);
            }
            catch
            {
                RestoreDestinationFromBackupIfNeeded(normalizedDestination, backupTempPath);
                throw;
            }
        }

        internal bool IsUnderSyncRoot(string path)
        {
            string normalizedPath = NormalizePath(path);
            if (normalizedPath.Length == 0)
            {
                return false;
            }

            foreach (string syncRoot in GetSyncRootCandidates())
            {
                string normalizedRoot = NormalizePath(syncRoot);
                if (normalizedRoot.Length == 0)
                {
                    continue;
                }

                string prefix = normalizedRoot.EndsWith("\\", StringComparison.Ordinal)
                    ? normalizedRoot
                    : normalizedRoot + "\\";
                if (normalizedPath.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(normalizedPath, normalizedRoot, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        internal string GetLocalTempWorkFolder(string subFolderName)
        {
            string rootPath = (Environment.GetEnvironmentVariable("LOCALAPPDATA") ?? string.Empty).Trim();
            if (rootPath.Length == 0)
            {
                rootPath = (Environment.GetEnvironmentVariable("TEMP") ?? string.Empty).Trim();
            }

            if (rootPath.Length == 0)
            {
                return string.Empty;
            }

            rootPath = NormalizePath(rootPath);
            string tempPath = string.IsNullOrWhiteSpace(subFolderName)
                ? rootPath
                : CombinePath(rootPath, subFolderName.Trim());

            try
            {
                Directory.CreateDirectory(tempPath);
                return NormalizePath(tempPath);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool PathExistsLocal(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            try
            {
                return File.Exists(path) || Directory.Exists(path);
            }
            catch
            {
                return false;
            }
        }

        private static bool IsHttpUrl(string value)
        {
            string normalized = (value ?? string.Empty).Trim();
            return normalized.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("https://", StringComparison.OrdinalIgnoreCase);
        }

        private string ExtractRelativePathFromCloudUrl(string url)
        {
            string trimmed = (url ?? string.Empty).Trim();
            if (trimmed.StartsWith("https://d.docs.live.net/", StringComparison.OrdinalIgnoreCase))
            {
                int fourthSlashIndex = FindSlashOccurrence(trimmed, 4);
                if (fourthSlashIndex > 0 && fourthSlashIndex + 1 < trimmed.Length)
                {
                    string encodedRelativePath = trimmed.Substring(fourthSlashIndex + 1);
                    LogHttpDecodeObservation("ExtractRelativePathFromCloudUrl", trimmed, encodedRelativePath);
                    return UrlDecode(encodedRelativePath).Replace("/", "\\");
                }
            }

            string[] markers = { "/Documents/", "/Shared%20Documents/", "/Shared Documents/" };
            foreach (string marker in markers)
            {
                int markerIndex = trimmed.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
                if (markerIndex > 0)
                {
                    string encodedRelativePath = trimmed.Substring(markerIndex + 1);
                    LogHttpDecodeObservation("ExtractRelativePathFromCloudUrl", trimmed, encodedRelativePath);
                    return UrlDecode(encodedRelativePath).Replace("/", "\\");
                }
            }

            return string.Empty;
        }

        private static IEnumerable<string> GetSyncRootCandidates()
        {
            var candidates = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            AddIfNotBlank(candidates, Environment.GetEnvironmentVariable("OneDrive"));
            AddIfNotBlank(candidates, Environment.GetEnvironmentVariable("OneDriveCommercial"));
            AddIfNotBlank(candidates, Environment.GetEnvironmentVariable("OneDriveConsumer"));
            AddIfNotBlank(candidates, Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "OneDrive"));
            return candidates;
        }

        private static void AddIfNotBlank(ISet<string> candidates, string value)
        {
            string trimmed = (value ?? string.Empty).Trim();
            if (trimmed.Length == 0)
            {
                return;
            }

            candidates.Add(trimmed);
        }

        private string BuildAdjacentTempFilePath(string destinationPath)
        {
            return BuildAdjacentWorkingFilePath(destinationPath, "tmp");
        }

        private string BuildAdjacentBackupFilePath(string destinationPath)
        {
            return BuildAdjacentWorkingFilePath(destinationPath, "bak");
        }

        private string BuildAdjacentWorkingFilePath(string destinationPath, string suffix)
        {
            string normalizedDestination = NormalizePath(destinationPath);
            string destinationFolder = GetParentFolderPath(normalizedDestination);
            string fileName = GetFileNameFromPath(normalizedDestination);
            string baseName = Path.GetFileNameWithoutExtension(fileName);
            string extension = Path.GetExtension(fileName);
            string tempFileName = (baseName.Length == 0 ? "temp" : baseName)
                + "." + (suffix ?? "tmp") + "_" + Guid.NewGuid().ToString("N")
                + extension;
            return CombinePath(destinationFolder, tempFileName);
        }

        private static string SanitizeFileName(string value)
        {
            string sanitized = value ?? string.Empty;
            char[] invalidChars = Path.GetInvalidFileNameChars();
            foreach (char invalidChar in invalidChars)
            {
                sanitized = sanitized.Replace(invalidChar, '_');
            }

            while (sanitized.Contains("  "))
            {
                sanitized = sanitized.Replace("  ", " ");
            }

            return sanitized.Trim().TrimEnd('.', ' ');
        }

        private static void WaitRetryTickMs(int milliseconds)
        {
            DateTime endAt = DateTime.UtcNow.AddMilliseconds(milliseconds);
            while (DateTime.UtcNow < endAt)
            {
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private static void TryDeleteFileQuietly(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }

        private static void RestoreDestinationFromBackupIfNeeded(string destinationPath, string backupPath)
        {
            if (string.IsNullOrWhiteSpace(destinationPath) || string.IsNullOrWhiteSpace(backupPath))
            {
                return;
            }

            try
            {
                if (!File.Exists(destinationPath) && File.Exists(backupPath))
                {
                    File.Move(backupPath, destinationPath);
                }
            }
            catch
            {
            }
        }

        private string ConvertOneDriveUrlToLocalBestEffort(string url)
        {
            string trimmed = (url ?? string.Empty).Trim();
            if (!trimmed.StartsWith("https://d.docs.live.net/", StringComparison.OrdinalIgnoreCase))
            {
                return trimmed;
            }

            int fourthSlashIndex = FindSlashOccurrence(trimmed, 4);
            if (fourthSlashIndex <= 0 || fourthSlashIndex + 1 >= trimmed.Length)
            {
                return trimmed;
            }

            string encodedRelativePath = trimmed.Substring(fourthSlashIndex + 1);
            LogHttpDecodeObservation("ConvertOneDriveUrlToLocalBestEffort", trimmed, encodedRelativePath);
            string relativePath = UrlDecode(encodedRelativePath);
            string oneDriveRoot = Environment.GetEnvironmentVariable("OneDrive");
            if (string.IsNullOrWhiteSpace(oneDriveRoot))
            {
                oneDriveRoot = Environment.GetEnvironmentVariable("OneDriveCommercial");
            }

            if (string.IsNullOrWhiteSpace(oneDriveRoot))
            {
                _logger?.Warn("PathCompatibilityService.ConvertOneDriveUrlToLocalBestEffort could not find a OneDrive root. input=" + trimmed + " relative=" + relativePath);
                return trimmed;
            }

            string combinedPath = Path.Combine(oneDriveRoot, relativePath);
            _logger?.Info("PathCompatibilityService.ConvertOneDriveUrlToLocalBestEffort combined docs.live URL with OneDrive root. input=" + trimmed + " relative=" + relativePath + " combined=" + combinedPath);
            return combinedPath;
        }

        private static int FindSlashOccurrence(string value, int occurrence)
        {
            int slashCount = 0;
            for (int index = 0; index < value.Length; index++)
            {
                if (value[index] != '/')
                {
                    continue;
                }

                slashCount++;
                if (slashCount == occurrence)
                {
                    return index;
                }
            }

            return -1;
        }

        private static string CollapseBackslashesSafe(string path)
        {
            bool isUncPath = path.StartsWith("\\\\", StringComparison.Ordinal);
            string collapsed = path;
            while (collapsed.Contains("\\\\"))
            {
                collapsed = collapsed.Replace("\\\\", "\\");
            }

            if (isUncPath)
            {
                collapsed = "\\" + collapsed;
            }

            if (collapsed.StartsWith("\\\\\\", StringComparison.Ordinal))
            {
                collapsed = "\\\\" + collapsed.Substring(3);
            }

            return collapsed;
        }

        private static string UrlDecode(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            var builder = new StringBuilder(value.Length);
            for (int index = 0; index < value.Length; index++)
            {
                char current = value[index];
                if (current == '%' && index + 2 < value.Length)
                {
                    string hex = value.Substring(index + 1, 2);
                    if (IsHex2(hex))
                    {
                        builder.Append((char)Convert.ToInt32(hex, 16));
                        index += 2;
                        continue;
                    }
                }

                builder.Append(current == '+' ? ' ' : current);
            }

            return builder.ToString();
        }

        private static bool IsHex2(string value)
        {
            if (string.IsNullOrEmpty(value) || value.Length != 2)
            {
                return false;
            }

            for (int index = 0; index < value.Length; index++)
            {
                char current = value[index];
                bool isHex = (current >= '0' && current <= '9')
                    || (current >= 'A' && current <= 'F')
                    || (current >= 'a' && current <= 'f');
                if (!isHex)
                {
                    return false;
                }
            }

            return true;
        }

        private void LogHttpDecodeObservation(string operationName, string originalUrl, string encodedRelativePath)
        {
            if (_logger == null)
            {
                return;
            }

            if (ContainsPercentEncodedHighByte(encodedRelativePath))
            {
                _logger.Warn("PathCompatibilityService." + operationName + " observed percent-encoded high-byte path data. input=" + originalUrl + " encodedRelative=" + encodedRelativePath);
                return;
            }

            if (encodedRelativePath.IndexOf('%') >= 0 || encodedRelativePath.IndexOf('+') >= 0)
            {
                _logger.Info("PathCompatibilityService." + operationName + " observed encoded path data. input=" + originalUrl + " encodedRelative=" + encodedRelativePath);
            }
        }

        private void LogHttpObservation(string operationName, string message)
        {
            _logger?.Info("PathCompatibilityService." + operationName + " " + (message ?? string.Empty));
        }

        private static bool ContainsPercentEncodedHighByte(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            for (int index = 0; index + 2 < value.Length; index++)
            {
                if (value[index] != '%')
                {
                    continue;
                }

                string hex = value.Substring(index + 1, 2);
                if (!IsHex2(hex))
                {
                    continue;
                }

                if (Convert.ToInt32(hex, 16) >= 0x80)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
