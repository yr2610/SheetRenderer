using System;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using ExcelDna.Integration;

internal enum AddInEdition
{
    Sync,
    Renderer
}

internal static class AddInProfile
{
    internal const string FileName = "addon-profile.json";

    private static readonly Lazy<ProfileState> current =
        new Lazy<ProfileState>(Load, true);

    internal static AddInEdition Edition
    {
        get { return current.Value.Edition; }
    }

    internal static bool CanUseSheetRenderCommands
    {
        get { return Edition == AddInEdition.Renderer; }
    }

    private static ProfileState Load()
    {
        string profilePath = Path.Combine(GetAddInDirectory(), FileName);

        try
        {
            if (!File.Exists(profilePath))
            {
                return UseSafeDefault(profilePath, "profile file was not found");
            }

            string json = File.ReadAllText(profilePath);
            var document = JsonSerializer.Deserialize<ProfileDocument>(
                json,
                new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true,
                    AllowTrailingCommas = true
                });

            if (document == null)
            {
                return UseSafeDefault(profilePath, "profile file was empty");
            }

            if (document.SchemaVersion != 1)
            {
                return UseSafeDefault(profilePath, "unsupported schemaVersion");
            }

            if (string.Equals(document.Edition, "renderer", StringComparison.OrdinalIgnoreCase))
            {
                return new ProfileState(AddInEdition.Renderer, profilePath);
            }

            if (string.Equals(document.Edition, "sync", StringComparison.OrdinalIgnoreCase))
            {
                return new ProfileState(AddInEdition.Sync, profilePath);
            }

            return UseSafeDefault(profilePath, "unknown edition");
        }
        catch (Exception ex)
        {
            return UseSafeDefault(profilePath, ex.Message);
        }
    }

    private static ProfileState UseSafeDefault(string profilePath, string reason)
    {
        Trace.TraceWarning(
            "SheetRenderer add-in profile could not be loaded. " +
            "SheetSync edition will be used. Path=" + profilePath +
            " Reason=" + reason);

        return new ProfileState(AddInEdition.Sync, profilePath);
    }

    private static string GetAddInDirectory()
    {
        try
        {
            string xllPath = ExcelDnaUtil.XllPath;
            if (!string.IsNullOrWhiteSpace(xllPath))
            {
                string directory = Path.GetDirectoryName(xllPath);
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    return directory;
                }
            }
        }
        catch
        {
        }

        return AppDomain.CurrentDomain.BaseDirectory;
    }

    private sealed class ProfileDocument
    {
        public int SchemaVersion { get; set; }

        public string Edition { get; set; }
    }

    private sealed class ProfileState
    {
        internal ProfileState(AddInEdition edition, string profilePath)
        {
            Edition = edition;
            ProfilePath = profilePath;
        }

        internal AddInEdition Edition { get; private set; }

        internal string ProfilePath { get; private set; }
    }
}
