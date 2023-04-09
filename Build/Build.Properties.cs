partial class Build
{
    readonly string[] Projects =
    {
        "Elnahla.ExcelToRevit"
    };

    public const string InstallerProject = "Installer";

    public const string BuildConfiguration = "Release";
    public const string InstallerConfiguration = "Installer";

    const string AddInBinPrefix = "AddIn";
    const string ArtifactsFolder = "output";
}