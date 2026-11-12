using Microsoft.Extensions.Logging;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>Publishes a newly created Excel file without replacing another writer's destination.</summary>
internal static class WorkbookCreationPublication
{
    internal static string ReserveStagingPath(string destination)
    {
        var directory = Path.GetDirectoryName(destination)!;
        var path = Path.Combine(directory, $".excelmcp-create-{Guid.NewGuid():N}{Path.GetExtension(destination)}");
        // Excel can overwrite this task's unique placeholder. It never receives the final path for SaveAs.
        using var reservation = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        return path;
    }

    internal static void Publish(string stagedPath, string destination)
    {
        try
        {
            // Staging beside the destination keeps publication on the same volume.
            File.Move(stagedPath, destination, overwrite: false);
        }
        catch (IOException ex) when (ex.HResult is unchecked((int)0x80070050) or unchecked((int)0x800700B7))
        {
            throw new WorkbookCreationPublicationException(destination, filePublished: false, ex);
        }
    }

    internal static string ReadIdentity(string path) =>
        WorkbookFileIdentity.Read(path).FileKey
        ?? throw new IOException("Could not confirm the new workbook file identity.");

    internal static void ConfirmIdentity(string path, string expectedIdentity)
    {
        if (!string.Equals(ReadIdentity(path), expectedIdentity, StringComparison.OrdinalIgnoreCase))
            throw new IOException("The published workbook path no longer identifies the file created by this request.");
    }

    internal static void DeleteStagingFile(string path, ILogger logger)
    {
        try
        {
            File.Delete(path);
        }
        catch (IOException ex)
        {
            logger.LogWarning(ex, "Could not remove the staged workbook at {Path}", path);
        }
        catch (UnauthorizedAccessException ex)
        {
            logger.LogWarning(ex, "Could not remove the staged workbook at {Path}", path);
        }
    }
}

internal sealed class WorkbookCreationPublicationException(string filePath, bool filePublished, Exception cause)
    : IOException(filePublished
        ? $"Workbook creation published '{filePath}', but Excel could not establish its session. Inspect the file before opening it. {cause.Message}"
        : $"File already exists: {filePath}. The existing workbook was not replaced.", cause)
{
    internal string FilePath { get; } = filePath;
    internal bool FilePublished { get; } = filePublished;
}
