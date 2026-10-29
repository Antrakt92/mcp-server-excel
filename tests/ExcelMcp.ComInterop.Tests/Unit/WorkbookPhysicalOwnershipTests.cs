using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Feature", "SessionLifecycle")]
[Trait("Layer", "ComInterop")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class WorkbookPhysicalOwnershipTests
{
    [Fact]
    public void Reservations_HardLinksCannotAcquireSeparateOwners()
    {
        WithDirectory((root, actual, alias) =>
        {
            var original = Path.Combine(actual, "source.xlsx");
            var linked = Path.Combine(root, "linked.xlsx");
            File.WriteAllText(original, "synthetic");
            Link(linked, original);
            var claims = new WorkbookPathReservations<string>();
            Assert.True(claims.TryAdd(original, "first"));
            Assert.False(claims.TryAdd(linked, "second"));
            Assert.True(claims.TryGetValue(linked, out var owner));
            Assert.Equal("first", owner);
            Assert.True(WorkbookLocation.MatchesReportedLocation(original, linked));
        });
    }

    [Fact]
    public void Reservations_CreatedDestinationKeepsClaimAcrossHardLinkTransition()
    {
        WithDirectory((root, actual, alias) =>
        {
            var target = Path.Combine(actual, "new.xlsx");
            var linked = Path.Combine(root, "linked.xlsx");
            var claims = new WorkbookPathReservations<string>();
            Assert.True(claims.TryAdd(target, "pending"));
            File.WriteAllText(target, "synthetic");
            Link(linked, target);
            Assert.False(claims.TryAdd(linked, "other"));
            Assert.True(claims.TryGetValue(Path.Combine(alias, "new.xlsx"), out var owner));
            Assert.Equal("pending", owner);
            Assert.True(claims.TryRemove(new(target, "pending")));
            Assert.True(claims.TryAdd(linked, "other"));
        });
    }

    [Fact]
    public void Reservations_SaveAsAliasKeepsDestinationAfterAtomicReplacement()
    {
        WithDirectory((root, actual, alias) =>
        {
            var source = Path.Combine(actual, "source.xlsx");
            var destination = Path.Combine(root, "destination.xlsx");
            var temporary = Path.Combine(root, "replacement.xlsx");
            File.WriteAllText(source, "synthetic-original");
            Link(destination, source);
            var claims = new WorkbookPathReservations<string>();
            Assert.True(claims.TryAdd(source, "session"));
            Assert.True(claims.TryGetValue(destination, out var owner));
            Assert.Equal("session", owner);
            claims[destination] = owner;
            File.WriteAllText(temporary, "synthetic-replacement");
            File.Move(temporary, destination, overwrite: true);
            Assert.True(claims.TryRemove(source, out _));
            Assert.True(claims.TryAdd(source, "other-source"));
            Assert.False(claims.TryAdd(destination, "other"));
            Assert.True(claims.TryGetValue(destination, out owner));
            Assert.Equal("session", owner);
            Assert.True(claims.TryRemove(new(destination, "session")));
            Assert.True(claims.TryAdd(destination, "other"));
        });
    }

    [Fact]
    public void Reservations_CompareAndRemoveCannotDropTransferredOrForeignOwner()
    {
        WithDirectory((root, actual, alias) =>
        {
            var source = Path.Combine(actual, "source.xlsx");
            var linked = Path.Combine(root, "linked.xlsx");
            File.WriteAllText(source, "synthetic");
            Link(linked, source);
            var claims = new WorkbookPathReservations<string>();
            Assert.True(claims.TryAdd(source, "pending"));
            Assert.True(claims.TryUpdate(source, "session", "pending"));
            Assert.False(claims.TryRemove(new(source, "pending")));
            Assert.False(claims.TryRemove(new(linked, "foreign")));
            Assert.Throws<InvalidOperationException>(() => claims[linked] = "foreign");
            Assert.True(claims.TryGetValue(linked, out var owner));
            Assert.Equal("session", owner);
            Assert.True(claims.TryRemove(new(linked, "session")));
            Assert.Empty(claims);
        });
    }

    [Fact]
    public void Reservations_RemoteQueriesRemainSeparateAndPrivate()
    {
        var first = WorkbookLocation.GetOwnershipKey("https://example.com/a.xlsx?token=one");
        var second = WorkbookLocation.GetOwnershipKey("https://example.com/a.xlsx?token=two");
        var claims = new WorkbookPathReservations<string>();
        Assert.True(claims.TryAdd(first, "one"));
        Assert.True(claims.TryAdd(second, "two"));
        Assert.True(claims.TryGetValue(first, out var owner));
        Assert.Equal("one", owner);
        Assert.DoesNotContain("token=", string.Join(',', claims.Select(entry => entry.Key)));
    }

    [Fact]
    public void OwnershipAliases_UnresolvedJunctionAndFileAncestorFailClosed()
    {
        WithDirectory((root, actual, alias) =>
        {
            var file = Path.Combine(actual, "not-a-directory");
            File.WriteAllText(file, "synthetic");
            Assert.Throws<IOException>(() => WorkbookLocation.GetOwnershipKey(Path.Combine(file, "new.xlsx")));
            File.Delete(file);
            Directory.Delete(actual);
            Assert.Throws<IOException>(() => WorkbookLocation.GetOwnershipKey(Path.Combine(alias, "new.xlsx")));
        });
    }

    [Fact]
    public void OwnershipAliases_JunctionUsesSameKeyBeforeAndAfterFileCreation()
    {
        WithDirectory((root, actual, alias) =>
        {
            var target = Path.Combine(actual, "new.xlsx");
            var viaAlias = Path.Combine(alias, "new.xlsx");
            var reserved = WorkbookLocation.GetOwnershipKey(target);
            Assert.Equal(reserved, WorkbookLocation.GetOwnershipKey(viaAlias));
            File.WriteAllText(target, "synthetic");
            Assert.Equal(reserved, WorkbookLocation.GetOwnershipKey(target));
            Assert.Equal(reserved, WorkbookLocation.GetOwnershipKey(viaAlias));
            Assert.True(WorkbookLocation.MatchesReportedLocation(reserved, viaAlias));
        });
    }

    [Fact]
    public void OwnershipAliases_MissingNestedDestinationUsesPhysicalAncestor()
    {
        WithDirectory((root, actual, alias) =>
        {
            Assert.Equal(WorkbookLocation.GetOwnershipKey(Path.Combine(actual, "missing", "new.xlsx")),
                WorkbookLocation.GetOwnershipKey(Path.Combine(alias, "missing", "new.xlsx")));
        });
    }

    private static void WithDirectory(Action<string, string, string> body)
    {
        var root = Path.Combine(Path.GetTempPath(), "excelmcp-physical-" + Guid.NewGuid().ToString("N"));
        var actual = Path.Combine(root, "actual");
        var alias = Path.Combine(root, "junction");
        Directory.CreateDirectory(actual);
        var primary = Record.Exception(() =>
        {
            CreateJunction(alias, actual);
            body(root, actual, alias);
        });
        var cleanup = Record.Exception(() =>
        {
            // Remove only the reparse entry, then this test's unique temporary root.
            if (new DirectoryInfo(alias).LinkTarget is not null) Directory.Delete(alias);
            Directory.Delete(root, recursive: true);
        });
        if (primary is not null && cleanup is not null) throw new AggregateException(primary, cleanup);
        if (primary is not null || cleanup is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture((primary ?? cleanup)!).Throw();
    }

    private static void CreateJunction(string alias, string target)
    {
        Directory.CreateDirectory(alias);
        using var handle = CreateFile(alias, 0x40000000, 7, IntPtr.Zero, 3, 0x02200000, IntPtr.Zero);
        if (handle.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error());
        var substitute = Encoding.Unicode.GetBytes(@"\??\" + target);
        var print = Encoding.Unicode.GetBytes(target);
        var data = new byte[16 + substitute.Length + 2 + print.Length + 2];
        BitConverter.GetBytes(0xA0000003u).CopyTo(data, 0);
        BitConverter.GetBytes((ushort)(data.Length - 8)).CopyTo(data, 4);
        BitConverter.GetBytes((ushort)substitute.Length).CopyTo(data, 10);
        BitConverter.GetBytes((ushort)(substitute.Length + 2)).CopyTo(data, 12);
        BitConverter.GetBytes((ushort)print.Length).CopyTo(data, 14);
        substitute.CopyTo(data, 16);
        print.CopyTo(data, 18 + substitute.Length);
        if (!DeviceIoControl(handle, 0x000900A4, data, data.Length, IntPtr.Zero, 0, out _, IntPtr.Zero))
            throw new Win32Exception(Marshal.GetLastWin32Error());
    }

    private static void Link(string alias, string original)
    {
        if (!CreateHardLink(alias, original, IntPtr.Zero)) throw new Win32Exception(Marshal.GetLastWin32Error());
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CreateHardLink(string alias, string original, IntPtr security);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeFileHandle CreateFile(string path, uint access, uint share, IntPtr security, uint mode, uint flags, IntPtr template);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DeviceIoControl(SafeFileHandle handle, uint code, byte[] input, int inputLength, IntPtr output, int outputLength, out int returned, IntPtr overlapped);
}
