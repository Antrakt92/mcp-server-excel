using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Generated;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "VBA")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class VbaProjectAuthenticationResultTests
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web)
    {
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
    };

    [Theory]
    [InlineData(null)]
    [InlineData(false)]
    [InlineData(true)]
    public void VerificationOutcome_RemainsDistinctFromSuccessfulOperation(bool? verified)
    {
        var value = new VbaProjectAuthenticationResult
        {
            Success = true,
            PasswordVerified = verified,
            AlreadyAccessible = !verified.HasValue,
            AuthenticationPerformed = verified.HasValue,
            AccessUnlocked = verified is true,
            ProjectName = "SyntheticProject",
            ProtectionBefore = verified.HasValue ? "Locked" : "None",
            ProtectionAfter = verified is false ? "Locked" : "None"
        };
        var json = JsonSerializer.Serialize(value, JsonOptions);
        var result = ResultJson.Read(json, "vba.verify-password",
            contract: ServiceRegistry.GetResultContract("vba.verify-password"));
        Assert.True(result.Success);
        using var document = JsonDocument.Parse(result.ToCliJson("vba.verify-password", 0));
        var actual = document.RootElement.GetProperty("data").GetProperty("passwordVerified");
        if (verified.HasValue) Assert.Equal(verified.Value, actual.GetBoolean());
        else Assert.Equal(JsonValueKind.Null, actual.ValueKind);
        Assert.False(document.RootElement.GetProperty("data").GetProperty("saved").GetBoolean());
        Assert.False(document.RootElement.GetProperty("data").GetProperty("reopened").GetBoolean());
    }
}
