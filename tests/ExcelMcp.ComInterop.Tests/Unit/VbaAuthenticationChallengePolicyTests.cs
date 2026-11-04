using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("RequiresExcel", "false")]
public sealed class VbaAuthenticationChallengePolicyTests
{
    [Theory]
    [InlineData("VBAProject Password", true)]
    [InlineData("OtherProject Password", false)]
    [InlineData("Password", false)]
    [InlineData("VBAProject Password - External", false)]
    [InlineData("VBAProject password", false)]
    public void MaskedLookalikeRequiresExactTargetCaption(string caption, bool expected)
    {
        Assert.Equal(expected, VbaAuthenticationChallengePolicy.CanRecognize("VBAProject", caption,
            true, false, true, 1, true, 5470, 2, true, true, false));
    }

    [Theory]
    [InlineData(false, false, true)]
    [InlineData(true, true, true)]
    [InlineData(true, false, false)]
    public void MatchingChallengeCannotBeAdoptedOutsideCommittedOwnedCommand(
        bool committed, bool preexisting, bool owned)
    {
        Assert.False(VbaAuthenticationChallengePolicy.CanRecognize("VBAProject", "VBAProject Password",
            committed, preexisting, owned, 1, true, 5470, 2, true, true, false));
    }

    [Theory]
    [InlineData(2, true, 5470, 2, true, true, false)]
    [InlineData(1, false, 5470, 2, true, true, false)]
    [InlineData(1, true, 5471, 2, true, true, false)]
    [InlineData(1, true, 0, 2, true, true, false)]
    [InlineData(1, true, 5470, 3, true, true, false)]
    [InlineData(1, true, 5470, 2, false, true, false)]
    [InlineData(1, true, 5470, 2, true, false, false)]
    [InlineData(1, true, 5470, 2, true, true, true)]
    public void UnexpectedControlSignatureCannotReceivePassword(int edits, bool masked, int editId,
        int buttons, bool ok, bool cancel, bool properties)
    {
        Assert.False(VbaAuthenticationChallengePolicy.CanRecognize("VBAProject", "VBAProject Password",
            true, false, true, edits, masked, editId, buttons, ok, cancel, properties));
    }

    [Theory]
    [InlineData(20, 20, true)]
    [InlineData(20, 10, false)]
    [InlineData(20, 30, false)]
    [InlineData(0, 0, false)]
    public void InvalidPasswordRequiresDirectChallengeOwner(int challenge, int directOwner, bool expected)
    {
        Assert.Equal(expected, VbaAuthenticationChallengePolicy.IsInvalidPasswordOwner(challenge, directOwner));
    }

    [Theory]
    [InlineData("VBAProject - Project Properties", true, false, 10, 10, 1, true)]
    [InlineData("VBAProject - Project Properties", true, false, 10, 20, 1, false)]
    [InlineData("VBAProject - Project Properties", true, true, 10, 10, 1, false)]
    [InlineData("VBAProject - Project Properties", false, false, 10, 10, 1, false)]
    [InlineData("VBAProject - Project Properties", true, false, 10, 10, 2, false)]
    [InlineData("VBAProject - Project Properties", true, false, 0, 0, 1, false)]
    [InlineData("OtherProject - Project Properties", true, false, 10, 10, 1, false)]
    [InlineData("VBAProject - Unrelated", true, false, 10, 10, 1, false)]
    public void PropertiesRequiresExactUniqueReplacementTransition(string caption, bool submitted,
        bool challengeAlive, int capturedEditorOwner, int directOwner, int modalCount, bool expected)
    {
        Assert.Equal(expected, VbaAuthenticationChallengePolicy.IsPropertiesReplacement("VBAProject", caption,
            submitted, challengeAlive, capturedEditorOwner, directOwner, modalCount));
    }

    [Theory]
    [InlineData(false, true, false)]
    [InlineData(true, false, false)]
    [InlineData(true, true, true)]
    public void CleanupNeverAdoptsUnrecordedLookalike(bool recorded, bool signature, bool expected)
    {
        Assert.Equal(expected, VbaAuthenticationChallengePolicy.CanCleanUp(recorded, signature));
    }
}
