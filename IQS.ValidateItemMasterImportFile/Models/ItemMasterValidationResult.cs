namespace IQS.ValidateItemMasterImportFile.Models;

public sealed record ItemMasterValidationResult(
    bool IsValid,
    IReadOnlyList<string> Errors);