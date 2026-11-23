using ClosedXML.Excel;
using IQS.ValidateItemMasterImportFile.Models;

namespace Validate_Master_Import;

public static class ItemMasterValidator
{
    private const int HeaderRow = 5;
    private const string WorksheetName = "Sheet1";

    private static readonly string[] ExpectedColumns =
    {
        "Item Number",
        "Description",
        "UOM",
        "Case To Bottle",
        "Case To Each",
        "Pallet Quantity",
        "Conv Unit of Measure",
        "NDC",
        "NSN",
        "List Price",
        "Standard Cost",
        "Lot Control",
        "Shelf Control",
        "Serial Control",
        "Agent Type",
        "Administrator",
        "Drug Type",
        "Drug Strength",
        "Rotatable",
        "SLEP Eligible",
        "Category I-V",
        "Indications",
        "Customer",
        "Adult/ped",
        "Pregnancy Class",
        "Product Identifier",
        "Item Type",
        "Item Status",
        "Temp Requirements",
        "Item Creation Date"
    };

    /// <summary>
    /// Validates the Item Master workbook and exports the valid item rows.
    /// Returns 0 on success and 1 on validation failure.
    /// </summary>
    public static int Run(string inputPath)
    {
        try
        {
            ItemMasterValidationResult validationResult = Validate(inputPath);

            if (!validationResult.IsValid)
            {
                string errorMessage = $"Validation failed with {validationResult.Errors.Count} error(s):\r\n\r\n" + string.Join("\r\n", validationResult.Errors);

                MessageBox.Show(
                    errorMessage,
                    "Validation Errors",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                return 1;
            }

            string exportPath = Export(inputPath);

            string successMessage =
                $"Successfully validated Excel file.\r\n\r\n" +
                $"Exported:\r\n{exportPath}";

            MessageBox.Show(
                successMessage,
                "Success",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            return 0;
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"An unexpected error occurred while processing the workbook.\r\n\r\n{ex.Message}",
                "Processing Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

            return 1;
        }
    }

    /// <summary>
    /// Validates the Item Master workbook structure.
    /// No files are created or modified by this method.
    /// </summary>
    public static ItemMasterValidationResult Validate(string inputPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);

        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException(
            "Input Excel file not found.",
            inputPath);
        }

        var errors = new List<string>();

        using var workbook = new XLWorkbook(inputPath);

        /*
        * Require the expected worksheet.
        *
        * Do not fall back to another worksheet because the worksheet
        * name is part of the expected Item Master format.
        */
        var worksheet = GetExpectedWorksheet(workbook);

        if (worksheet is null)
        {
            errors.Add(
            $"Expected worksheet '{WorksheetName}' was not found.");

            return new ItemMasterValidationResult(
            false,
            errors);
        }

        /*
        * Locate the expected header row.
        */
        var headerRow = worksheet.Row(HeaderRow);
        var lastHeaderCell = headerRow.LastCellUsed();

        if (lastHeaderCell is null)
        {
            errors.Add(
            $"No data was found in header row {HeaderRow}.");

            return new ItemMasterValidationResult(
            false,
            errors);
        }

        /*
        * Validate the total number of columns.
        *
        * This catches missing expected columns as well as unexpected
        * additional columns.
        */
        int actualColumnCount =
        lastHeaderCell.Address.ColumnNumber;

        if (actualColumnCount != ExpectedColumns.Length)
        {
            errors.Add(
            $"Expected {ExpectedColumns.Length} columns, " +
            $"but found {actualColumnCount} columns.");
        }

        /*
        * Validate each expected column name and position.
        *
        * Header comparisons are case-insensitive and leading/trailing
        * whitespace is ignored.
        */
        for (int i = 0; i < ExpectedColumns.Length; i++)
        {
            string expectedHeader = ExpectedColumns[i];

            string actualHeader =
            headerRow
            .Cell(i + 1)
            .GetString()
            .Trim();

            if (!string.Equals(
            actualHeader,
            expectedHeader,
            StringComparison.OrdinalIgnoreCase))
            {
                errors.Add(
                $"Column {i + 1}: Expected '{expectedHeader}' " +
                $"but found '{actualHeader}'.");
            }
        }

        return new ItemMasterValidationResult(
        errors.Count == 0,
        errors);
    }

    /// <summary>
    /// Exports valid Item Master rows to a new workbook.
    /// The input workbook is expected to have already passed validation.
    /// </summary>
    public static string Export(string inputPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);

        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException(
            "Input Excel file not found.",
            inputPath);
        }

        using var workbook = new XLWorkbook(inputPath);

        /*
        * Export should only be called after successful validation.
        * Failure to locate the expected worksheet here indicates a
        * programming or sequencing error rather than a validation result.
        */
        var worksheet = GetExpectedWorksheet(workbook)
        ?? throw new InvalidOperationException(
        $"Expected worksheet '{WorksheetName}' was not found.");

        int firstDataRow = HeaderRow + 1;

        int lastDataRow =
        worksheet.LastRowUsed()?.RowNumber()
        ?? (firstDataRow - 1);

        int itemNumberColumn =
        Array.FindIndex(
        ExpectedColumns,
        header => string.Equals(
        header,
        "Item Number",
        StringComparison.OrdinalIgnoreCase))
        + 1;

        if (itemNumberColumn <= 0)
        {
            throw new InvalidOperationException(
            "Could not locate the 'Item Number' column.");
        }

        string exportPath = BuildExportPath(inputPath);

        using var outputWorkbook = new XLWorkbook();

        var outputWorksheet =
        outputWorkbook.Worksheets.Add(WorksheetName);

        /*
        * Write the expected headers.
        */
        for (int column = 0;
        column < ExpectedColumns.Length;
        column++)
        {
            outputWorksheet.Cell(1, column + 1).Value =
            ExpectedColumns[column];
        }

        /*
        * Copy valid Item Master rows.
        *
        * Exclude:
        * - Rows with a blank Item Number
        * - Rows where Item Number is "CUI"
        */
        int outputRow = 2;

        for (int row = firstDataRow;
        row <= lastDataRow;
        row++)
        {
            string itemNumber =
            worksheet
            .Cell(row, itemNumberColumn)
            .GetString()
            .Trim();

            if (string.IsNullOrWhiteSpace(itemNumber))
            {
                continue;
            }

            if (string.Equals(
            itemNumber,
            "CUI",
            StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            for (int column = 1;
            column <= ExpectedColumns.Length;
            column++)
            {
                outputWorksheet
                .Cell(outputRow, column)
                .Value =
                worksheet
                .Cell(row, column)
                .Value;
            }

            outputRow++;
        }

        /*
        * Auto-size columns while preventing excessively narrow
        * or wide output columns.
        */
        int lastOutputRow =
        Math.Max(1, outputRow - 1);

        outputWorksheet
        .Columns(1, ExpectedColumns.Length)
        .AdjustToContents(
        1,
        lastOutputRow,
        minWidth: 6,
        maxWidth: 60);

        outputWorkbook.SaveAs(exportPath);

        return exportPath;
    }

    /// <summary>
    /// Gets the expected Item Master worksheet.
    /// </summary>
    private static IXLWorksheet? GetExpectedWorksheet(
    XLWorkbook workbook)
    {
        return workbook.Worksheets.FirstOrDefault(
        worksheet => string.Equals(
        worksheet.Name,
        WorksheetName,
        StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Builds the validated Item Master export path.
    /// </summary>
    private static string BuildExportPath(string inputPath)
    {
        string baseName =
        Path.GetFileNameWithoutExtension(inputPath);

        /*
        * Preserve the existing naming convention by using the
        * portion of the input filename before the first underscore.
        */
        string prefix =
        baseName
        .Split('_')[0]
        .Trim();

        string exportFileName =
        $"{prefix} Validated ItemMaster - New Items.xlsx";

        string exportDirectory =
        Path.GetDirectoryName(inputPath)
        ?? Environment.CurrentDirectory;

        return Path.Combine(
        exportDirectory,
        exportFileName);
    }
}
