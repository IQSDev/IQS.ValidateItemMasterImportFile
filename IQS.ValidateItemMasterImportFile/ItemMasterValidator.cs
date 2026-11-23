using ClosedXML.Excel;

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
            var (valid, errors, exportPath) = ValidateAndExport(inputPath);

            if (valid)
            {
                string message =
                    $"Successfully validated Excel file.\r\n\r\n" +
                    $"Exported:\r\n{exportPath}";

                MessageBox.Show(
                    message,
                    "Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                return 0;
            }

            string errorMessage =
                $"Validation failed with {errors.Count} error(s):\r\n\r\n" +
                string.Join("\r\n", errors);

            MessageBox.Show(
                errorMessage,
                "Validation Errors",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

            return 1;
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
    /// Validates the expected worksheet and column structure.
    /// If validation succeeds, exports rows containing valid Item Number values.
    /// </summary>
    public static (
        bool valid,
        List<string> errors,
        string exportPath)
        ValidateAndExport(string inputPath)
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
         * We intentionally do not fall back to the first worksheet because
         * the worksheet name is part of the expected Item Master format.
         */
        var worksheet = workbook.Worksheets.FirstOrDefault(
            w => string.Equals(
                w.Name,
                WorksheetName,
                StringComparison.OrdinalIgnoreCase));

        if (worksheet is null)
        {
            errors.Add(
                $"Expected worksheet '{WorksheetName}' was not found.");

            return (false, errors, string.Empty);
        }

        /*
         * Validate the expected header row.
         */
        var headerRow = worksheet.Row(HeaderRow);
        var lastHeaderCell = headerRow.LastCellUsed();

        if (lastHeaderCell is null)
        {
            errors.Add(
                $"No data was found in header row {HeaderRow}.");

            return (false, errors, string.Empty);
        }

        /*
         * Validate the number of columns.
         *
         * This catches both missing expected columns and unexpected
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
                headerRow.Cell(i + 1)
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

        /*
         * Do not create an export if the workbook structure is invalid.
         */
        if (errors.Count > 0)
        {
            return (false, errors, string.Empty);
        }

        /*
         * Determine the boundaries of the source data.
         */
        int firstDataRow = HeaderRow + 1;

        int lastDataRow =
            worksheet.LastRowUsed()?.RowNumber()
            ?? (firstDataRow - 1);

        /*
         * Locate Item Number using the expected header definition instead
         * of hard-coding its column number.
         */
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
            errors.Add(
                "Could not locate the 'Item Number' column.");

            return (false, errors, string.Empty);
        }

        /*
         * Build the output filename.
         *
         * This preserves the original PowerShell behavior of using the
         * portion of the source filename before the first underscore.
         */
        string baseName =
            Path.GetFileNameWithoutExtension(inputPath);

        string prefix =
            baseName
                .Split('_')[0]
                .Trim();

        string exportFileName =
            $"{prefix} Validated ItemMaster - New Items.xlsx";

        string exportDirectory =
            Path.GetDirectoryName(inputPath)
            ?? Environment.CurrentDirectory;

        string exportPath =
            Path.Combine(
                exportDirectory,
                exportFileName);

        /*
         * Create the validated export workbook.
         */
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
         *   - Rows with a blank Item Number
         *   - Rows where Item Number is "CUI"
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
         * Adjust exported column widths while preventing excessively
         * narrow or wide columns.
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

        return (
            true,
            errors,
            exportPath);
    }
}