using ClosedXML.Excel;
using Validate_Master_Import;

namespace IQS.ValidateItemMasterImportFile.Tests;

public sealed class ItemMasterValidatorTests
{
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

    [Fact]
    public void Validate_ValidWorkbook_ReturnsValid()
    {
        // Arrange
        string testDirectory = Path.Combine(
            Path.GetTempPath(),
            "IQS.ValidateItemMasterImportFile.Tests",
            Guid.NewGuid().ToString());

        Directory.CreateDirectory(testDirectory);

        string testFile = Path.Combine(
            testDirectory,
            "ValidItemMaster.xlsx");

        try
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");

                for (int column = 0; column < ExpectedColumns.Length; column++)
                {
                    worksheet.Cell(5, column + 1).Value =
                        ExpectedColumns[column];
                }

                workbook.SaveAs(testFile);
            }

            // Act
            var result = ItemMasterValidator.Validate(testFile);

            // Assert
            Assert.True(
                result.IsValid,
                $"Expected workbook to be valid. Errors: {string.Join(", ", result.Errors)}");

            Assert.Empty(result.Errors);
        }
        finally
        {
            if (Directory.Exists(testDirectory))
            {
                Directory.Delete(
                    testDirectory,
                    recursive: true);
            }
        }
    }

    [Fact]
    public void Validate_MissingWorksheet_ReturnsValidationError()
    {
        // Arrange
        string testFile = CreateTestWorkbook(
            "MissingWorksheet.xlsx",
            worksheetName: "WrongSheet");

        try
        {
            // Act
            var result = ItemMasterValidator.Validate(testFile);

            // Assert
            Assert.False(result.IsValid);

            Assert.Single(result.Errors);

            Assert.Equal(
                "Expected worksheet 'Sheet1' was not found.",
                result.Errors[0]);
        }
        finally
        {
            DeleteTestWorkbook(testFile);
        }
    }

    [Fact]
    public void Validate_EmptyHeaderRow_ReturnsValidationError()
    {
        // Arrange
        string testFile = CreateTestWorkbook(
            "EmptyHeaderRow.xlsx",
            includeExpectedHeaders: false,
            configureWorksheet: worksheet =>
            {
                worksheet.Cell(1, 1).Value = "Test";
            });

        try
        {
            // Act
            var result = ItemMasterValidator.Validate(testFile);

            // Assert
            Assert.False(result.IsValid);

            Assert.Single(result.Errors);

            Assert.Equal(
                "No data was found in header row 5.",
                result.Errors[0]);
        }
        finally
        {
            DeleteTestWorkbook(testFile);
        }
    }

    [Fact]
    public void Validate_MissingColumn_ReturnsValidationError()
    {
        // Arrange
        string testDirectory = Path.Combine(
            Path.GetTempPath(),
            "IQS.ValidateItemMasterImportFile.Tests",
            Guid.NewGuid().ToString());

        Directory.CreateDirectory(testDirectory);

        string testFile = Path.Combine(
            testDirectory,
            "MissingColumn.xlsx");

        try
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");

                // Intentionally omit the final expected column.
                for (int column = 0; column < ExpectedColumns.Length - 1; column++)
                {
                    worksheet.Cell(5, column + 1).Value =
                        ExpectedColumns[column];
                }

                workbook.SaveAs(testFile);
            }

            // Act
            var result = ItemMasterValidator.Validate(testFile);

            // Assert
            Assert.False(result.IsValid);

            Assert.Contains(
                result.Errors,
                error => error == "Expected 30 columns, but found 29 columns.");

            Assert.Contains(
                result.Errors,
                error => error == "Column 30: Expected 'Item Creation Date' but found ''.");
        }
        finally
        {
            if (Directory.Exists(testDirectory))
            {
                Directory.Delete(
                    testDirectory,
                    recursive: true);
            }
        }
    }

    [Fact]
    public void Validate_ExtraColumn_ReturnsValidationError()
    {
        // Arrange
        string testDirectory = Path.Combine(
            Path.GetTempPath(),
            "IQS.ValidateItemMasterImportFile.Tests",
            Guid.NewGuid().ToString());

        Directory.CreateDirectory(testDirectory);

        string testFile = Path.Combine(
            testDirectory,
            "ExtraColumn.xlsx");

        try
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");

                // Create all expected headers.
                for (int column = 0; column < ExpectedColumns.Length; column++)
                {
                    worksheet.Cell(5, column + 1).Value =
                        ExpectedColumns[column];
                }

                // Intentionally add an unexpected 31st column.
                worksheet.Cell(5, ExpectedColumns.Length + 1).Value =
                    "Unexpected Column";

                workbook.SaveAs(testFile);
            }

            // Act
            var result = ItemMasterValidator.Validate(testFile);

            // Assert
            Assert.False(result.IsValid);

            Assert.Single(result.Errors);

            Assert.Equal(
                "Expected 30 columns, but found 31 columns.",
                result.Errors[0]);
        }
        finally
        {
            if (Directory.Exists(testDirectory))
            {
                Directory.Delete(
                    testDirectory,
                    recursive: true);
            }
        }
    }

    [Fact]
    public void Validate_IncorrectHeader_ReturnsValidationError()
    {
        // Arrange
        string testFile = CreateTestWorkbook(
            "IncorrectHeader.xlsx",
            configureWorksheet: worksheet =>
            {
                worksheet.Cell(5, 2).Value = "Item Description";
            });

        try
        {
            // Act
            var result = ItemMasterValidator.Validate(testFile);

            // Assert
            Assert.False(result.IsValid);

            Assert.Single(result.Errors);

            Assert.Equal(
                "Column 2: Expected 'Description' but found 'Item Description'.",
                result.Errors[0]);
        }
        finally
        {
            DeleteTestWorkbook(testFile);
        }
    }

    [Fact]
    public void Validate_ColumnsOutOfOrder_ReturnsValidationError()
    {
        // Arrange
        string testFile = CreateTestWorkbook(
            "ColumnsOutOfOrder.xlsx",
            configureWorksheet: worksheet =>
            {
                worksheet.Cell(5, 2).Value = "UOM";
                worksheet.Cell(5, 3).Value = "Description";
            });

        try
        {
            // Act
            var result = ItemMasterValidator.Validate(testFile);

            // Assert
            Assert.False(result.IsValid);

            Assert.Equal(2, result.Errors.Count);

            Assert.Contains(
                "Column 2: Expected 'Description' but found 'UOM'.",
                result.Errors);

            Assert.Contains(
                "Column 3: Expected 'UOM' but found 'Description'.",
                result.Errors);
        }
        finally
        {
            DeleteTestWorkbook(testFile);
        }
    }

    [Fact]
    public void Validate_HeaderCaseDifference_IsAccepted()
    {
        // Arrange
        string testFile = CreateTestWorkbook(
            "HeaderCaseDifference.xlsx",
            configureWorksheet: worksheet =>
            {
                worksheet.Cell(5, 1).Value = "ITEM NUMBER";
                worksheet.Cell(5, 2).Value = "description";
                worksheet.Cell(5, 3).Value = "uom";
                worksheet.Cell(5, 30).Value = "ITEM CREATION DATE";
            });

        try
        {
            // Act
            var result = ItemMasterValidator.Validate(testFile);

            // Assert
            Assert.True(
                result.IsValid,
                $"Expected workbook to be valid. Errors: {string.Join(", ", result.Errors)}");

            Assert.Empty(result.Errors);
        }
        finally
        {
            DeleteTestWorkbook(testFile);
        }
    }

    [Fact]
    public void Validate_HeaderWhitespace_IsIgnored()
    {
        // Arrange
        string testFile = CreateTestWorkbook(
            "HeaderWhitespace.xlsx",
            configureWorksheet: worksheet =>
            {
                worksheet.Cell(5, 1).Value = "  Item Number";
                worksheet.Cell(5, 2).Value = "Description  ";
                worksheet.Cell(5, 3).Value = "  UOM  ";
                worksheet.Cell(5, 30).Value = " Item Creation Date ";
            });

        try
        {
            // Act
            var result = ItemMasterValidator.Validate(testFile);

            // Assert
            Assert.True(
                result.IsValid,
                $"Expected workbook to be valid. Errors: {string.Join(", ", result.Errors)}");

            Assert.Empty(result.Errors);
        }
        finally
        {
            DeleteTestWorkbook(testFile);
        }
    }

    // Refactor Helper Methods:
    private static string CreateTestWorkbook(
    string fileName,
    string worksheetName = "Sheet1",
    Action<IXLWorksheet>? configureWorksheet = null,
    bool includeExpectedHeaders = true)
    {
        string testDirectory = Path.Combine(
            Path.GetTempPath(),
            "IQS.ValidateItemMasterImportFile.Tests",
            Guid.NewGuid().ToString());

        Directory.CreateDirectory(testDirectory);

        string testFile = Path.Combine(
            testDirectory,
            fileName);

        using var workbook = new XLWorkbook();

        var worksheet = workbook.Worksheets.Add(worksheetName);

        if (includeExpectedHeaders)
        {
            for (int column = 0; column < ExpectedColumns.Length; column++)
            {
                worksheet.Cell(5, column + 1).Value =
                    ExpectedColumns[column];
            }
        }

        configureWorksheet?.Invoke(worksheet);

        workbook.SaveAs(testFile);

        return testFile;
    }

    private static void DeleteTestWorkbook(string testFile)
    {
        string? testDirectory = Path.GetDirectoryName(testFile);

        if (testDirectory is not null &&
            Directory.Exists(testDirectory))
        {
            Directory.Delete(
                testDirectory,
                recursive: true);
        }
    }
}