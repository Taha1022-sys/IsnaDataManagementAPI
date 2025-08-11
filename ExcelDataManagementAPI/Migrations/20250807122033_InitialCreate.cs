using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace ExcelDataManagementAPI.Migrations
{
    /// <inheritdoc />
    public partial class InitialCreate : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "ExcelDataRows",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    FileName = table.Column<string>(type: "nvarchar(255)", maxLength: 255, nullable: false),
                    SheetName = table.Column<string>(type: "nvarchar(255)", maxLength: 255, nullable: false),
                    RowIndex = table.Column<int>(type: "int", nullable: false),
                    RowData = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    CreatedDate = table.Column<DateTime>(type: "datetime2", nullable: false),
                    ModifiedDate = table.Column<DateTime>(type: "datetime2", nullable: true),
                    IsDeleted = table.Column<bool>(type: "bit", nullable: false),
                    Version = table.Column<int>(type: "int", nullable: false),
                    ModifiedBy = table.Column<string>(type: "nvarchar(255)", maxLength: 255, nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ExcelDataRows", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "ExcelFiles",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    FileName = table.Column<string>(type: "nvarchar(255)", maxLength: 255, nullable: false),
                    OriginalFileName = table.Column<string>(type: "nvarchar(255)", maxLength: 255, nullable: false),
                    FilePath = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: false),
                    FileSize = table.Column<long>(type: "bigint", nullable: false),
                    UploadDate = table.Column<DateTime>(type: "datetime2", nullable: false),
                    UploadedBy = table.Column<string>(type: "nvarchar(255)", maxLength: 255, nullable: true),
                    IsActive = table.Column<bool>(type: "bit", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ExcelFiles", x => x.Id);
                });

            migrationBuilder.CreateIndex(
                name: "IX_ExcelDataRows_CreatedDate",
                table: "ExcelDataRows",
                column: "CreatedDate");

            migrationBuilder.CreateIndex(
                name: "IX_ExcelDataRows_FileName_SheetName",
                table: "ExcelDataRows",
                columns: new[] { "FileName", "SheetName" });

            migrationBuilder.CreateIndex(
                name: "IX_ExcelDataRows_IsDeleted",
                table: "ExcelDataRows",
                column: "IsDeleted");

            migrationBuilder.CreateIndex(
                name: "IX_ExcelDataRows_ModifiedDate",
                table: "ExcelDataRows",
                column: "ModifiedDate");

            migrationBuilder.CreateIndex(
                name: "IX_ExcelDataRows_RowIndex",
                table: "ExcelDataRows",
                column: "RowIndex");

            migrationBuilder.CreateIndex(
                name: "IX_ExcelFiles_FileName",
                table: "ExcelFiles",
                column: "FileName");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "ExcelDataRows");

            migrationBuilder.DropTable(
                name: "ExcelFiles");
        }
    }
}
