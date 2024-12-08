package service

import (
	"bytes"
	"fmt"
	"github.com/xuri/excelize/v2"
	"golang.org/x/text/cases"
	"golang.org/x/text/language"
	"strings"
)

type CSharpGeneratorService struct{}

// GenerateCSharpClassFromExcel reads an Excel file from a stream and returns a C# class definition as a string.
func (s *CSharpGeneratorService) GenerateCSharpClassFromExcel(fileData []byte, sheetName string) (string, error) {
	// Open the Excel file from the byte slice (stream).
	f, err := excelize.OpenReader(bytes.NewReader(fileData))
	if err != nil {
		return "", err
	}

	// Get rows from the specified sheet.
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return "", err
	}

	var sb strings.Builder

	// Assuming the table name is in the first cell of the first row.
	tableName := rows[0][0]
	sb.WriteString(fmt.Sprintf("```csharp\npublic class %s\n{\n", strings.Title(tableName)))

	// Skip the first row which contains the table name.
	for _, row := range rows[1:] {
		if len(row) >= 2 { // Ensure there are at least two columns (name and type).
			sb.WriteString(s.generateCSharpProperty(row[0], row[1]))
		}
	}

	sb.WriteString("}\n```")
	return sb.String(), nil
}

func (s *CSharpGeneratorService) generateCSharpProperty(colName, colType string) string {
	// Convert the column type to C# equivalent if necessary.
	csColType := s.mapCSharpType(colType)

	// Use cases.Title for capitalization
	caser := cases.Title(language.Und) // Use "Und" (undefined) language for generic capitalization
	capitalizedColName := caser.String(colName)

	return fmt.Sprintf("\tpublic %s %s { get; set; }\n", csColType, capitalizedColName)
}

func (s *CSharpGeneratorService) mapCSharpType(typeName string) string {
	switch strings.ToLower(typeName) {
	case "int":
		return "int"
	case "string":
		return "string"
	case "datetime":
		return "DateTime"
	default:
		return "object" // Fallback type
	}
}

// GenerateCSharpInsertMethodFromExcel Excel生成Insert语句
func (s *CSharpGeneratorService) GenerateCSharpInsertMethodFromExcel(fileData []byte, sheetName string) (string, error) {
	f, err := excelize.OpenReader(bytes.NewReader(fileData))
	if err != nil {
		return "", err
	}
	defer func(f *excelize.File) {
		err := f.Close()
		if err != nil {

		}
	}(f)

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return "", err
	}

	// Assuming the class name is in the first cell of the first row.
	tableName := strings.Title(rows[0][0])
	var sb strings.Builder

	// Start building the C# method.
	sb.WriteString(fmt.Sprintf("public static void Insert(%s obj, SqlTransaction trans)\n{\n", tableName))
	sb.WriteString("    string strSql = @\"insert into IBP_ManagementList(\n")

	// Collect column names and prepare SQL parameter list.
	var columns []string
	var params []string
	for _, row := range rows[1:] { // Skip the first row which contains the table name.
		if len(row) >= 2 { // Ensure there are at least two columns (name and type).
			columnName := row[0]
			columnType := s.mapExcelToSqlDbType(row[1])

			columns = append(columns, columnName)
			params = append(params, "@"+columnName)

			// Add SqlParameter initialization to the builder.
			sb.WriteString(fmt.Sprintf("        new SqlParameter(\"%s\", SqlDbType.%s),\n", columnName, columnType))
		}
	}

	// Complete the SQL command string.
	sb.WriteString(strings.Join(columns, ",") + ") values(")
	sb.WriteString(strings.Join(params, ",") + ")\";\n\n")

	// Initialize the SqlParameter array.
	sb.WriteString("    SqlParameter[] parameters = new SqlParameter[] {\n")
	for i, p := range params {
		if i > 0 {
			sb.WriteString(",\n")
		}
		sb.WriteString(fmt.Sprintf("        %s", p))
	}
	sb.WriteString("\n    };\n\n")

	// Assign values to parameters.
	sb.WriteString("    int i = -1;\n")
	for _, col := range columns {
		sb.WriteString(fmt.Sprintf("    parameters[++i].Value = obj.%s;\n", col))
	}

	// Execute the command.
	sb.WriteString("    if (trans != null)\n")
	sb.WriteString("        FXSZMIS.Data.SQLHelper.ExecuteNonQuery(trans, CommandType.Text, strSql, parameters);\n")
	sb.WriteString("    else\n")
	sb.WriteString(fmt.Sprintf("        FXSZMIS.Data.SQLHelper.ExecuteNonQuery(%s.Connection, CommandType.Text, strSql, parameters);\n", tableName))
	sb.WriteString("}\n")

	return sb.String(), nil
}

func (s *CSharpGeneratorService) mapExcelToSqlDbType(excelType string) string {
	switch strings.ToLower(excelType) {
	case "datetime":
		return "DateTime"
	case "varchar":
		return "VarChar"
	case "int":
		return "Int"
	default:
		return "VarChar" // Default to VarChar as a fallback.
	}
}
