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
