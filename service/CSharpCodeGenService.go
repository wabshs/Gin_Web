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

func (s *CSharpGeneratorService) GenerateCSharpInsertMethodFromExcel(fileData []byte, sheetName string) (string, error) {
	f, err := excelize.OpenReader(bytes.NewReader(fileData))
	if err != nil {
		return "", err
	}
	defer func(f *excelize.File) {
		if closeErr := f.Close(); closeErr != nil {
			// 这里可以记录错误日志
		}
	}(f)

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return "", err
	}

	// 获取表名并将其格式化为首字母大写
	tableName := strings.Title(rows[0][0])
	var sb strings.Builder

	// 开始构建 C# 方法
	sb.WriteString("```csharp\n")
	sb.WriteString(fmt.Sprintf("public static void Add(%s obj, SqlTransaction trans)\n{\n", tableName))

	// SQL 语句及列名
	sb.WriteString(fmt.Sprintf("    string strSql = @\"INSERT INTO %s (\n", tableName))

	var params []string
	for i, row := range rows[1:] { // 跳过第一行
		if len(row) >= 2 { // 确保至少有两列
			columnName := row[0]
			// 添加参数到 SQL 插入语句
			params = append(params, fmt.Sprintf("@%s", columnName))

			// 在最后一行前面添加逗号
			if i != len(rows[1:])-1 {
				sb.WriteString(fmt.Sprintf("        %s,\n", columnName))
			} else {
				sb.WriteString(fmt.Sprintf("        %s\n", columnName))
			}
		}
	}

	sb.WriteString(") VALUES (" + strings.Join(params, ", ") + ")\";\n\n")

	// 添加参数设置
	sb.WriteString("    SqlParameter[] params = new SqlParameter[] {\n")
	for _, row := range rows[1:] { // 跳过第一行
		if len(row) >= 2 {
			columnName := row[0]
			columnType := s.mapExcelToSqlDbType(row[1]) // 使用 mapExcelToSqlDbType 函数
			sb.WriteString(fmt.Sprintf("        new SqlParameter(\"%s\", SqlDbType.%s, 50),\n", columnName, columnType))
		}
	}
	sb.WriteString("    };\n\n")

	// 添加参数赋值
	sb.WriteString("    int i = -1;\n")
	for _, row := range rows[1:] { // 跳过第一行
		if len(row) >= 2 {
			columnName := row[0]
			sb.WriteString(fmt.Sprintf("    params[++i].Value = obj.%s;\n", columnName))
		}
	}

	// 添加执行 SQL 的代码
	sb.WriteString("    if (trans != null)\n")
	sb.WriteString("        FXSZMIS.Data.SQLHelper.ExecuteNonQuery(trans, CommandType.Text, strSql, params);\n")
	sb.WriteString("    else\n")
	sb.WriteString(fmt.Sprintf("        FXSZMIS.Data.SQLHelper.ExecuteNonQuery(%s.Connection, CommandType.Text, strSql, params);\n", tableName))
	sb.WriteString("}\n```\n")
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
