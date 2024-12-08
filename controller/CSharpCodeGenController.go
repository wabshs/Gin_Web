package controller

import (
	"bytes"
	"io/ioutil"
	"mime/multipart"
	"net/http"

	"Go3/service"
	"github.com/gin-gonic/gin"
	"github.com/xuri/excelize/v2"
)

type CSharpGeneratorController struct {
	Service *service.CSharpGeneratorService
}

// NewCSharpGeneratorController initializes a new controller with the service.
func NewCSharpGeneratorController(svc *service.CSharpGeneratorService) *CSharpGeneratorController {
	return &CSharpGeneratorController{Service: svc}
}

// GetSheets 读取Excel
func (c *CSharpGeneratorController) GetSheets(ctx *gin.Context) {
	// Parse the file from the form data.
	file, _, err := ctx.Request.FormFile("file")
	if err != nil {
		ctx.JSON(http.StatusBadRequest, gin.H{"error": "Invalid file"})
		return
	}
	defer func(file multipart.File) {
		err := file.Close()
		if err != nil {

		}
	}(file)

	// Read the file into a byte slice.
	fileData, err := ioutil.ReadAll(file)
	if err != nil {
		ctx.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to read file"})
		return
	}

	// Open the Excel file from the byte slice (stream).
	f, err := excelize.OpenReader(bytes.NewReader(fileData))
	if err != nil {
		ctx.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to open Excel file"})
		return
	}

	// Get all sheet names.
	sheetNames := f.GetSheetList()

	// Return the sheet names as a response.
	ctx.JSON(http.StatusOK, gin.H{"sheets": sheetNames})
}

// Generate handles the HTTP request to generate a C# class from an Excel file.
func (c *CSharpGeneratorController) Generate(ctx *gin.Context) {
	// Parse the file from the form data.
	file, _, err := ctx.Request.FormFile("file")
	if err != nil {
		ctx.JSON(http.StatusBadRequest, gin.H{"error": "Invalid file"})
		return
	}
	defer func(file multipart.File) {
		err := file.Close()
		if err != nil {

		}
	}(file)

	// Read the file into a byte slice.
	fileData, err := ioutil.ReadAll(file)
	if err != nil {
		ctx.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to read file"})
		return
	}

	// Get the sheet name from the request form.
	sheetName := ctx.PostForm("sheet_name")
	if sheetName == "" {
		ctx.JSON(http.StatusBadRequest, gin.H{"error": "Sheet name is required"})
		return
	}

	// Generate the C# class from the file data.
	csharpClass, err := c.Service.GenerateCSharpClassFromExcel(fileData, sheetName)
	if err != nil {
		ctx.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to generate C# class", "details": err.Error()})
		return
	}

	// Return the generated class as a response.
	ctx.JSON(http.StatusOK, gin.H{"CS_Code": csharpClass})
}
