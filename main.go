package main

import (
	"Go3/controller"
	"Go3/cors"

	"Go3/service"
	"github.com/gin-gonic/gin"
)

func main() {
	r := gin.Default()
	r.Use(cors.Cors())

	// Initialize service and controller.
	csharpService := &service.CSharpGeneratorService{}
	csharpController := controller.NewCSharpGeneratorController(csharpService)

	// Define a route group "CSClass"
	csClassGroup := r.Group("/CSharpClass") // 创建路由组
	{
		// Define API routes within the CSClass group.
		csClassGroup.POST("/generate", csharpController.Generate)    // C# 类生成
		csClassGroup.POST("/get_sheets", csharpController.GetSheets) // 获取工作表列表
	}

	// Start the server.
	err := r.Run(":8080")
	if err != nil {
		return
	} // Default listens on port 8080.
}
