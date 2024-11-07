package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"time"
)

func ExportToExcel(data []map[string]interface{}, fileName, sheetName string, columnOrder []string) error {
	f := excelize.NewFile()
	index, _ := f.NewSheet(sheetName)
	f.SetActiveSheet(index)
	if sheetName != "Sheet1" {
		_ = f.DeleteSheet("Sheet1")
	}
	if len(data) == 0 {
		log.Fatalln("data is empty!")
	}

	// 创建居中对齐的样式
	style, err := f.NewStyle(&excelize.Style{Alignment: &excelize.Alignment{Horizontal: "center"}})
	if err != nil {
		return err
	}

	// 写入表头并设置样式和列宽
	for i, header := range columnOrder {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		_ = f.SetCellValue(sheetName, cell, header)
		_ = f.SetCellStyle(sheetName, cell, cell, style)
		col := string(rune('A' + i))
		_ = f.SetColWidth(sheetName, col, col, float64(len(header)+5))
	}

	// 写入数据
	for rowIndex, row := range data {
		for colIndex, header := range columnOrder {
			cell, _ := excelize.CoordinatesToCellName(colIndex+1, rowIndex+2)
			_ = f.SetCellValue(sheetName, cell, row[header])
		}
	}

	// 保存文件
	if err := f.SaveAs(fileName); err != nil {
		return err
	}
	log.Printf("save excel in : %s\n", fileName)
	return nil
}

func main() {
	// 定义命令行参数
	jsonPath := flag.String("jp", "", "Path to the JSON file")
	filePath := flag.String("sp", fmt.Sprintf("%d.xlsx", time.Now().Unix()), "Path to save the Excel file")
	sheetName := flag.String("s", "Sheet1", "Sheet name")
	columnOrderStr := flag.String("ord", "", "Column order")

	// 解析命令行参数
	flag.Parse()

	// 读取和解析 JSON 文件
	fileContent, err := os.ReadFile(*jsonPath)
	if err != nil {
		log.Println("Error reading JSON file:", err)
		return
	}

	var data []map[string]interface{}
	if err := json.Unmarshal(fileContent, &data); err != nil {
		log.Println("Error parsing data:", err)
		return
	}

	// 解析列顺序
	var columnOrder []string
	if err := json.Unmarshal([]byte(*columnOrderStr), &columnOrder); err != nil {
		log.Println("Error parsing column order:", err)
		return
	}
	start := time.Now()
	// 调用函数生成 Excel 文件
	if err := ExportToExcel(data, *filePath, *sheetName, columnOrder); err != nil {
		log.Println("Error:", err)
	} else {
		log.Println("Excel file created successfully.")
	}
	elapsed := time.Since(start)
	log.Printf("Cost time: %s\n", elapsed)
}

/*
Call Demo

./lst2xlsx_mac -jp ./data/lst.json -sp ./data/data1.xlsx -s sheet2 -ord '["name", "addr", "age"]'

*/
