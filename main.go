package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"regexp"
	"strconv"
	"strings"
)

var sheetName = "打卡时间"
var before = 0
var after = 0
var delay = 0

var filename string

func init() {
	fmt.Print("Enter your filename: ")
	_, err := fmt.Scanln(&filename)
	if err != nil {
		return
	}
}

func main() {
	f, err := excelize.OpenFile(filename + ".xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := f.SaveAs("result.xlsx"); err != nil {
			fmt.Println(err)
		}
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	rows, err := f.GetRows(sheetName)
	if err != nil {
		fmt.Println(err)
		return
	}
	setTitle(f)
	for i, row := range rows {
		if i == 0 {
			continue
		}
		delay, after, before = 0, 0, 0
		for j, colCell := range row {
			cellName, _ := excelize.CoordinatesToCellName(j+1, i+1)
			if colCell == "" || isFullChinese(colCell) {
				setColor(f, cellName, "FFFF00")
				continue
			}
			color := handle(colCell)
			if color != "" {
				setColor(f, cellName, color)
			}
		}
		setCount(f, i+1)
	}
}

func handle(cell string) (color string) {
	cell = strings.TrimSpace(cell)
	re := regexp.MustCompile(`\s+`)
	output := re.ReplaceAllString(cell, "\n")
	explode := strings.Split(output, "\n")
	count := len(explode)
	if count > 2 {
		color = "FFA500"
	}
	for i, v := range explode {
		if checkChinese(v) {
			color = "FFA500"
			continue
		}
		v = strings.ReplaceAll(v, ":", "")
		v, _ := strconv.Atoi(v)
		if v > 900 && i == 0 {
			color = "FF0000"
			delay++
		} else if v <= 850 {
			before++
		} else if v >= 1710 {
			after++
		}
	}
	return
}

func checkChinese(s string) bool {
	for _, r := range s {
		if r >= 0x4E00 && r <= 0x9FFF {
			return true
		}
	}
	return false
}

func isFullChinese(s string) bool {
	v := strings.ReplaceAll(s, ":", "")
	v = strings.ReplaceAll(v, "\n", "")
	v = strings.ReplaceAll(v, " ", "")
	for _, r := range v {
		if r < 0x4E00 || r > 0x9FFF {
			return false
		}
	}
	return true
}

func setColor(f *excelize.File, cellName string, color string) {
	style, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{color}, // 黄色
		},
	})
	if err != nil {
		fmt.Println(err)
	}
	err = f.SetCellStyle(sheetName, cellName, cellName, style)
	if err != nil {
		fmt.Println(err)
	}
}

func setTitle(f *excelize.File) {
	setCountValue(f, fmt.Sprintf("%s%d", "AB", 1), "早到次数")
	setCountValue(f, fmt.Sprintf("%s%d", "AC", 1), "晚走次数")
	setCountValue(f, fmt.Sprintf("%s%d", "AD", 1), "迟到次数")
}

func setCount(f *excelize.File, row int) {
	setCountValue(f, fmt.Sprintf("%s%d", "AB", row), before)
	setCountValue(f, fmt.Sprintf("%s%d", "AC", row), after)
	setCountValue(f, fmt.Sprintf("%s%d", "AD", row), delay)
}

func setCountValue(f *excelize.File, column string, value interface{}) {
	err := f.SetCellValue(sheetName, column, value)
	if err != nil {
		fmt.Println(err)
	}
}
