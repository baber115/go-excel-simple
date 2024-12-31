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

var vacation = []string{
	"早到次数",
	"晚走次数",
	"迟到次数",
	"年假",
	"婚假",
	"病假",
	"调休",
	"事假",
	"产假",
	"丧假",
}

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
	rowOptions := excelize.Options{
		RawCellValue: true,
	}
	rows, _ := f.GetRows(sheetName, rowOptions)
	cols, _ := f.GetCols(sheetName)
	rowLen := len(rows)
	colLen := len(cols)
	setTitle(f, colLen)
	for i := 1; i < rowLen; i++ {
		delay, after, before = 0, 0, 0
		vacationDays := map[string][]string{}

		for j := 2; j < colLen; j++ {
			cellName, _ := excelize.CoordinatesToCellName(j+1, i+1)
			colCell, _ := f.GetCellValue(sheetName, cellName)
			str := isVacation(colCell)
			if str != "" {
				vacationDays[str] = append(vacationDays[str], rows[0][j])
			}
			if colCell == "" || isFullChinese(colCell) {
				setColor(f, cellName, "FFFF00")
				continue
			}
			color := handle(colCell)
			if color != "" {
				setColor(f, cellName, color)
			}
		}
		setValidation(f, colLen, i, vacationDays)
		setCount(f, colLen, i)
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

func isVacation(str string) string {
	for _, v := range vacation {
		if strings.Count(str, v) >= 1 {
			return v
		}
	}

	return ""
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

func setTitle(f *excelize.File, colLen int) {
	for k, v := range vacation {
		celName, _ := excelize.CoordinatesToCellName(colLen+k+1, 1)
		setCountValue(f, celName, v)
	}
}

func setCount(f *excelize.File, colLen int, row int) {
	for k, v := range []int{before, after, delay} {
		celName, _ := excelize.CoordinatesToCellName(colLen+k+1, row+1)
		setCountValue(f, celName, v)
	}
}

func setValidation(f *excelize.File, colLen int, row int, validationDays map[string][]string) {
	for k, v := range vacation {
		if days, exists := validationDays[v]; exists {
			day := strings.Join(days, "、")
			celName, _ := excelize.CoordinatesToCellName(colLen+k+1, row+1)
			setCountValue(f, celName, day)
		}
	}
}

func setCountValue(f *excelize.File, column string, value interface{}) {
	err := f.SetCellValue(sheetName, column, value)
	if err != nil {
		fmt.Println(err)
	}
}
