package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"regexp"
	"strconv"
	"strings"
)

var chineseRules = []string{
	"外勤", "外出", "病假", "年假", "调休", "补卡", "事假", "婚假", "产假", "丧假", "哺乳假", "产检", "体检", "年休",
}
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
	for i, row := range rows {
		if i == 0 {
			continue
		}
		delay, after, before = 0, 0, 0
		for j, colCell := range row {
			cellName, _ := excelize.CoordinatesToCellName(j+1, i+1)
			if colCell == "" || checkChinese(&colCell) {
				setColor(f, cellName)
			} else {
				if handle(colCell) == false {
					setColor(f, cellName)
				}
			}
		}
		setCount(f, i+1)
	}
}

func handle(cell string) bool {
	cell = strings.TrimSpace(cell)
	re := regexp.MustCompile(`\s+`)
	output := re.ReplaceAllString(cell, "\n")

	explode := strings.Split(output, "\n")
	count := len(explode)
	if count == 1 {
		return false
	}
	for i, v := range explode {
		v = strings.ReplaceAll(v, ":", "")
		v, _ := strconv.Atoi(v)
		if v > 900 && i == 0 {
			delay++
		} else if v <= 850 {
			before++
		} else if v >= 1710 {
			after++
		}
	}
	return true
}

func checkChinese(s *string) bool {
	for _, v := range chineseRules {
		if strings.Contains(*s, v) {
			*s = strings.ReplaceAll(*s, v, "")
			fmt.Println(*s)
			return false
		}
	}

	//return strings.Contains(s, "外勤")
	for _, r := range *s {
		if r >= 0x4E00 && r <= 0x9FFF {
			return true
		}
	}
	return false
}

func setColor(f *excelize.File, cellName string) {
	style, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{"FFFF00"}, // 黄色
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

func setCount(f *excelize.File, row int) {
	err := f.SetCellValue(sheetName, fmt.Sprintf("%s%d", "Y", row), before)
	if err != nil {
		fmt.Println(err)
	}
	err = f.SetCellValue(sheetName, fmt.Sprintf("%s%d", "Z", row), after)
	if err != nil {
		fmt.Println(err)
	}
	err = f.SetCellValue(sheetName, fmt.Sprintf("%s%d", "AA", row), delay)
	if err != nil {
		fmt.Println(err)
	}
}
