package main

import (
	"bufio"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"strconv"
	"strings"
)

type specificCity struct {
	Name      string
	InfectNum int
}

type specificProvince struct {
	Name     string
	TotalNum int
	Cities   []specificCity
}

type statisticValue struct {
	Name      string
	TotalNum  int
	provinces []specificProvince
}

func main() {
	//确诊数据处理
	var inputConfirmedValue string
	fmt.Println("请输入确诊数据")
	input1 := bufio.NewScanner(os.Stdin)
	input1.Scan()
	inputConfirmedValue = input1.Text()
	// 确诊病例
	confirmedCase(inputConfirmedValue)
	// 无症状病例
	var inputAsymptomaticValue string
	fmt.Println("请输入无症状数据")
	input2 := bufio.NewScanner(os.Stdin)
	input2.Scan()
	inputAsymptomaticValue = input2.Text()
	asymptomaticCase(inputAsymptomaticValue)
}

func asymptomaticCase(rawValue string) {
	firstComma := strings.Index(rawValue, "，")
	secondComma := (firstComma + 3) + strings.Index(rawValue[firstComma+3:], "，")
	localConfirmedCase(rawValue[secondComma+3:], "本土无症状", "localNoSymptom.xlsx")
}

func confirmedCase(rawValue string) {
	sentenceSplit := strings.Split(rawValue, "。")
	splitValue := sentenceSplit[1]
	// 第二句按;拆分
	firstSemicolon := strings.Index(splitValue, "；")
	// 境外输入
	overseasConfirmedCase(splitValue[0:firstSemicolon])
	// 本土确诊
	localConfirmedCase(splitValue[firstSemicolon+3:], "本土确诊", "localConfirmed.xlsx")
}

func localConfirmedCase(rawValue string, sheetName string, exportName string) {
	// 记录数据的起始字段
	start := 0
	var localConfirmed statisticValue
	localConfirmed.Name = sheetName
	localConfirmed.TotalNum, start = getNum(rawValue, start, len(rawValue))
	// 所有省份的数据
	firstBracketIndex := strings.Index(rawValue, "（")
	secondBracketIndex := strings.Index(rawValue, "）")
	allProvinceValue := rawValue[firstBracketIndex+3 : secondBracketIndex]
	// 各个省份的数据
	eachProvinceValue := strings.Split(allProvinceValue, "；")
	for i := 0; i < len(eachProvinceValue); i++ {
		start = 0
		length := len(eachProvinceValue[i])
		var province specificProvince
		province.Name, start = getName(eachProvinceValue[i], start, length)
		province.TotalNum, start = getNum(eachProvinceValue[i], start, length)
		start += 6
		for start < length {
			var city specificCity
			city.Name, start = getName(eachProvinceValue[i], start, length)
			city.InfectNum, start = getNum(eachProvinceValue[i], start, length)
			start += 6
			province.Cities = append(province.Cities, city)
		}
		localConfirmed.provinces = append(localConfirmed.provinces, province)
	}
	//fmt.Println(localConfirmed)
	export(localConfirmed, exportName)
}

func overseasConfirmedCase(rawValue string) {
	// 记录数据的起始字段
	start := 0
	var overseasInput statisticValue
	overseasInput.Name = "境外输入确诊"
	overseasInput.TotalNum, start = getNum(rawValue, start, len(rawValue))
	// 各个省份的处理
	firstBracketIndex := strings.Index(rawValue, "（")
	secondBracketIndex := strings.Index(rawValue, "）")
	provinceValue := rawValue[firstBracketIndex+3 : secondBracketIndex]
	start = 0
	length := len(provinceValue)
	for start < length {
		var province specificProvince
		province.Name, start = getName(provinceValue, start, length)
		province.TotalNum, start = getNum(provinceValue, start, length)
		start += 6
		overseasInput.provinces = append(overseasInput.provinces, province)
	}
	//fmt.Println(overseasInput)
	export(overseasInput, "overseasInput.xlsx")
}

func getName(str string, index int, length int) (string, int) {
	var start, end int
	start = index
	end = start
	for {
		if end >= length {
			return str[start:end], length + 1
		} else if str[end] >= '0' && str[end] <= '9' {
			break
		} else {
			end++
		}
	}
	return str[start:end], end
}

func getNum(str string, index int, length int) (int, int) {
	for {
		// 边界判断
		if index >= length {
			return 0, length + 1
		}
		if str[index] >= '0' && str[index] <= '9' {
			break
		}
		index++
	}
	var num int
	for str[index] >= '0' && str[index] <= '9' {
		num = num*10 + int(str[index]-'0')
		if index < length {
			index++
		} else {
			return 0, length + 1
		}
	}
	return num, index
}

func export(exportValue statisticValue, fileName string) {
	f := excelize.NewFile()
	style, err := f.NewStyle(`{
    "border": [
		{
			"type": "left",
			"color": "#000000",
			"style": 1
		},
		{
			"type": "top",
			"color": "#000000",
			"style": 1
		},
		{
			"type": "bottom",
			"color": "#000000",
			"style": 1
		},
		{
			"type": "right",
			"color": "#000000",
			"style": 1
		}],
	"font":{"bold":true,"size":12},
    "alignment": {
			"horizontal": "center",
			"vertical": "center"
		}
	}`)
	if err != nil {
		fmt.Println(err)
	}
	sheet := f.NewSheet("Sheet1")
	// 设置表头
	var exportName = []string{"类型", "总人数", "省/直辖市", "人数", "城市/地区", "人数"}
	var SheetName = []string{"A1", "B1", "C1", "D1", "E1", "F1"}
	for sheetIndex, value := range exportName {
		if err := f.SetCellStr("Sheet1", SheetName[sheetIndex], value); err != nil {
			fmt.Printf("[Export] ser cell str error {err=%#v}", err)
		}
	}
	// 填写数据
	if err := f.SetCellValue("Sheet1", "A2", exportValue.Name); err != nil {
		fmt.Printf("[Export] set cell vaule error {err=%#v, name=%#v, value=%#v}", err, "A2", exportValue.Name+strconv.Itoa(exportValue.TotalNum))
	}
	if err := f.SetCellValue("Sheet1", "B2", strconv.Itoa(exportValue.TotalNum)); err != nil {
		fmt.Printf("[Export] set cell vaule error {err=%#v, name=%#v, value=%#v}", err, "B2", exportValue.Name+strconv.Itoa(exportValue.TotalNum))
	}
	i := 2
	for _, province := range exportValue.provinces {
		index := 2
		start := i
		if err := f.SetCellValue("Sheet1", SheetName[index][:1]+strconv.Itoa(i), province.Name); err != nil {
			fmt.Printf("[Export] set cell vaule error {err=%#v, name=%#v, value=%#v}", err, SheetName[index][:1]+strconv.Itoa(i), exportValue.Name+strconv.Itoa(exportValue.TotalNum))
		}
		index += 1
		if err := f.SetCellValue("Sheet1", SheetName[index][:1]+strconv.Itoa(i), strconv.Itoa(province.TotalNum)); err != nil {
			fmt.Printf("[Export] set cell vaule error {err=%#v, name=%#v, value=%#v}", err, SheetName[index][:1]+strconv.Itoa(i), exportValue.Name+strconv.Itoa(exportValue.TotalNum))
		}
		index += 1
		if len(province.Cities) != 0 {
			for _, city := range province.Cities {
				if err := f.SetCellValue("Sheet1", SheetName[index][:1]+strconv.Itoa(i), city.Name); err != nil {
					fmt.Printf("[Export] set cell vaule error {err=%#v, name=%#v, value=%#v}", err, SheetName[index][:1]+strconv.Itoa(i), exportValue.Name+strconv.Itoa(exportValue.TotalNum))
				}
				index += 1
				if err := f.SetCellValue("Sheet1", SheetName[index][:1]+strconv.Itoa(i), strconv.Itoa(city.InfectNum)); err != nil {
					fmt.Printf("[Export] set cell vaule error {err=%#v, name=%#v, value=%#v}", err, SheetName[index][:1]+strconv.Itoa(i), exportValue.Name+strconv.Itoa(exportValue.TotalNum))
				}
				index -= 1
				i++
			}
		} else {
			i++
		}
		f.MergeCell("Sheet1", "C"+strconv.Itoa(start), "C"+strconv.Itoa(i-1))
		f.MergeCell("Sheet1", "D"+strconv.Itoa(start), "D"+strconv.Itoa(i-1))
	}
	f.MergeCell("Sheet1", "A2", "A"+strconv.Itoa(i-1))
	f.MergeCell("Sheet1", "B2", "B"+strconv.Itoa(i-1))
	if err := f.SetCellStyle("Sheet1", "A1", "F"+strconv.Itoa(i-1), style); err != nil {
		fmt.Printf("[Export] ser cell style error {err=%#v}", err)
	}

	f.SetActiveSheet(sheet)
	f.SaveAs(fileName)
}
