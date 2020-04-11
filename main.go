package main


import (
"fmt"

"github.com/tealeg/xlsx"
)

func main() {
	excelFileName := "机会升舱数据.xlsx"

	nameValus, nameTimes := readExcelFile(excelFileName)
	saveExcelFile(excelFileName,nameValus,nameTimes)
}
func readExcelFile1(excelFileName string) (nameValue map[string]int64, nameTimes map[string]int64){
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		fmt.Printf("open failed: %s\n", err)
	}

	count := 0
	nameList := make(map[string]int64)
	nameListTimes := make(map[string]int64)
	for _, sheet := range xlFile.Sheets {
		fmt.Printf("Sheet Name: %s\n", sheet.Name)

		if sheet.Name == "Sheet1" {
			for _, row := range sheet.Rows {
				fmt.Println("--------------------------------")
				var tempNmae string
				var firstName string
				var secondName string
				var tempValue int64
				for i :=0; i < 2;i++ {
					cell :=row.Cells[i]

					if i == 0 {
						firstName = cell.String()
					}
					if i == 1 {
						secondName = cell.String()
					}
				}
				if len(nameList) == 0 {
					tempNmae = firstName+ secondName
					nameList[tempNmae] = tempValue
					nameListTimes[tempNmae] = 1
					fmt.Println("tempNmae:",tempNmae, "value:",nameListTimes[tempNmae])

				} else {
					tempNmae = firstName + secondName
					val, ok := nameList[tempNmae] //ok为true时，代表有key
					if ok {
						val = val + tempValue
						nameList[tempNmae] = val
						nameListTimes[tempNmae] = nameListTimes[tempNmae] + 1
						fmt.Println("tempNmae:",tempNmae, "value:",nameListTimes[tempNmae])
					} else {
						tempNmae = secondName + firstName
						val, ok := nameList[tempNmae] //ok为true时，代表有key
						if ok {
							val = val + tempValue
							nameList[tempNmae] = val
							nameListTimes[tempNmae] = nameListTimes[tempNmae] + 1
							fmt.Println("tempNmae:",tempNmae, "value:",nameListTimes[tempNmae])

						} else {
							nameList[tempNmae] = tempValue
							nameListTimes[tempNmae] = 1
							fmt.Println("tempNmae:",tempNmae, "value:",nameListTimes[tempNmae])

						}
					}
				}
				count++
			}
		}
		fmt.Println("count：",count)
		fmt.Println("nameList:",nameList)
		fmt.Println("nameListTimes:",nameListTimes)
	}
	return nameList, nameListTimes

}
func readExcelFile(excelFileName string) (nameValue map[string]int64, nameTimes map[string]int64){
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		fmt.Printf("open failed: %s\n", err)
	}

	count := 0
	nameList := make(map[string]int64)
	nameListTimes := make(map[string]int64)
	for _, sheet := range xlFile.Sheets {
		fmt.Printf("Sheet Name: %s\n", sheet.Name)

		if sheet.Name == "Sheet1" {
			for _, row := range sheet.Rows {
				fmt.Println("--------------------------------")
				var tempNmae string
				var firstName string
				var secondName string
				var tempValue int64
				for i :=0; i < 5;i++ {
					cell :=row.Cells[i]
					if i == 4 {
						tempValue,_ = cell.Int64()
					}
					if i == 0 {
						firstName = cell.String()
					}
					if i == 1 {
						secondName = cell.String()
					}
				}
				if len(nameList) == 0 {
					tempNmae = firstName+ secondName
					nameList[tempNmae] = tempValue
					nameListTimes[tempNmae] = 1
					fmt.Println("tempNmae:",tempNmae, "value:",nameListTimes[tempNmae])

				} else {
					tempNmae = firstName + secondName
					val, ok := nameList[tempNmae] //ok为true时，代表有key
					if ok {
						val = val + tempValue
						nameList[tempNmae] = val
						nameListTimes[tempNmae] = nameListTimes[tempNmae] + 1
						fmt.Println("tempNmae:",tempNmae, "value:",nameListTimes[tempNmae])
					} else {
						tempNmae = secondName + firstName
						val, ok := nameList[tempNmae] //ok为true时，代表有key
						if ok {
							val = val + tempValue
							nameList[tempNmae] = val
							nameListTimes[tempNmae] = nameListTimes[tempNmae] + 1
							fmt.Println("tempNmae:",tempNmae, "value:",nameListTimes[tempNmae])

						} else {
							nameList[tempNmae] = tempValue
							nameListTimes[tempNmae] = 1
							fmt.Println("tempNmae:",tempNmae, "value:",nameListTimes[tempNmae])

						}
					}
				}
				count++
			}
		}
		fmt.Println("count：",count)
		fmt.Println("nameList:",nameList)
		fmt.Println("nameListTimes:",nameListTimes)
	}
	return nameList, nameListTimes

}
//把数据存入Excel
func saveExcelFile(excelFileName string, nameValue map[string]int64, nameTimes map[string]int64) {
	xlFile, err := xlsx.OpenFile(excelFileName)

	sheet, err := xlFile.AddSheet("完成数据2")
	if err != nil {
		fmt.Printf(err.Error())
	}
	for name,value := range nameValue {
		times := nameTimes[name]
		row := sheet.AddRow()
		cell := row.AddCell()
		cell.Value = name
		cell = row.AddCell()
		cell.SetInt64(value)
		cell = row.AddCell()
		cell.SetInt64(times)
	}
	err = xlFile.Save(excelFileName)
	if err != nil {
		fmt.Printf(err.Error())
	}
}
