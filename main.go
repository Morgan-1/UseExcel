package main


import (
"fmt"
	"strconv"

	"github.com/tealeg/xlsx"
)

func main() {
	excelFileName := "统计航段信息.xlsx"
	lineTimes, lineAccounts, linePeoples, lineCountAge := readExceFileData(excelFileName)
	saveExcelFileData(excelFileName,lineTimes, lineAccounts, linePeoples, lineCountAge)
	//nameValus, nameTimes, nameAge20List,nameAge30List,nameAge50List,nameAge100List,nameAveAge20List1 ,
	//nameAveAge30List1, nameAveAge50List1 ,nameAveAge100List1  := readExcelFile2(excelFileName)
	//saveExcelFile1(excelFileName,nameValus,nameTimes, nameAge20List,nameAge30List,nameAge50List,nameAge100List,nameAveAge20List1 ,
		//nameAveAge30List1, nameAveAge50List1 ,nameAveAge100List1)
	//nameValus, nameTimes := readExcelFile1(excelFileName)
	//saveExcelFile(excelFileName,nameValus,nameTimes)
}

func readExcelFile2(excelFileName string) (nameValue map[string]int64, nameTimes map[string]int64,
nameAge20List1 map[string]int64,nameAge30List1 map[string]int64,
nameAge50List1 map[string]int64,nameAge100List1 map[string]int64,
nameAveAge20List1 map[string]int64,nameAveAge30List1 map[string]int64,
nameAveAge50List1 map[string]int64,nameAveAge100List1 map[string]int64){
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		fmt.Printf("open failed: %s\n", err)
	}

	count := 0
	nameList := make(map[string]int64)
	nameListTimes := make(map[string]int64)
	nameAge20List := make(map[string]int64)   //<20
	nameAge30List := make(map[string]int64)  //>=20 <30
	nameAge50List := make(map[string]int64)  //>=30 <50
	nameAge100List := make(map[string]int64) //>=50
	nameAveAge20List := make(map[string]int64)   //<20
	nameAveAge30List := make(map[string]int64)  //>=20 <30
	nameAveAge50List := make(map[string]int64)  //>=30 <50
	nameAveAge100List := make(map[string]int64) //>=50
	for _, sheet := range xlFile.Sheets {
		fmt.Printf("Sheet Name: %s\n", sheet.Name)

		if sheet.Name == "复购主要航段" {
			for _, row := range sheet.Rows {
				fmt.Println("--------------------------------")
				var tempNmae string
				var firstName string
				var secondName string
				var tempValue int64
				var tempage int64
				for i :=0; i < 5;i++ {
					cell :=row.Cells[i]
					if i == 4 {
						tempValue=1
					}

					if i == 1 {
						tempage,_ = cell.Int64()
					}
					if i == 2 {
						firstName = cell.String()
					}
					if i == 3 {
						secondName = cell.String()
					}
				}
				if len(nameList) == 0 {
					tempNmae = firstName+ secondName
					nameList[tempNmae] = tempValue
					nameListTimes[tempNmae] = 1
					if tempage < 20{
						nameAge20List[tempNmae] =nameAge20List[tempNmae] + 1
						nameAveAge20List[tempNmae] =nameAveAge20List[tempNmae] + tempage
					}else if tempage >= 20 && tempage <30 {
						nameAge30List[tempNmae] = nameAge30List[tempNmae]+1
						nameAveAge30List[tempNmae] =nameAveAge30List[tempNmae] + tempage
					}else if tempage >= 30 && tempage <50 {
						nameAge50List[tempNmae] = nameAge50List[tempNmae]+1
						nameAveAge50List[tempNmae] =nameAveAge50List[tempNmae] + tempage
					}else  {
						nameAge100List[tempNmae] = nameAge100List[tempNmae]+1
						nameAveAge100List[tempNmae] =nameAveAge100List[tempNmae] + tempage
					}
					fmt.Println("tempNmae:",tempNmae, "value:",nameListTimes[tempNmae] ," nameAgeList:",nameAge20List,nameAge30List,nameAge50List,nameAge100List)

				} else {
					tempNmae = firstName + secondName
					val, ok := nameList[tempNmae] //ok为true时，代表有key
					if ok {
						val = val + tempValue
						nameList[tempNmae] = val
						nameListTimes[tempNmae] = nameListTimes[tempNmae] + 1
						if tempage < 20{
							nameAge20List[tempNmae] =nameAge20List[tempNmae] + 1
							nameAveAge20List[tempNmae] =nameAveAge20List[tempNmae] + tempage
						}else if tempage >= 20 && tempage <30 {
							nameAge30List[tempNmae] = nameAge30List[tempNmae]+1
							nameAveAge30List[tempNmae] =nameAveAge30List[tempNmae] + tempage
						}else if tempage >= 30 && tempage <50 {
							nameAge50List[tempNmae] = nameAge50List[tempNmae]+1
							nameAveAge50List[tempNmae] =nameAveAge50List[tempNmae] + tempage
						}else  {
							nameAge100List[tempNmae] = nameAge100List[tempNmae]+1
							nameAveAge100List[tempNmae] =nameAveAge100List[tempNmae] + tempage
						}
						fmt.Println("tempNmae:",tempNmae, "value:",nameListTimes[tempNmae])
					} else {
						tempNmae1 := secondName + firstName
						val, ok := nameList[tempNmae1] //ok为true时，代表有key
						if ok {
							val = val + tempValue
							nameList[tempNmae1] = val
							nameListTimes[tempNmae1] = nameListTimes[tempNmae1] + 1
							if tempage < 20{
								nameAge20List[tempNmae1] =nameAge20List[tempNmae1] + 1
								nameAveAge20List[tempNmae1] =nameAveAge20List[tempNmae1] + tempage
							}else if tempage >= 20 && tempage <30 {
								nameAge30List[tempNmae1] = nameAge30List[tempNmae1]+1
								nameAveAge30List[tempNmae1] =nameAveAge30List[tempNmae1] + tempage
							}else if tempage >= 30 && tempage <50 {
								nameAge50List[tempNmae1] = nameAge50List[tempNmae1]+1
								nameAveAge50List[tempNmae1] =nameAveAge50List[tempNmae1] + tempage
							}else  {
								nameAge100List[tempNmae1] = nameAge100List[tempNmae1]+1
								nameAveAge100List[tempNmae1] =nameAveAge100List[tempNmae1] + tempage
							}
							fmt.Println("tempNmae:",tempNmae1, "value:",nameListTimes[tempNmae])

						} else {
							nameList[tempNmae] = tempValue
							nameListTimes[tempNmae] =1
							if tempage < 20{
								nameAge20List[tempNmae] =nameAge20List[tempNmae] + 1
								nameAveAge20List[tempNmae] =nameAveAge20List[tempNmae] + tempage
							}else if tempage >= 20 && tempage <30 {
								nameAge30List[tempNmae] = nameAge30List[tempNmae]+1
								nameAveAge30List[tempNmae] =nameAveAge30List[tempNmae] + tempage
							}else if tempage >= 30 && tempage <50 {
								nameAge50List[tempNmae] = nameAge50List[tempNmae]+1
								nameAveAge50List[tempNmae] =nameAveAge50List[tempNmae] + tempage
							}else  {
								nameAge100List[tempNmae] = nameAge100List[tempNmae]+1
								nameAveAge100List[tempNmae] =nameAveAge100List[tempNmae] + tempage
							}
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
		fmt.Println(" nameAgeList:",nameAge20List,nameAge30List,nameAge50List,nameAge100List)

	}
	return nameList, nameListTimes, nameAge20List,nameAge30List,nameAge50List,nameAge100List, nameAveAge20List,nameAveAge30List,nameAveAge50List,nameAveAge100List

}
//求航线   交易次数     总的收入    旅客人数    旅客平均年龄
func readExceFileData(excelFileName string)(lineTimes map[string]int64, lineAccounts map[string]int64,
	linePeoples map[string]map[string]int64, lineCountAge map[string]int64){
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		fmt.Printf("open failed: %s\n", err)
	}

	count := 0
	lineTimes = make(map[string]int64)
	lineAccounts = make(map[string]int64)
	linePeoples = make(map[string]map[string]int64)
	lineCountAge = make(map[string]int64)
	for _, sheet := range xlFile.Sheets {
		fmt.Printf("Sheet Name: %s\n", sheet.Name)

		if sheet.Name == "Sheet1" {
			for i := 1; i < len(sheet.Rows);i++ {
				fmt.Println("--------------------------------")
				//从1 开始是为了去掉第一行标题
				row := sheet.Rows[i]
				//出发地
				var firstName string
				//目的地
				var secondName string
				//航线名
				var lineName string
				//航线次数
				var tempTimes int64
				//名字
				var tempName string
				//年龄
				var tempAge int64
				//价格
				var tempAccount int64
				for i :=0; i < 6;i++ {
					cell :=row.Cells[i]
					if i == 0 {
						firstName = cell.String()
					}
					if i == 1 {
						secondName = cell.String()
					}
					if i == 2 {
						tempName = cell.String()
					}
					if i == 3 {
						tempAge, err = cell.Int64()
						if err != nil  {
							tempAge = 0
						}
					}
					if i == 4 {
						tempAccount,err = cell.Int64()
						if err != nil  {
							tempAccount = 0
						}
					}
					if i == 5 {
						tempTimes,err = cell.Int64()
						if err != nil  {
							tempTimes = 0
						}
					}

				}
				lineName = firstName+ secondName
				//第一次的时候添加数据，lineTimes没有数据
				if len(lineTimes) == 0 {

					lineTimes,lineAccounts,linePeoples,lineCountAge = setLists(lineName,tempTimes,tempName,tempAge,tempAccount,
						lineTimes,lineAccounts,linePeoples,lineCountAge)
				} else {
					_, ok := lineTimes[lineName] //ok为true时，代表有key
					if ok {
						lineTimes,lineAccounts,linePeoples,lineCountAge = setLists(lineName,tempTimes,tempName,tempAge,tempAccount,
							lineTimes,lineAccounts,linePeoples,lineCountAge)
					} else {
						//反向航线
						lineName2 := secondName + firstName
						_, ok := lineTimes[lineName2] //ok为true时，代表有key
						if ok {
							lineTimes,lineAccounts,linePeoples,lineCountAge = setLists(lineName2,tempTimes,tempName,tempAge,tempAccount,
								lineTimes,lineAccounts,linePeoples,lineCountAge)
						} else {
							lineTimes,lineAccounts,linePeoples,lineCountAge = setLists(lineName,tempTimes,tempName,tempAge,tempAccount,
								lineTimes,lineAccounts,linePeoples,lineCountAge)
						}
					}
				}
				count++
			}
		}

	}
	return lineTimes, lineAccounts, linePeoples, lineCountAge

}
func setLists(lineName string, tempTimes int64, tempName string, tempAge int64,  tempAccount int64,lineTimes map[string]int64, lineAccounts map[string]int64, linePeoples map[string]map[string]int64, lineCountAge map[string]int64) (lineTimes1 map[string]int64, lineAccounts1 map[string]int64, linePeoples1 map[string]map[string]int64, lineCountAge1 map[string]int64) {
	lineTimes[lineName] = lineTimes[lineName] + tempTimes
	lineAccounts[lineName] = lineAccounts[lineName] + tempAccount
	peoples,isTrue :=  linePeoples[lineName]
	if !isTrue {
		peoples = make(map[string]int64)
	}
	//判断该航线是否存在这个人
	_, ok := peoples[tempName]
	if !ok  {
		peoples[tempName] = tempAge
	}
	linePeoples[lineName] = peoples
	lineCountAge[lineName] = lineCountAge[lineName] +tempAge
	fmt.Println("lineName:",lineName, " lineAccounts:",lineAccounts, " linePeoples:",linePeoples," lineCountAge:",lineCountAge)
	return lineTimes, lineAccounts, linePeoples, lineCountAge
}
//把数据存入Excel
func saveExcelFileData(excelFileName string, lineTimes map[string]int64, lineAccounts map[string]int64, linePeoples map[string]map[string]int64, lineCountAge map[string]int64) {
	fmt.Println("open file save excel")
	xlFile, err := xlsx.OpenFile(excelFileName)

	sheet, err := xlFile.AddSheet("统计结果")
	if err != nil {
		fmt.Printf(err.Error())
	}
	row := sheet.AddRow()
	cell := row.AddCell()
	cell.Value = "航段名字"
	cell = row.AddCell()
	cell.Value = "总共乘坐次数"
	cell = row.AddCell()
	cell.Value = "总共航段的收入"
	cell = row.AddCell()
	cell.Value = "旅客人数"
	cell = row.AddCell()
	cell.Value = "旅客信息"
	cell = row.AddCell()
	cell.Value = "旅客平均年龄"

	for name,value := range lineTimes {
		times := value
		accounts := lineAccounts[name]
		peoples := linePeoples[name]
		peoplesNum := len(peoples)
		var peoplesStr = " "
		for key,value := range peoples {
			peoplesStr = peoplesStr+"name:"+ key+ "/  age :"+strconv.FormatInt(value,10) + " - "
		}
		countAge := lineCountAge[name]
		aveAge := countAge / times
		row := sheet.AddRow()
		cell := row.AddCell()
		cell.Value = name

		cell = row.AddCell()
		cell.SetInt64(times)
		cell = row.AddCell()
		cell.SetInt64(accounts)
		cell = row.AddCell()
		cell.SetInt64(int64(peoplesNum))
		cell = row.AddCell()
		cell.Value = peoplesStr
		cell = row.AddCell()
		cell.SetInt64(aveAge)

	}
	err = xlFile.Save(excelFileName)
	if err != nil {
		fmt.Printf(err.Error())
	}
	fmt.Println("close file save excel")

}

//判断有来回的航线占比
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

		if sheet.Name == "机上" {
			for _, row := range sheet.Rows {
				fmt.Println("--------------------------------")
				var tempNmae string
				var firstName string
				var secondName string
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
					nameList[tempNmae] = 0
					nameListTimes[tempNmae] =  nameListTimes[tempNmae] + 1
					fmt.Println("tempNmae:",tempNmae, "value:",nameListTimes[tempNmae])

				} else {
					tempNmae = firstName + secondName
					val, ok := nameList[tempNmae] //ok为true时，代表有key
					if ok {
						nameList[tempNmae] = val
						nameListTimes[tempNmae] = nameListTimes[tempNmae] + 1
						fmt.Println("tempNmae:",tempNmae, "value:",nameListTimes[tempNmae])
					} else {
						tempNmae2 := secondName + firstName
						val, ok := nameList[tempNmae2] //ok为true时，代表有key
						if ok {
							val = val + 1
							nameList[tempNmae2] = val
							nameListTimes[tempNmae2] = nameListTimes[tempNmae2] + 1
							fmt.Println("tempNmae:",tempNmae, "value:",nameListTimes[tempNmae])

						} else {
							nameList[tempNmae] = 0
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
				var tempValue1 int64
				for i :=0; i < 5;i++ {
					cell :=row.Cells[i]
					if i == 3 {
						tempValue,_ = cell.Int64()
					}
					if i == 4 {
						tempValue1,_ = cell.Int64()
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
					nameListTimes[tempNmae] = tempValue1
					fmt.Println("tempNmae:",tempNmae, "value:",nameListTimes[tempNmae])

				} else {
					tempNmae = firstName + secondName
					val, ok := nameList[tempNmae] //ok为true时，代表有key
					if ok {
						val = val + tempValue
						nameList[tempNmae] = val
						nameListTimes[tempNmae] = nameListTimes[tempNmae] + tempValue1
						fmt.Println("tempNmae:",tempNmae, "value:",nameListTimes[tempNmae])
					} else {
						tempNmae = secondName + firstName
						val, ok := nameList[tempNmae] //ok为true时，代表有key
						if ok {
							val = val + tempValue
							nameList[tempNmae] = val
							nameListTimes[tempNmae] = nameListTimes[tempNmae] + tempValue1
							fmt.Println("tempNmae:",tempNmae, "value:",nameListTimes[tempNmae])

						} else {
							nameList[tempNmae] = tempValue
							nameListTimes[tempNmae] =tempValue1
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

	sheet, err := xlFile.AddSheet("完成数据3")
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

//把数据存入Excel
func saveExcelFile1(excelFileName string, nameValue map[string]int64,
	nameTimes map[string]int64,nameAge20List map[string]int64,
	nameAge30List map[string]int64,nameAge50List map[string]int64,nameAge100List map[string]int64,
	nameAveAge20List map[string]int64,nameAveAge30List map[string]int64,
	nameAveAge50List map[string]int64,nameAveAge100List map[string]int64) {
	xlFile, err := xlsx.OpenFile(excelFileName)

	sheet, err := xlFile.AddSheet("完成数据")
	if err != nil {
		fmt.Printf(err.Error())
	}
	row := sheet.AddRow()
	cell := row.AddCell()
	cell.Value = "航段名字"
	cell = row.AddCell()
	cell.Value = "总的收入"
	cell = row.AddCell()
	cell.Value = "次数"
	cell = row.AddCell()
	cell.Value = "20岁以下人数"
	cell = row.AddCell()
	cell.Value = "20-30岁人数"
	cell = row.AddCell()
	cell.Value = "30-50岁人数"
	cell = row.AddCell()
	cell.Value = "50岁以上人数"
	cell = row.AddCell()
	cell.Value = "20岁以下人数平均年龄"
	cell = row.AddCell()
	cell.Value = "20-30岁人数平均年龄"
	cell = row.AddCell()
	cell.Value = "30-50岁人数平均年龄"
	cell = row.AddCell()
	cell.Value = "50岁以上人数平均年龄"
	for name,value := range nameValue {
		times := nameTimes[name]
		age20 := nameAge20List[name]
		age30 := nameAge30List[name]
		age50 := nameAge50List[name]
		age100 := nameAge100List[name]
		var aveAge20, aveAge30,aveAge50, aveAge100 int64
		if age20 == 0  {
			aveAge20 = 0
		}else {
			aveAge20 = nameAveAge20List[name]/age20
		}
		if age30 == 0  {
			aveAge30 = 0
		}else {
			aveAge30 = nameAveAge30List[name]/age30
		}
		if age50 == 0  {
			aveAge50 = 0
		}else {
			aveAge50 = nameAveAge50List[name]/age50
		}
		if age100 == 0  {
			aveAge100 = 0
		}else {
			aveAge100 = nameAveAge100List[name]/age100
		}

		row := sheet.AddRow()
		cell := row.AddCell()
		cell.Value = name
		cell = row.AddCell()
		cell.SetInt64(value)
		cell = row.AddCell()
		cell.SetInt64(times)
		cell = row.AddCell()
		cell.SetInt64(age20)
		cell = row.AddCell()
		cell.SetInt64(age30)
		cell = row.AddCell()
		cell.SetInt64(age50)
		cell = row.AddCell()
		cell.SetInt64(age100)
		cell = row.AddCell()
		cell.SetInt64(aveAge20)
		cell = row.AddCell()
		cell.SetInt64(aveAge30)
		cell = row.AddCell()
		cell.SetInt64(aveAge50)
		cell = row.AddCell()
		cell.SetInt64(aveAge100)
	}
	err = xlFile.Save(excelFileName)
	if err != nil {
		fmt.Printf(err.Error())
	}
}
