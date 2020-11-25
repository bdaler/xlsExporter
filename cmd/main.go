package main

import (
	"github.com/360EntSecGroup-Skylar/excelize"
	"log"
	"strconv"
)

var headerAxis = [...]string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
	"N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

func main() {
	f := excelize.NewFile()

	headerStyle, err := f.NewStyle(`{"font":{"bold":true}}`)
	if err != nil {
		log.Println("headerStyle err: ",err)
	}

	headers := [...]string{"Наименование перевозчика", "Номер", "Рег номер", "Название", "Вид тарифа", "Вид маршрута",
		"Количество рейсов", "Номер ПН", "План", "Факт", "Дельта", "План", "Факт", "Дельта"}

	rows := [...][]string{
		{"row data 1", "row data 1_2", "row data 1_3", "row data 1_4", "row data 1_5"},
		{"row data 2", "row data 2_2", "row data 2_3", "row data 2_4", "row data 2_5", "row data 2_6", "row data 2_7", "row data 2_8", "row data 2_9", "row data 2_10", "row data 2_11", "row data 2_12", "row data 2_13", "row data 2_14"},
		{"row data 3", "row data 3_2", "row data 3_3", "row data 3_4", "row data 3_5"},
		{"row data 4", "row data 4_2", "row data 4_3", "row data 4_4", "row data 4_5"},
		{"row data 5", "row data 5_2", "row data 5_3", "row data 5_4", "row data 5_5"},
		{"row data 6", "row data 6_2", "row data 6_3", "row data 6_4", "row data 6_5"},
		{"row data 7", "row data 7_2", "row data 7_3", "row data 7_4", "row data 7_5"},
		{"row data 8", "row data 8_2", "row data 8_3", "row data 8_4", "row data 8_5"},
		{"row data 9", "row data 9_2", "row data 9_3", "row data 9_4", "row data 9_5"},
		{"row data 10", "row data 10_2", "row data 10_3", "row data 10_4", "row data 10_5"},
		{"row data 11", "row data 11_2", "row data 11_3", "row data 11_4", "row data 11_5"},
	}

	for i, header := range headers {
		f.SetCellValue("Sheet1", headerAxis[i] + strconv.Itoa(1), header)
		f.SetCellStyle("Sheet1", headerAxis[i] + strconv.Itoa(1), headerAxis[i] + strconv.Itoa(1), headerStyle)
	}

	axis := 2
	for _, row := range rows {
		for i, data := range row {
			f.SetCellValue("Sheet1", headerAxis[i] + strconv.Itoa(axis), data)
		}
		axis++
	}
	if err := f.SaveAs("Book1.xlsx"); err != nil{
		log.Println("err to write file", err)
	}
}
