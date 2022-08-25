package excel

import (
	"errors"
	"fmt"
	"io"
	"reflect"
	"strconv"
	"strings"
	"time"

	"github.com/shopspring/decimal"
	"github.com/siddontang/go/log"
	"github.com/xuri/excelize/v2"
	"golang.org/x/exp/slices"
)

type ExcelExport interface {
	ExcelString() string
}

type IsDemoAccount bool

func (d IsDemoAccount) ExcelString() string {
	if d {
		return "测试帐号"
	}
	return "正式帐号"
}

type Exporter struct {
	timezone      *time.Location
	sheetname     string
	decimalDigits *int32
}

func New(options ...func(*Exporter)) *Exporter {
	e := &Exporter{}
	for _, option := range options {
		option(e)
	}
	if e.timezone == nil {
		e.timezone = time.Local
	}
	if e.sheetname == "" {
		e.sheetname = "Sheet1"
	}
	if e.decimalDigits == nil {
		e.decimalDigits = new(int32)
		*e.decimalDigits = 2
	}
	return e
}
func WithTimezone(timezone string) func(*Exporter) {
	return func(e *Exporter) {
		loc, err := time.LoadLocation(timezone)
		if err != nil {
			log.Errorf("error load timezone %s : %v", timezone, err)
		} else {
			e.timezone = loc
		}
	}
}

func WithSheetname(sheetname string) func(*Exporter) {
	return func(e *Exporter) {
		e.sheetname = sheetname
	}
}

func (e *Exporter) Export(data interface{}) (io.Reader, error) {
	// loc, _ := time.LoadLocation("Asia/Hong_Kong")
	// sheetname := "报表(导出时间:" + time.Now().In(e.timezone).Format("2006-01-02 15:04:05") + ")"
	// excelExportType := reflect.TypeOf((ExcelExport)(nil)).Elem()

	file := excelize.NewFile()

	file.SetSheetName("Sheet1", e.sheetname)

	// // check if the data is a slice
	if reflect.TypeOf(data).Kind() != reflect.Slice {
		return nil, errors.New("data must be a slice")
	}
	//type of the underline struct
	refType := reflect.TypeOf(data).Elem()

	//value of the underline struct
	refValue := reflect.ValueOf(data)

	// check if data is empty
	if refValue.Len() == 0 {
		return nil, errors.New("data is empty")
	}

	//check if the element is a struct
	if refType.Kind() != reflect.Struct {
		return nil, errors.New("element of data must be struct")
	}

	//number of fields in the struct
	fieldCount := refType.NumField()

	//construct header row
	Header := make([]interface{}, fieldCount)
	var tsFields []int   //position of timestamp fields
	var omitFields []int //position of omit fields

	/*
		tags example "excel:omit,timestamp,title=colume title"
	*/

	for i := 0; i < fieldCount; i++ {
		//analyze every field tag of the struct
		field := refType.Field(i)
		fieldName := field.Name

		if excelTag, ok := field.Tag.Lookup("excel"); ok {
			tags := strings.Split(excelTag, ",")
			for _, tag := range tags {
				fmt.Println("tagis:", tag)
				if tag == "omit" {
					omitFields = append(omitFields, i)
				}
				if tag == "timestamp" {
					tsFields = append(tsFields, i)
				}
				if strings.HasPrefix(tag, "title=") {
					fieldName = tag[6:]
				}
			}

		}

		Header[i] = fieldName
	}

	//write header row
	file.SetSheetRow(e.sheetname, "A1", &Header)

	//write rows
	for i := 0; i < refValue.Len(); i++ {
		row := make([]interface{}, fieldCount)
		for j := 0; j < fieldCount; j++ {

			if !slices.Contains(omitFields, j) {
				//get the value of the field
				fieldValue := refValue.Index(i).Field(j)
				fieldTypeOf := reflect.TypeOf(fieldValue.Interface())
				fieldKind := fieldTypeOf.Kind()
				t := reflect.TypeOf((*ExcelExport)(nil)).Elem()
				switch {
				case fieldValue.Type().Implements(t):
					row[j] = fieldValue.Interface().(ExcelExport).ExcelString()
				case fieldKind == reflect.String:
					row[j] = fieldValue.String()
				case fieldTypeOf == reflect.TypeOf(decimal.Decimal{}):
					row[j] = fieldValue.Interface().(decimal.Decimal).StringFixed(*e.decimalDigits)
				case slices.Contains(tsFields, j):
					row[j] = time.Unix(fieldValue.Int(), 0).In(e.timezone).Format("2006-01-02 15:04:05")

				default:
					row[j] = fieldValue.Interface()
				}

				// if slices.Contains(tsFields, j) {
				// 	row[j] = time.Unix(refValue.Index(i).Field(j).Interface().(int64), 0).In(e.timezone).Format("2006-01-02 15:04:05")
				// } else {

				// 	typeof := reflect.TypeOf(refValue.Index(i).Field(j).Interface())

				// 	if typeof == reflect.TypeOf(decimal.Decimal{}) {

				// 		row[j] = refValue.Index(i).Field(j).Interface().(decimal.Decimal).InexactFloat64() //what if float64 is not exact?

				// 	} else if typeof.ConvertibleTo(excelExportType) {
				// 		row[j] = refValue.Index(i).Field(j).Interface().(ExcelExport).ExcelString()

				// 	} else {
				// 		row[j] = refValue.Index(i).Field(j).Interface()
				// 	}

				// }
			}
		}

		file.SetSheetRow(e.sheetname, "A"+strconv.Itoa(i+2), &row)
	}

	//remove omited columns
	offset := 0

	for _, col := range omitFields {

		column := string(rune(col + 65 - offset))

		file.RemoveCol(e.sheetname, column)

		offset++
	}

	//format column width
	cols, err := file.GetCols(e.sheetname)

	if err != nil {
		return nil, err
	}

	for idx, col := range cols {
		largestWidth := 0

		for _, rowCell := range col {

			// cellWidth := utf8.RuneCountInString(rowCell) + 2 // + 2 for margin

			cellWidth := len(rowCell) + 2 //chinese char is counted as 3

			if cellWidth > largestWidth {
				largestWidth = cellWidth
			}
		}

		name, err := excelize.ColumnNumberToName(idx + 1)

		if err != nil {
			return nil, err
		}

		file.SetColWidth(e.sheetname, name, name, float64(largestWidth))
	}

	ioR, err := file.WriteToBuffer()

	return ioR, err

}
