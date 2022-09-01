package excel

import (
	"errors"
	"fmt"
	"io"
	"reflect"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"

	"github.com/shopspring/decimal"
	"github.com/siddontang/go/log"
	"github.com/xuri/excelize/v2"
	"golang.org/x/exp/slices"
)

// Exporter example
// type isdemo bool
//
// func (d isdemo) ExcelString() string {
// 	if d {
// 		return "测试帐号"
// 	}
// 	return "正式帐号"
// }

type ExcelExport interface {
	ExcelString() string
}

type Exporter struct {
	timezone       *time.Location
	sheetname      string
	decimalDigits  *int32
	datetimeformat string
}

// default sheetname is "Sheet1"
// default timezone is "Asia/Hong_Kong"
// default decimal digits is 2
// default datetime format is "2006-01-02 15:04:05"

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
	if e.datetimeformat == "" {
		e.datetimeformat = "2006-01-02 15:04:05"
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

func WithDecimalDigits(digits int32) func(*Exporter) {
	return func(e *Exporter) {
		e.decimalDigits = &digits
	}
}

func WithSheetname(sheetname string) func(*Exporter) {
	return func(e *Exporter) {
		e.sheetname = sheetname
	}
}

// use standard time format, a mulformed datetime format will cause wrong output without warning
func WithDatetimeFormat(format string) func(*Exporter) {
	return func(e *Exporter) {
		e.datetimeformat = format
	}
}

// data must be a non empty slice of struct
func (e *Exporter) Export(data interface{}) (io.Reader, error) {

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

	file := excelize.NewFile()

	file.SetSheetName("Sheet1", e.sheetname)
	//write header row
	err := file.SetSheetRow(e.sheetname, "A1", &Header)
	if err != nil {
		return nil, fmt.Errorf("error set sheet header row: %v", err)
	}

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
					switch {
					case fieldKind <= reflect.Int64 && fieldKind >= reflect.Int:
						row[j] = time.Unix(fieldValue.Int(), 0).In(e.timezone).Format(e.datetimeformat)
						//case time.Time
					case fieldTypeOf == reflect.TypeOf(time.Time{}):
						row[j] = fieldValue.Interface().(time.Time).In(e.timezone).Format(e.datetimeformat)

					}

				default:
					row[j] = fieldValue.Interface()
				}
			}
		}

		err := file.SetSheetRow(e.sheetname, "A"+strconv.Itoa(i+2), &row)
		if err != nil {
			return nil, fmt.Errorf("error set data row %d: %v", i, err)
		}
	}

	//remove omited columns
	offset := 0

	for _, col := range omitFields {

		column := string(rune(col + 65 - offset))

		err := file.RemoveCol(e.sheetname, column)

		if err != nil {
			return nil, fmt.Errorf("error remove column %s: %v", column, err)
		}

		offset++
	}

	//set column width
	cols, err := file.GetCols(e.sheetname)

	if err != nil {
		return nil, err
	}

	for idx, col := range cols {
		var largestWidth float64

		for _, rowCell := range col {

			var cellWidth float64
			r := utf8.RuneCountInString(rowCell)
			l := len(rowCell)
			if r != l {
				//utf8 chars 's width x 1.5 + 2 margin
				cellWidth = float64(r)*1.5 + 2
			} else {
				cellWidth = float64(l) + 2
			}

			if cellWidth > largestWidth {
				largestWidth = cellWidth
			}
		}

		name, err := excelize.ColumnNumberToName(idx + 1)

		if err != nil {
			return nil, err
		}

		err = file.SetColWidth(e.sheetname, name, name, float64(largestWidth))
		if err != nil {
			return nil, fmt.Errorf("error set column width %d: %v", idx+1, err)
		}
	}

	ioR, err := file.WriteToBuffer()

	return ioR, err

}
