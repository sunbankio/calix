package excelwriter

import (
	"encoding/json"
	"fmt"
	"io"
	"reflect"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"

	"github.com/shopspring/decimal"
	"github.com/xuri/excelize/v2"
	"golang.org/x/exp/slices"
)

// CellStyle represents styling options for a cell
type CellStyle struct {
	FontSize     float64
	FontColor    string
	Background   string
	Bold         bool
	Italic       bool
	Alignment    string // left, center, right
	NumberFormat string
}

// ColumnConfig holds configuration for a column
type ColumnConfig struct {
	Title      string
	Width      float64
	Style      *CellStyle
	Format     string // For dates, numbers, etc.
	Hidden     bool
	Freeze     bool // For freezing columns
	Validation *DataValidation
}

// DataValidation represents data validation rules
type DataValidation struct {
	Type      string // decimal, whole, date, time, textLength, list
	Operator  string // between, notBetween, equal, notEqual, greaterThan, lessThan, greaterOrEqual, lessOrEqual
	Formula1  string
	Formula2  string
	AllowList []string // For dropdown lists
}

type Export interface {
	ExcelString() string
	ExcelStyle() *CellStyle // New method for custom styling
}

type Exporter struct {
	timezone       *time.Location
	sheetName      string
	decimalDigits  *int32
	datetimeFormat string
	defaultStyle   *CellStyle
	headerStyle    *CellStyle
	columnConfigs  map[string]*ColumnConfig
	password       string
	autoFilter     bool
	freezeHeader   bool
	activeSheet    bool // Make the sheet active when opened
}

// ExcelTag represents the structure of excel tag options
type ExcelTag struct {
	Title      string
	Width      float64
	Format     string
	Hidden     bool
	Omit       bool
	Timestamp  bool
	Style      *CellStyle
	Validation *DataValidation
	Freeze     bool
	ValueMap   map[string]string // Add value mapping support
}

func parseTag(tag string) ExcelTag {
	var et ExcelTag
	parts := strings.Split(tag, ";")

	for _, part := range parts {
		kv := strings.SplitN(part, "=", 2)
		key := strings.TrimSpace(kv[0])

		if len(kv) == 1 {
			switch key {
			case "omit":
				et.Omit = true
			case "timestamp":
				et.Timestamp = true
			case "freeze":
				et.Freeze = true
			case "hidden":
				et.Hidden = true
			}
			continue
		}

		value := strings.TrimSpace(kv[1])
		switch key {
		case "title":
			et.Title = value
		case "width":
			if w, err := strconv.ParseFloat(value, 64); err == nil {
				et.Width = w
			}
		case "format":
			et.Format = value
		case "style":
			var style CellStyle
			if err := json.Unmarshal([]byte(value), &style); err == nil {
				et.Style = &style
			}
		case "validation":
			var validation DataValidation
			if err := json.Unmarshal([]byte(value), &validation); err == nil {
				et.Validation = &validation
			}
		case "valuemap":
			var valueMap map[string]string
			if err := json.Unmarshal([]byte(value), &valueMap); err == nil {
				et.ValueMap = valueMap
			}
		}
	}
	return et
}

func (e *Exporter) formatValue(value interface{}, tag ExcelTag) string {
	if tag.ValueMap != nil {
		if strVal, ok := tag.ValueMap[fmt.Sprint(value)]; ok {
			return strVal
		}
	}
	return fmt.Sprint(value)
}

func (e *Exporter) Export(data interface{}) (io.Reader, error) {
	// check if the data is a slice
	if reflect.TypeOf(data).Kind() != reflect.Slice {
		return nil, fmt.Errorf("data is not a slice")
	}

	//type of the underline struct
	refType := reflect.TypeOf(data).Elem()

	//value of the underline struct
	refValue := reflect.ValueOf(data)

	// check if data is empty
	if refValue.Len() == 0 {
		return nil, fmt.Errorf("data is empty")
	}

	//check if the element is a struct
	if refType.Kind() != reflect.Struct {
		return nil, fmt.Errorf("data is not a struct")
	}

	//number of fields in the struct
	fieldCount := refType.NumField()

	// Initialize column configs if not already done
	if e.columnConfigs == nil {
		e.columnConfigs = make(map[string]*ColumnConfig)
	}

	//construct header row and parse field configurations
	Header := make([]interface{}, fieldCount)
	var tsFields []int
	var omitFields []int

	for i := 0; i < fieldCount; i++ {
		field := refType.Field(i)

		if excelTag, ok := field.Tag.Lookup("excel"); ok {
			tag := parseTag(excelTag)

			// Store column config
			e.columnConfigs[field.Name] = &ColumnConfig{
				Title:      tag.Title,
				Width:      tag.Width,
				Style:      tag.Style,
				Format:     tag.Format,
				Hidden:     tag.Hidden,
				Freeze:     tag.Freeze,
				Validation: tag.Validation,
			}

			if tag.Omit {
				omitFields = append(omitFields, i)
				continue
			}
			if tag.Timestamp {
				tsFields = append(tsFields, i)
			}

			Header[i] = tag.Title
			if Header[i] == "" {
				Header[i] = field.Name
			}
		} else {
			Header[i] = field.Name
		}
	}

	file := excelize.NewFile()

	if err := file.SetSheetName("Sheet1", e.sheetName); err != nil {
		return nil, fmt.Errorf("failed to set sheet name: %w", err)
	}

	// Apply column configurations
	for fieldName, config := range e.columnConfigs {
		if config.Hidden {
			if err := file.SetColVisible(e.sheetName, fieldName, false); err != nil {
				return nil, fmt.Errorf("failed to hide column %s: %w", fieldName, err)
			}
		}
		if config.Width > 0 {
			if err := file.SetColWidth(e.sheetName, fieldName, fieldName, config.Width); err != nil {
				return nil, fmt.Errorf("failed to set column width for %s: %w", fieldName, err)
			}
		}
	}

	//write header row
	err := file.SetSheetRow(e.sheetName, "A1", &Header)
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
				t := reflect.TypeOf((*Export)(nil)).Elem()
				switch {
				case fieldValue.Type().Implements(t):
					row[j] = fieldValue.Interface().(Export).ExcelString()
				case fieldKind == reflect.String:
					row[j] = fieldValue.String()
				case fieldTypeOf == reflect.TypeOf(decimal.Decimal{}):
					row[j] = fieldValue.Interface().(decimal.Decimal).StringFixed(*e.decimalDigits)
				case slices.Contains(tsFields, j):
					switch {
					case fieldKind <= reflect.Int64 && fieldKind >= reflect.Int:
						row[j] = time.Unix(fieldValue.Int(), 0).In(e.timezone).Format(e.datetimeFormat)
						//case time.Time
					case fieldTypeOf == reflect.TypeOf(time.Time{}):
						row[j] = fieldValue.Interface().(time.Time).In(e.timezone).Format(e.datetimeFormat)

					}
				case fieldKind == reflect.Bool:
					tag, ok := e.columnConfigs[refType.Field(j).Name]
					if ok {
						excelTag := ExcelTag{
							Title:      tag.Title,
							Width:      tag.Width,
							Format:     tag.Format,
							Hidden:     tag.Hidden,
							Style:      tag.Style,
							Validation: tag.Validation,
							Freeze:     tag.Freeze,
						}
						row[j] = e.formatValue(fieldValue.Interface(), excelTag)
					} else {
						row[j] = fieldValue.Interface()
					}
				default:
					row[j] = fieldValue.Interface()
				}
			}
		}

		err := file.SetSheetRow(e.sheetName, "A"+strconv.Itoa(i+2), &row)
		if err != nil {
			return nil, fmt.Errorf("error set data row %d: %v", i, err)
		}
	}

	//remove omited columns
	offset := 0

	for _, col := range omitFields {

		column := string(rune(col + 65 - offset))

		err := file.RemoveCol(e.sheetName, column)

		if err != nil {
			return nil, fmt.Errorf("error remove column %s: %v", column, err)
		}

		offset++
	}

	//set column width
	cols, err := file.GetCols(e.sheetName)

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

		err = file.SetColWidth(e.sheetName, name, name, float64(largestWidth))
		if err != nil {
			return nil, fmt.Errorf("error set column width %d: %v", idx+1, err)
		}
	}

	ioR, err := file.WriteToBuffer()

	return ioR, err
}

func New(options ...func(*Exporter)) *Exporter {
	e := &Exporter{}
	for _, option := range options {
		option(e)
	}
	if e.timezone == nil {
		e.timezone = time.Local
	}
	if e.sheetName == "" {
		e.sheetName = "Sheet1"
	}
	if e.decimalDigits == nil {
		e.decimalDigits = new(int32)
		*e.decimalDigits = 2
	}
	if e.datetimeFormat == "" {
		e.datetimeFormat = "2006-01-02 15:04:05"
	}
	return e
}

func WithTimezone(timezone string) func(*Exporter) {
	return func(e *Exporter) {
		loc, err := time.LoadLocation(timezone)
		if err != nil {
			fmt.Printf("[excelExporter] invalid timezone: %s, use local timezone instead", timezone)
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

func WithSheetName(sheetName string) func(*Exporter) {
	return func(e *Exporter) {
		e.sheetName = sheetName
	}
}

// WithDatetimeFormat set datetime format use standard time format, a malformed datetime format will cause wrong output without warning
func WithDatetimeFormat(format string) func(*Exporter) {
	return func(e *Exporter) {
		e.datetimeFormat = format
	}
}

func WithDefaultStyle(style *CellStyle) func(*Exporter) {
	return func(e *Exporter) {
		e.defaultStyle = style
	}
}

func WithHeaderStyle(style *CellStyle) func(*Exporter) {
	return func(e *Exporter) {
		e.headerStyle = style
	}
}

func WithPassword(password string) func(*Exporter) {
	return func(e *Exporter) {
		e.password = password
	}
}

func WithAutoFilter(enable bool) func(*Exporter) {
	return func(e *Exporter) {
		e.autoFilter = enable
	}
}

func WithFreezeHeader(enable bool) func(*Exporter) {
	return func(e *Exporter) {
		e.freezeHeader = enable
	}
}

func WithActiveSheet(enable bool) func(*Exporter) {
	return func(e *Exporter) {
		e.activeSheet = enable
	}
}
