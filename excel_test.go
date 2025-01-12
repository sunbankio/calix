package excelwriter

import (
	"bufio"
	"io"
	"os"
	"testing"
	"time"

	"github.com/shopspring/decimal"
)

type isdemo bool

func (d isdemo) ExcelString() string {
	if d {
		return "测试帐号"
	}
	return "正式帐号"
}

func (d isdemo) ExcelColor() string {
	if d {
		return "red"
	}
	return "green"
}

type person struct {
	Name       string `excel:"title=姓名"`
	Age        int
	Birthday   time.Time       `excel:"title=B-Day,timestamp"`
	Salary     decimal.Decimal `excel:"title=工资"`
	EntryTime  int64           `excel:"title=xy,timestamp,omit"`
	IsDemo     isdemo          `excel:"title=是否测试帐号"`
	PureNumber int64           `excel:"title=纯数字"`
}

func TestExcel(t *testing.T) {
	data := []person{
		{
			Name:       "张三",
			Age:        18,
			Birthday:   time.Now(),
			Salary:     decimal.NewFromFloat(1000),
			EntryTime:  time.Now().Unix(),
			PureNumber: 123456789,
		},
		{
			Name:       "李四",
			Age:        19,
			Birthday:   time.Now(),
			Salary:     decimal.NewFromFloat(2000.03120),
			EntryTime:  time.Now().Unix(),
			PureNumber: 123789,
		},
	}

	excel := New(WithTimezone("Asia/Hong_Kong"), WithSheetName("报表2"), WithDatetimeFormat("2006年01月02日15:04:05"))
	reader, err := excel.Export(data)
	if err != nil {
		t.Error(err)
	}
	// t.Log(reader)
	//write reader to file
	_ = WriteToFile(reader, "test.xlsx")

}

type ValidationTestStruct struct {
	Name     string          `excel:"title=Name;validation={'type':'textLength','operator':'between','formula1':'5','formula2':'10'}"`
	Price    decimal.Decimal `excel:"title=Price;validation={'type':'decimal','operator':'between','formula1':'10','formula2':'100'}"`
	Category string          `excel:"title=Category;validation={'type':'list','allowList':['A','B','C']}"`
}

func TestValidation(t *testing.T) {
	data := []ValidationTestStruct{
		{
			// Valid data
			Name:     "John Doe",    // 8 chars - valid
			Price:    decimal.NewFromFloat(50.0), // valid
			Category: "A",           // valid
		},
		{
			// Invalid data but will still be written
			Name:     "Bob",         // 3 chars - too short
			Price:    decimal.NewFromFloat(5.0),  // too low
			Category: "D",           // not in list
		},
	}

	excel := New(
		WithSheetName("Validation Test"),
	)

	reader, err := excel.Export(data)
	if err != nil {
		t.Fatal(err)
	}

	// Write to file for manual inspection
	if err := WriteToFile(reader, "validation_test.xlsx"); err != nil {
		t.Fatal(err)
	}
}

func WriteToFile(reader io.Reader, filename string) error {
	file, err := os.Create(filename)
	if err != nil {
		return err
	}
	defer file.Close()

	writer := bufio.NewWriter(file)
	defer writer.Flush()

	_, err = io.Copy(writer, reader)
	return err
}

/*
f, _ := os.OpenFile("test.xlsx", os.O_CREATE|os.O_WRONLY, 0666)
	w := bufio.NewWriter(f)

	// make a buffer to keep chunks that are read
	buf := make([]byte, 1024)
	for {
		// read a chunk
		n, err := reader.Read(buf)
		if err != nil && err != io.EOF {
			panic(err)
		}
		if n == 0 {
			break
		}

		// write a chunk
		if _, err := w.Write(buf[:n]); err != nil {
			panic(err)
		}
	}

	if err = w.Flush(); err != nil {
		panic(err)
	}

*/
