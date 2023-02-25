package calix

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

func TestExcel(t *testing.T) {
	type person struct {
		Name       string `excel:"title=姓名"`
		Age        int
		Birthday   time.Time       `excel:"title=B-Day,timestamp"`
		Salary     decimal.Decimal `excel:"title=工资"`
		EntryTime  int64           `excel:"title=xy,timestamp,omit"`
		IsDemo     isdemo          `excel:"title=是否测试帐号"`
		PureNumber int64           `excel:"title=纯数字"`
	}

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
