package excel

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
func TestExcel(t *testing.T) {
	type person struct {
		Name      string `excel:"title=姓名"`
		Age       int
		Birthday  int //`excel:"title=xy,timestamp"`
		Salary    decimal.Decimal
		EntryTime int    `excel:"title=xy,timestamp"`
		IsDemo    isdemo `excel:"title=是否测试帐号"`
	}

	data := []person{
		{
			Name:      "张三",
			Age:       18,
			Birthday:  int(time.Now().Unix()),
			Salary:    decimal.NewFromFloat(1000),
			EntryTime: int(time.Now().Unix()),
		},
		{
			Name:      "李四",
			Age:       19,
			Birthday:  int(time.Now().Unix()),
			Salary:    decimal.NewFromFloat(2000.03120),
			EntryTime: int(time.Now().Unix()),
		},
	}
	// data := []person{}

	// data := person{
	// 	Name:      "张三",
	// 	Age:       18,
	// 	Birthday:  time.Date(2000, 1, 1, 0, 0, 0, 0, time.Local),
	// 	Salary:    decimal.NewFromFloat(1000.00),
	// 	EntryTime: time.Now().Unix(),
	// }
	// data := []int{5, 8, 9}
	excel := New(WithTimezone("Asia/Hong_Kong"), WithSheetname("报表2"))
	reader, err := excel.Export(data)
	if err != nil {
		t.Error(err)
	}
	// t.Log(reader)
	//write reader to file
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

}
