package excelwriter

import "errors"

var (
	ErrDataIsNotSlice  = errors.New("data must be a slice")
	ErrDataIsEmpty     = errors.New("data is empty")
	ErrDataIsNotStruct = errors.New("element of data must be struct")
)
