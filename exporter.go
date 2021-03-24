package exportizer

import (
	"errors"
	"fmt"
	"io"
	"net/http"
	"reflect"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)


type Config struct {

}
type ExcelExporter struct {
	xlsx   *excelize.File
	sheets map[string]*Sheet
	styles []int
}

func NewExcelExporter() *ExcelExporter {
	xlsx := excelize.NewFile()

	s1,err:=xlsx.NewStyle(`{"fill":{"type":"pattern","color":["#F0F0FF"],"pattern":1}}`)
	if err != nil {
		panic(err)
	}
	return &ExcelExporter{
		xlsx:   xlsx,
		sheets: make(map[string]*Sheet, 0),
		styles: []int{s1},
	}
}

func (e *ExcelExporter) SaveToFile(filename string) error {
	for s, sheet := range e.sheets {
		err := sheet.Close()
		if err != nil {
			return fmt.Errorf("%s close err:%w",s,err)
		}
	}
	err := e.xlsx.SaveAs(filename)
	if err != nil {
		return err
	}
	return nil
}

func (e *ExcelExporter) SaveToWriter(filename string,w http.ResponseWriter)error  {
	for s, sheet := range e.sheets {
		err := sheet.Close()
		if err != nil {
			return fmt.Errorf("%s close err:%w",s,err)
		}
	}
	header := w.Header()
	header.Set("Content-Type","application/octet-stream")
	header.Set("Content-Disposition",fmt.Sprintf("attachment; filename=%s",filename))
	header.Set("Content-Transfer-Encoding","binary")
	_, err := e.xlsx.WriteTo(w)
	if err != nil {
		return err
	}
	return nil
}

func (e *ExcelExporter) ExportToReader() (io.Reader,error) {
	for s, sheet := range e.sheets {
		err := sheet.Close()
		if err != nil {
			return nil,fmt.Errorf("%s close err:%w",s,err)
		}
	}
	return e.xlsx.WriteToBuffer()
}

type Sheet struct {
	exporter *ExcelExporter
	name     string
	sheet    int
	row      uint64
	column   int32
	writer   *excelize.StreamWriter
}

func (e *ExcelExporter) NewSheet(sheet string) (*Sheet, error) {
	s := e.xlsx.NewSheet(sheet)
	writer, err := e.xlsx.NewStreamWriter(sheet)
	if err != nil {
		return nil, err
	}
	item := &Sheet{
		name:     sheet,
		writer:   writer,
		exporter: e,
		sheet:    s,
		row:      0,
		column:   0,
	}
	e.sheets[sheet] = item
	return item, nil
}

func (s *Sheet) AddRow(slice interface{}) error {
	s.row += 1
	sValues := reflect.ValueOf(slice)
	if sValues.Kind() != reflect.Array && sValues.Kind() != reflect.Slice {
		return errors.New("expect type of array || slice")
	}
	var cells = make([]interface{}, 0, sValues.Len())
	for i := 0; i < sValues.Len(); i++ {
		cells = append(cells, &excelize.Cell{
			StyleID: s.exporter.styles[0],
			Value:   sValues.Index(i).Interface(),
		})
	}
	err := s.writer.SetRow(fmt.Sprintf("A%d", s.row), cells)
	if err != nil {
		return err
	}
	return nil
}

func (s *Sheet) Close() error {
	return s.writer.Flush()
}
