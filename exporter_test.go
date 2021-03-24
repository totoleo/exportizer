package exportizer

import "testing"

func TestNew(t *testing.T) {
	exporter := NewExcelExporter()

	sheet, err := exporter.NewSheet("Sheet1")
	if err != nil {
		t.Error(err)
		return
	}
	err = sheet.AddRow([]int{1, 2, 3})
	if err != nil {
		t.Error(err)
		return
	}

	err = exporter.SaveToFile("test.xlsx")
	if err != nil {
		t.Error(err)
	}
}
