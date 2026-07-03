package excel

import (
	"bytes"
	"encoding/base64"
	"fmt"
	"math/big"
	"path/filepath"
	"reflect"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func Parse(blob string, filename string) ([]byte, error) {
	fileData, err := base64.StdEncoding.DecodeString(blob)

	if err != nil {
		return nil, fmt.Errorf("failed to decode base64 data: %v", err)
	}

	ext := filepath.Ext(filename)

	if ext == ".xlsx" || ext == ".xlsm" {
		return fileData, nil
	} else {
		return nil, fmt.Errorf("unsupported file type: %s", ext)
	}
}

func ReadBytes[T any](data []byte, sheet string) ([]T, error) {
	reader := bytes.NewReader(data)

	f, err := excelize.OpenReader(reader)

	if err != nil {
		return nil, err
	}

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	return readSheet[T](f, sheet)
}

func ReadBook[T any](filePath string, sheet string) ([]T, error) {
	f, err := excelize.OpenFile(filePath)

	if err != nil {
		return nil, err
	}

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	return readSheet[T](f, sheet)
}

func WriteBook[T any](file string, sheet string, items []T) {

	f := excelize.NewFile()

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	sw, err := f.NewStreamWriter(sheet)

	if err != nil {
		panic(err)
	}

	err = saveStruct(f, sw, items)

	if err != nil {
		panic(err)
	}

	// Save spreadsheet by the given path.
	if err := f.SaveAs(file); err != nil {
		fmt.Println(err)
	}
}

func readSheet[T any](f *excelize.File, sheet string) ([]T, error) {
	// Get all the rows in the Sheet1.
	rows, err := f.GetRows(sheet)

	if err != nil {
		return nil, err
	}

	maxSize := len(rows[0])

	for i, row := range rows {
		rowSize := len(row)

		if rowSize >= maxSize {
			continue
		}

		padding := make([]string, maxSize-rowSize)

		rows[i] = append(rows[i], padding...)
	}

	return loadStruct[T](rows)
}

func getColumns[T any]() ([]string, []any, error) {
	t := reflect.TypeFor[T]()

	if t.Kind() != reflect.Struct {
		return nil, nil, fmt.Errorf("%s must be a struct", t.Name())
	}

	keys := []string{}
	columns := make([]any, 0)

	for field := range t.Fields() {
		if field.PkgPath != "" {
			continue
		}

		if jsonTag, ok := field.Tag.Lookup("json"); ok {

			jsonColumn, _, _ := strings.Cut(jsonTag, ",")

			if len(jsonColumn) > 0 {
				keys = append(keys, field.Name)
				columns = append(columns, jsonColumn)
			}
		}
	}

	if len(columns) == 0 {
		return nil, nil, fmt.Errorf("No Columns found in struct '%s'. Please include the missing json tags", t.Name())
	}

	return keys, columns, nil
}

func saveStruct[T any](f *excelize.File, sw *excelize.StreamWriter, items []T) error {

	if reflect.TypeFor[T]().Kind() != reflect.Struct {
		return fmt.Errorf("T must be a struct")
	}

	keys, columns, err := getColumns[T]()

	if err != nil {
		return err
	}

	cell := "A" + strconv.Itoa(1)

	if err := sw.SetRow(cell, columns); err != nil {
		return err
	}

	if len(items) == 0 {
		return nil
	}

	size := len(columns)

	row := make([]any, size)

	for i, item := range items {

		reflectValue := reflect.ValueOf(&item).Elem()

		row = row[:size]

		for j, key := range keys {

			fieldValue := reflectValue.FieldByName(key)

			valueType := fieldValue.Type()

			if isDuration(valueType) {

				duraration, err := durationCell(f, fieldValue.Interface().(time.Duration))

				if err != nil {
					return fmt.Errorf("Cannot Convert value '%v' of column %s to time.Duration.", fieldValue, columns[j])
				}

				row[j] = duraration
				// row[j] = formatHHMMSS(fieldValue.Interface().(time.Duration))

				continue
			}

			row[j] = fieldValue.Interface()

		}

		cell = "A" + strconv.Itoa(i+2)

		if err := sw.SetRow(cell, row); err != nil {
			return err
		}
	}

	return sw.Flush()
}

func loadStruct[T any](values [][]string) ([]T, error) {

	if reflect.TypeFor[T]().Kind() != reflect.Struct {
		return nil, fmt.Errorf("T must be a struct")
	}

	if len(values) == 0 {
		return []T{}, nil
	}

	headers := values[0]

	headerIndex := make(map[string]int)

	for i, header := range headers {
		trimedHeader := strings.TrimSpace(header)

		if trimedHeader == "" {
			continue
		}

		headerIndex[strings.TrimSpace(header)] = i
	}

	result := make([]T, 0, len(values)-1)

	for _, row := range values[1:] {
		var item T

		v := reflect.ValueOf(&item).Elem()
		t := v.Type()

		for i := 0; i < t.NumField(); i++ {
			structField := t.Field(i)
			valueField := v.Field(i)

			if !valueField.CanSet() {
				continue
			}

			jsonTag := structField.Tag.Get("json")
			columnName, _, _ := strings.Cut(jsonTag, ",")

			if columnName == "" {
				continue
			}

			columnIndex, exists := headerIndex[columnName]

			if !exists {
				return nil, fmt.Errorf("Missing '%s' column in file.", columnName)
			}

			if columnIndex >= len(row) {
				continue
			}

			value := row[columnIndex]

			valueType := valueField.Type()

			if isTime(valueType) {
				// TODO: NOT WORKING
				parsedTime, err := time.Parse("02:09:00", value)

				if err != nil {
					return nil, fmt.Errorf("Cannot Convert value '%v' of column %s to time.Time.", value, columnName)
				}

				valueField.Set(reflect.ValueOf(parsedTime))

				continue
			} else if isDuration(valueType) {
				parseDuration, err := parseHHMMSS(value)

				if err != nil {
					return nil, fmt.Errorf("Cannot Convert value '%v' of column %s to time.Duration.", value, columnName)
				}

				valueField.Set(reflect.ValueOf(parseDuration))

				continue
			}

			switch valueField.Kind() {
			case reflect.String:
				valueField.SetString(value)

			case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
				if value == "" {
					valueField.SetInt(0)
					continue
				}

				valueInt, err := strconv.ParseInt(value, 10, bitSize(valueField))

				if err != nil {
					return nil, fmt.Errorf("Cannot Convert value '%v' of column %s to int.", value, columnName)
				}

				valueField.SetInt(valueInt)

			case reflect.Float32, reflect.Float64:
				if value == "" {
					valueField.SetFloat(0)
					continue
				}

				valueFloat, err := strconv.ParseFloat(value, bitSize(valueField))

				if err != nil {
					return nil, fmt.Errorf("Cannot Convert value '%v' of column %s to float.", value, columnName)
				}

				valueField.SetFloat(valueFloat)

			case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
				if value == "" {
					valueField.SetUint(0)
					continue
				}

				valueUint, err := strconv.ParseUint(value, 10, bitSize(valueField))

				if err != nil {
					return nil, fmt.Errorf("Cannot Convert value '%v' of column %s to uint.", value, columnName)
				}

				valueField.SetUint(valueUint)

			case reflect.Bool:
				if value == "" {
					valueField.SetBool(false)
					continue
				}

				valueBool, err := strconv.ParseBool(value)

				if err != nil {
					return nil, fmt.Errorf("Cannot Convert value '%v' of column %s to bool.", value, columnName)
				}

				valueField.SetBool(valueBool)

			default:

				if isBigFloat(valueType) {

					if value == "" {
						valueField.Set(reflect.ValueOf(big.NewFloat(0)))
						continue
					}

					precision := uint(len(value[strings.Index(value, ".")+1:]))

					valueFloat, _, err := big.ParseFloat(value, 10, precision*4, big.RoundingMode(big.Exact))

					if err != nil {
						return nil, fmt.Errorf("Cannot Convert value '%v' of column %s to uint.", value, columnName)
					}

					valueField.Set(reflect.ValueOf(valueFloat))

				}
			}
		}

		result = append(result, item)
	}

	return result, nil
}

func bitSize(reflectValue reflect.Value) int {

	return int(reflectValue.Type().Size() * 8)
}

func isBigFloat(t reflect.Type) bool {
	tPtrBigFloat := reflect.TypeFor[*big.Float]()

	return t == tPtrBigFloat
}

func isTime(t reflect.Type) bool {
	tTimeDuration := reflect.TypeFor[time.Time]()

	return t == tTimeDuration
}

func isDuration(t reflect.Type) bool {
	tTimeDuration := reflect.TypeFor[time.Duration]()

	return t == tTimeDuration
}

func parseHHMMSS(input string) (time.Duration, error) {
	if input == "" {
		return 0, nil
	}

	var h, m, s int

	count := strings.Count(input, ":")

	switch count {
	case 2:
		_, err := fmt.Sscanf(input, "%d:%d:%d", &h, &m, &s)

		if err != nil {
			return 0, err
		}

	case 1:
		_, err := fmt.Sscanf(input, "%d:%d", &h, &m)

		if err != nil {
			return 0, err
		}

	default:
		_, err := fmt.Sscanf(input, "%d", &h)

		if err != nil {
			return 0, err
		}
	}

	return time.Duration(h)*time.Hour +
		time.Duration(m)*time.Minute +
		time.Duration(s)*time.Second, nil
}

func formatHHMMSS(d time.Duration) string {
	totalSeconds := int64(d.Seconds())

	hours := totalSeconds / 3600
	minutes := (totalSeconds % 3600) / 60
	seconds := totalSeconds % 60

	return fmt.Sprintf("%02d:%02d:%02d", hours, minutes, seconds)
}

func durationToExcelTime(d time.Duration) float64 {
	return float64(d) / float64(24*time.Hour)
}

func durationCell(f *excelize.File, d time.Duration) (excelize.Cell, error) {
	format := "hh:mm:ss" // use "hh:mm:ss" if you never go above 24h

	styleID, err := f.NewStyle(&excelize.Style{
		CustomNumFmt: &format,
	})

	if err != nil {
		return excelize.Cell{}, err
	}

	return excelize.Cell{
		StyleID: styleID,
		Value:   durationToExcelTime(d),
	}, nil
}
