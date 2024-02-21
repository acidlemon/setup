package gosheets

import (
	"encoding/json"
	"time"
)

func SheetToMap(sheet [][]interface{}) []map[string]interface{} {
	// Get header
	header := sheet[0]
	// Get data
	data := sheet[1:]

	// Convert to map
	var result []map[string]interface{}
	for _, row := range data {
		m := make(map[string]interface{})
		for i, cell := range row {
			m[header[i].(string)] = cell
		}
		result = append(result, m)
	}
	return result
}

type SerialTime struct {
	f float64
}

func (s SerialTime) Time(loc *time.Location) time.Time {
	return SerialToTime(float64(s.f), loc)
}

func (s *SerialTime) UnmarshalJSON(b []byte) error {
	var f float64
	if err := json.Unmarshal(b, &f); err != nil {
		return err
	}
	s.f = f
	return nil
}

func SerialToTime(serial float64, loc *time.Location) time.Time {
	d := time.Unix(int64((serial-25569)*86400), 0).UTC()
	return time.Date(d.Year(), d.Month(), d.Day(), d.Hour(), d.Minute(), d.Second(), 0, loc)
}
