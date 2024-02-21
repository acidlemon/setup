package gosheets

import (
	"context"
	"encoding/base64"
	"fmt"
	"log"
	"net/http"
	"os"

	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

type SheetService struct {
	*sheets.Service
}

func NewSheetService() *SheetService {
	// Create Service using ADC
	ctx := context.Background()
	srv, err := sheets.NewService(ctx)
	if err == nil {
		log.Println("New Service by GOOGLE_APPLICATION_CREDENTIALS")
		return &SheetService{srv}
	} else {
		log.Println(err)
	}

	// Fallback: Base64 encoded Credential JSON from GOOGLE_APPLICATION_CREDENTIALS_BASE64_JSON
	client := prepareClient()
	srv, err = sheets.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	}

	return &SheetService{srv}
}

func (ss *SheetService) SheetRange(sheet, getRange string) string {
	return fmt.Sprintf("%s!%s", sheet, getRange)
}

func (ss *SheetService) Get(key string, sheetRange string) [][]interface{} {
	return sheetGet(ss.Service, key, sheetRange)
}

func (ss *SheetService) BatchGet(key string, sheetRanges []string) [][][]interface{} {
	return sheetBatchGet(ss.Service, key, sheetRanges)
}

// ADCを使わない場合のクライアント生成
func prepareClient() *http.Client {
	src := os.Getenv("GOOGLE_APPLICATION_CREDENTIALS_BASE64_JSON")
	if src == "" {
		log.Fatal(`env GOOGLE_APPLICATION_CREDENTIALS_BASE64_JSON is empty ('-'#)`)
	}
	cred, err := base64.StdEncoding.DecodeString(src)
	if err != nil {
		log.Fatalf(`failed to DecodeString(): err=%s`, err.Error())
	}
	conf, err := google.JWTConfigFromJSON([]byte(cred), "https://www.googleapis.com/auth/spreadsheets")
	if err != nil {
		log.Fatalf(`failed to prepare client: %s`, err.Error())
	}
	return conf.Client(context.Background())
}

func sheetGet(srv *sheets.Service, key string, sheetRange string) [][]interface{} {
	valueRenderOption := "UNFORMATTED_VALUE"
	res, err := srv.Spreadsheets.Values.Get(key, string(sheetRange)).ValueRenderOption(valueRenderOption).Do()
	if err != nil {
		log.Printf("Unable to retrieve data from sheet: %v", err)
		return nil
	}

	return res.Values
}

func sheetBatchGet(srv *sheets.Service, key string, sheetRanges []string) [][][]interface{} {
	valueRenderOption := "UNFORMATTED_VALUE"
	res, err := srv.Spreadsheets.Values.BatchGet(key).Ranges(sheetRanges...).ValueRenderOption(valueRenderOption).Do()
	if err != nil {
		log.Printf("Unable to retrieve data from sheet: %v", err)
		return nil
	}

	values := [][][]interface{}{}
	for _, v := range res.ValueRanges {
		values = append(values, v.Values)
	}

	return values
}
