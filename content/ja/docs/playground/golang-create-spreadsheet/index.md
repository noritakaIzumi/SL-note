---
title: "Golang Create Spreadsheet"
description: "Go でスプレッドシートを操作してみる"
lead: "Go でスプレッドシートを操作してみる"
date: 2021-12-28T23:30:49+09:00
lastmod: 2021-12-28T23:30:49+09:00
draft: false
images: []
menu:
  docs:
    parent: "playground"
weight: 999
toc: true
---

---

Go 言語でスプレッドシートを作成することはできるのか？やってみる。

参考にするページ

[Go quickstart | Sheets API | Google Developers](https://developers.google.com/sheets/api/quickstart/go)

[sheets package - google.golang.org/api/sheets/v4 - pkg.go.dev](https://pkg.go.dev/google.golang.org/api/sheets/v4)

```go
package main

import (
	"context"
	"encoding/json"
	"fmt"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
	"io/ioutil"
	"log"
	"net/http"
	"os"
)

func main() {
	srv := getService()

	// スプレッドシートを作成する
	// resp, err := srv.Spreadsheets.Create(&sheets.Spreadsheet{}).Do()
	// if err != nil {
	// 	log.Fatalf("Unable to create a spreadsheet: %v", err)
	// }

	// スプレッドシートを取得する
	spreadsheetId := "1sBicl6K3ok_MsvM8MGWLNEYl06uN4U9IMeFE3dtsZYQ"
	resp, err := srv.Spreadsheets.Get(spreadsheetId).Do()
	if err != nil {
		log.Fatalf("%v", err)
	}
	for _, sheet := range resp.Sheets {
		fmt.Println(sheet.Properties.Title)
	}

	// スプレッドシートに書き込む
	_, err = srv.Spreadsheets.Values.Update(
		spreadsheetId,
		"A1:B2",
		&sheets.ValueRange{
			Values: [][]interface{}{
				{1, 2},
				{3, 4},
			},
		}).ValueInputOption("RAW").Do()
	if err != nil {
		log.Fatalf("%v", err)
	}
	// 数式を書き込むこともできる
	_, err = srv.Spreadsheets.Values.Update(
		spreadsheetId,
		"A3:B3",
		&sheets.ValueRange{
			Values: [][]interface{}{
				{"= SUM(A1:A2)", "= SUM(B1:B2)"},
			},
		}).ValueInputOption("USER_ENTERED").Do()
	if err != nil {
		log.Fatalf("%v", err)
	}
}

func getService() *sheets.Service {
	ctx := context.Background()
	b, err := ioutil.ReadFile("credentials.json")
	if err != nil {
		log.Fatalf("Unable to read client secret file: %v", err)
	}

	// If modifying these scopes, delete your previously saved token.json.
	config, err := google.ConfigFromJSON(b, sheets.SpreadsheetsScope)
	if err != nil {
		log.Fatalf("Unable to parse client secret file to config: %v", err)
	}
	client := getClient(config)

	srv, err := sheets.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	}

	return srv
}

func getClient(config *oauth2.Config) *http.Client {
	// The file token.json stores the user's access and refresh tokens, and is
	// created automatically when the authorization flow completes for the first
	// time.
	tokFile := "token.json"
	tok, err := tokenFromFile(tokFile)
	if err != nil {
		tok = getTokenFromWeb(config)
		saveToken(tokFile, tok)
	}
	return config.Client(context.Background(), tok)
}

func saveToken(path string, token *oauth2.Token) {
	fmt.Printf("Saving credential file to: %s\n", path)
	f, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		log.Fatalf("Unable to cache oauth token: %v", err)
	}
	defer func(f *os.File) {
		err := f.Close()
		if err != nil {
			log.Fatalf("Failed to close file: %v", path)
		}
	}(f)
	err = json.NewEncoder(f).Encode(token)
	if err != nil {
		log.Fatalf("Failed to encode file: %v", path)
	}
}

// Request a token from the web, then returns the retrieved token.
func getTokenFromWeb(config *oauth2.Config) *oauth2.Token {
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Printf("Go to the following link in your browser then type the "+
		"authorization code: \n%v\n", authURL)

	var authCode string
	if _, err := fmt.Scan(&authCode); err != nil {
		log.Fatalf("Unable to read authorization code: %v", err)
	}

	tok, err := config.Exchange(context.TODO(), authCode)
	if err != nil {
		log.Fatalf("Unable to retrieve token from web: %v", err)
	}
	return tok
}

// Retrieves a token from a local file.
func tokenFromFile(file string) (*oauth2.Token, error) {
	f, err := os.Open(file)
	if err != nil {
		return nil, err
	}
	defer func(f *os.File) {
		err := f.Close()
		if err != nil {
			log.Fatalf("Failed to close file: %v", file)
		}
	}(f)
	tok := &oauth2.Token{}
	err = json.NewDecoder(f).Decode(tok)
	return tok, err
}
```

## 大変だったところ

### そもそもコードを動かすためにクレデンシャルの作成が必要だった

クレデンシャルの中でも `OAuth client ID` というのを取得しないと、ユーザが管理しているスプレッドシートを編集したり、新たにスプレッドシートを作成したり、ということができなかった。

参考：

[[C#]Google Sheets Api(OAuth)でシート作成！ | codelikeなブログ](https://codelikes.com/csharp-spreadsheet-api/#toc4)

### Sheets API のガイドには Quickstart しかなく、具体的なコードのイメージがしにくい

認証部分は Quickstart を丸パクでも行けるのですが、実際にスプレッドシートを作成したり値を書き換えたりするには
[こちら](https://pkg.go.dev/google.golang.org/api/sheets/v4)
を読み解く必要があり、これがかなり難しい。

Sheets API のガイドにも例は載っているものの、Go のものが存在していないため、頑張って想像で書くしかないのが大変。

例えば、セル内の値を書き換えるという [こちら](https://developers.google.com/sheets/api/guides/values#writing_to_a_single_range) のページ。

Python の例を見て雰囲気で Go で書いてみるしかない。

```python
values = [
    [
        # Cell values ...
    ],
    # Additional rows ...
]
body = {
    'values': values
}
result = service.spreadsheets().values().update(
    spreadsheetId=spreadsheet_id, range=range_name,
    valueInputOption=value_input_option, body=body).execute()
print('{0} cells updated.'.format(result.get('updatedCells')))
```

これを書き換えると、以下のようになる。具体的な値を入れている。

```go
range_ := "A1:B2"
values := [][]interface{}{
	{1, 2},
	{3, 4},
}
res, err := srv.Spreadsheets.Values.Update(
	spreadsheetId,
	range_,
	&sheets.ValueRange{
		Values: values,
	}).ValueInputOption("RAW").Do()
if err != nil {
	log.Fatalf("%v", err)
}
fmt.Printf("%v cells updated.\n", res.UpdatedCells)
```

### このコードどうやって思いついたの？

手持ちの IDE に強力なコード補完を用意してやることで、res の中に `UpdatedCells` というのがあるんだな～というのは都度学習しながら書くことができる。

JetBrains 社の GoLand は特におすすめ。あるいは（使ったことないけど）VSCode の Go 系のプラグインもよい。

Go が標準で gofmt というフォーマッタを搭載していることから、保存時に gofmt が実行されるように IDE を設定すればさらに便利。

時代は IDE というが、IDE の時代がやってきたというよりは、こういった Go のような言語が IDE 時代をけん引してきたともいえる。
