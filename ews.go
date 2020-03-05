// 26 august 2016
package ews

import (
	"bytes"
	"crypto/tls"
	"errors"
	"net/http"
	"regexp"
	"strings"
	"time"

	httpntlm "github.com/vadimi/go-http-ntlm"
)

// https://msdn.microsoft.com/en-us/library/office/dd877045(v=exchg.140).aspx
// https://arvinddangra.wordpress.com/2011/09/29/send-email-using-exchange-smtp-and-ews-exchange-web-service/
// https://msdn.microsoft.com/en-us/library/office/dn789003(v=exchg.150).aspx

var soapheader = `<?xml version="1.0" encoding="utf-8" ?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2007_SP1" />
  </soap:Header>
  <soap:Body>
`

var (
	EWSAddr  string
	Username string // mail or domin\account format
	Password string
)

// SendMail just send mail :)
func SendMail(to []string, cc []string, topic string, content string) (*http.Response, error) {
	// check username format
	b, err := BuildTextEmail(Username, to, cc, topic, []byte(content))
	if err != nil {
		return nil, err
	}

	return Issue(EWSAddr, Username, Password, b)
}

func Issue(ewsAddr string, username string, password string, body []byte) (*http.Response, error) {

	b2 := []byte(soapheader)
	b2 = append(b2, body...)
	b2 = append(b2, "\n  </soap:Body>\n</soap:Envelope>"...)
	req, err := http.NewRequest("POST", ewsAddr, bytes.NewReader(b2))
	if err != nil {
		return nil, err
	}

	var client *http.Client
	re := regexp.MustCompile("^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$")
	isMail := re.MatchString(Username)

	if !isMail {
		// use domain
		l := strings.SplitN(Username, "\\", 2)
		if len(l) < 2 {
			return nil, errors.New("Wrong format of username, not email or format with domain\\account")
		}

		domain := l[0]
		account := l[1]
		client = &http.Client{
			Transport: &httpntlm.NtlmTransport{
				Domain:          domain,
				User:            account,
				Password:        password,
				TLSClientConfig: &tls.Config{InsecureSkipVerify: true},
			},
		}
	} else {
		//for office365, no ntlm, emial as username
		req.SetBasicAuth(username, password)
		client = &http.Client{
			Transport: &http.Transport{TLSClientConfig: &tls.Config{InsecureSkipVerify: true}},
		}
	}

	req.Header.Set("Content-Type", "text/xml")
	client.Timeout = 10 * time.Second
	client.CheckRedirect = func(req *http.Request, via []*http.Request) error { return http.ErrUseLastResponse }
	return client.Do(req)
}
