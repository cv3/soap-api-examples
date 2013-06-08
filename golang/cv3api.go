/*
Blake Ellis <blake@commercev3.com>

Usage:
      api := cv3api.NewApi()
      api.SetCredentials("store-name","user-name","password")
      api.GetProductSingle(ctx.num)
      data := api.Execute()
      fmt.Printf(string(data))

*/

package cv3api

import (
	"bytes"
	"encoding/base64"
	"encoding/xml"
	"fmt"
	"io"
	"io/ioutil"
	"net/http"
	"strings"
)

const (
	cv3_endpoint = "https://service.commercev3.com/"
	soapEnvelope = "<SOAP-ENV:Envelope xmlns:SOAP-ENV=\"http://www.w3.org/2001/12/soap-envelope\" SOAP-ENV:encodingStyle=\"http://www.w3.org/2001/12/soap-encoding\">\n  <SOAP-ENV:Body>\n<m:CV3Data xmlns:m=\"http://soapinterop.org/\" SOAP-ENV:encodingStyle=\"http://schemas.xmlsoap.org/soap/encoding/\">\n<data xsi:type=\"xsd:string\">%v</data>\n</m:CV3Data>\n</SOAP-ENV:Body>\n</SOAP-ENV:Envelope>\n\n"
)

type Credentials struct {
	XMLName   xml.Name `xml:"authenticate"`
	User      string   `xml:"user"`
	Password  string   `xml:"pass"`
	ServiceID string   `xml:"serviceID"`
}

type RequestBody struct {
	XMLName  xml.Name `xml:"request"`
	Auth     Credentials
	Requests []Request `xml:"requests"`
}

type Request struct {
	Request string `xml:",innerxml"`
}

type wrapper struct {
	XMLName xml.Name `xml:"CV3Data"`
	Wrap    RequestBody
}

type response struct {
	XMLName xml.Name `xml:"Envelope"`
	Data    string   `xml:"Body>CV3DataResponse>return"`
}

type nopCloser struct {
	io.Reader
}

func toBase64(data string) string {
	var buf bytes.Buffer
	encoder := base64.NewEncoder(base64.StdEncoding, &buf)
	encoder.Write([]byte(data))
	encoder.Close()
	return buf.String()
}

type Api struct {
	user      string
	pass      string
	serviceID string
	request   string
}

func NewApi() *Api {
	api := new(Api)
	return api
}

func (self *Api) SetCredentials(username, password, serviceID string) {
	self.user = username
	self.pass = password
	self.serviceID = serviceID
}

func (self *Api) GetCustomerGroups() {
	self.request = "<reqCustomerInformation members_only=\"false\"/>"
}

func (self *Api) GetProductSingle(o string) {
	self.request = "<reqProducts><reqProductSingle>" + o + "</reqProductSingle></reqProducts>"
}

func (self *Api) Execute() (n []byte) {
  var pre_n []byte 
	w := Credentials{User: self.user, Password: self.pass, ServiceID: self.serviceID}
	x := Request{Request: self.request}
	t := RequestBody{Auth: w, Requests: []Request{x}}
	v := &wrapper{Wrap: t}
	xmlbytes, err := xml.MarshalIndent(v, "  ", "    ")
	xmlstring := string(xmlbytes)
	xmlstring = strings.Replace(xmlstring, "<CV3Data>", "<CV3Data version=\"2.0\">", -1)
	//fmt.Printf(xmlstring)
	encodedString := toBase64(xmlstring)
	xmlstring = xml.Header + fmt.Sprintf(soapEnvelope, encodedString)
	if err == nil {
		client := &http.Client{}
		body := nopCloser{bytes.NewBufferString(xmlstring)}
		if err == nil {
			req, err := http.NewRequest("POST", cv3_endpoint, body)
			if err == nil {
				req.Header.Add("Accept", "text/xml")
				req.Header.Add("Content-Type", "text/xml; charset=utf-8")
				req.Header.Add("SOAPAction", "http://service.commercev3.com/index.php/CV3Data")
				req.ContentLength = int64(len(string(xmlstring)))
				//preq, _ := ioutil.ReadAll(req.Body)
				resp, err := client.Do(req)
				if err != nil {
					fmt.Printf("Request error: %v", err)
					return
				}
				res, err := ioutil.ReadAll(resp.Body)
				resp.Body.Close()
				if err != nil {
					fmt.Printf("Read Response Error: %v", err)
					return
				}
				y := response{}
				err = xml.Unmarshal([]byte(res), &y)
				if err != nil {
					fmt.Printf("Unmarshal error: %v", err)
					return
				}
        pre_n, err = base64.StdEncoding.DecodeString(y.Data)
				if err != nil {
					fmt.Printf("Decoding error: %v", err)
					return
				}
			}
		}
	}
  err = xml.Unmarshal(pre_n,&n)
  if err != nil {
    fmt.Printf("second unmarshal unsuccessful: %v", err)
    return
  }
  fmt.Printf(string(n))
	return
}
