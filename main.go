package main

import (
	"log"
	"reflect"

	"github.com/Pallinder/go-randomdata"
	"github.com/unidoc/unioffice/spreadsheet"
)

const amountOfPeople = 2 * 1000

func main() {
	generateSheet()
}

func generateSheet() {
	type PersonData struct {
		FirstName   string `header:"-"`
		LastName    string `header:"Lastname"`
		NickName    string `header:"Nickname"`
		Email       string `header:"Email"`
		City        string `header:"City"`
		Street      string `header:"Street"`
		PhoneNumber string `header:"Phonenumber"`
	}

	peoples := make([]PersonData, amountOfPeople) // knowing the size makes it much faster
	for i := 0; i < amountOfPeople; i++ {
		person := PersonData{
			FirstName:   randomdata.FirstName(randomdata.RandomGender),
			LastName:    randomdata.LastName(),
			NickName:    randomdata.SillyName(),
			Email:       randomdata.Email(),
			City:        randomdata.City(),
			Street:      randomdata.Street(),
			PhoneNumber: randomdata.PhoneNumber(),
		}
		peoples = append(peoples, person)
	}

	ss := spreadsheet.New()
	// add a single sheet
	sheet := ss.AddSheet()

	// header
	row := sheet.AddRow()
	p := PersonData{}
	s := reflect.ValueOf(&p).Elem()
	typeOfT := s.Type()
	for i := 0; i < s.NumField(); i++ {
		// Use tag or fallback to struct name
		h := typeOfT.Field(i).Tag.Get("header")
		if h == "-" {
			continue
		}
		if h == "" {
			h = typeOfT.Field(i).Name
		}
		cell := row.AddCell()
		cell.SetString(h)
	}

	// rows
	for _, p := range peoples {
		row := sheet.AddRow()
		// and cells

		// reflect is maybe a bit slow
		s := reflect.ValueOf(&p).Elem()
		//typeOfT := s.Type()

		for i := 0; i < s.NumField(); i++ {
			h := typeOfT.Field(i).Tag.Get("header")
			if h == "-" {
				continue
			}
			f := s.Field(i)
			// fmt.Printf("%d: %s %s = %v\n", i, typeOfT.Field(i).Name, f.Type(), f.Interface())
			cell := row.AddCell()
			cell.SetString(f.Interface().(string))
		}

	}

	if err := ss.Validate(); err != nil {
		log.Fatalf("error validating sheet: %s", err)
	}

	ss.SaveToFile("simple.xlsx")
}
