// Package rels parses OOXML relationship XML files (.rels).
//
// It exists to eliminate duplicated parseRelsXML / xmlRelationships code from
// workbook/ and worksheet/, which cannot share the code directly due to the
// import graph.
package rels

import (
	"encoding/xml"
	"fmt"
)

// Relationships is the root element of a .rels XML document.
type Relationships struct {
	Relationships []Relationship `xml:"Relationship"`
}

// Relationship is one entry in a .rels XML document.
type Relationship struct {
	ID     string `xml:"Id,attr"`
	Target string `xml:"Target,attr"`
}

// ParseRelsXML parses the raw bytes of a .rels XML file and returns a map of
// relationship ID → target string.
func ParseRelsXML(data []byte) (map[string]string, error) {
	var r Relationships
	if err := xml.Unmarshal(data, &r); err != nil {
		return nil, fmt.Errorf("parse rels XML: %w", err)
	}
	m := make(map[string]string, len(r.Relationships))
	for _, rel := range r.Relationships {
		m[rel.ID] = rel.Target
	}
	return m, nil
}
