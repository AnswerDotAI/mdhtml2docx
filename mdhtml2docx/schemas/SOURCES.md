# Schema provenance

The `*.xsd` files except `xml.xsd` are the wordprocessing import closure (12 of 26 files) of the ECMA-376 Part 4 5th edition transitional schema set, from `OfficeOpenXML-XMLSchema-Transitional.zip` inside https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip (sml/pml and their closures omitted: spreadsheets and presentations are out of scope).

`xml.xsd` is the W3C schema for the XML namespace, from https://www.w3.org/2001/xml.xsd.

Local modification, needed because lxml cannot resolve a namespace-only import: in `wml.xsd` and `shared-math.xsd`, the `<xsd:import namespace="http://www.w3.org/XML/1998/namespace"/>` element gained `schemaLocation="xml.xsd"`. No other edits.
