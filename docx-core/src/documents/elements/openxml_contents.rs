use super::*;
use serde::Serialize;

use crate::documents::BuildXML;
use crate::xml_builder::*;

#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct OpenXmlContents {
    pub xml_texts: Vec<String>,
    pub draw_data: Option<Vec<Pic>>,
}

impl Default for OpenXmlContents {
    fn default() -> Self {
        Self::new()
    }
}

impl OpenXmlContents {
    pub fn new() -> Self {
        Self {
            xml_texts: Vec::new(),
            draw_data: None,
        }
    }

    pub fn add_xml_text(mut self, item: &str) -> Self {
        self.xml_texts.push(String::from(item));
        self
    }

    pub fn set_draw_data(mut self, pic: &Pic) -> Self {
        if self.draw_data.is_none() {
            self.draw_data = Some(Vec::new());
        }
        self.draw_data.as_mut().unwrap().push(pic.to_owned());
        self
    }
}

impl BuildXML for OpenXmlContents {
    fn build(&self) -> Vec<u8> {
        let b = XMLBuilder::new().add_xml_texts(&self.xml_texts);
        if self.draw_data.is_some() {
            for item in self.draw_data.as_ref().unwrap() {
                item.build();
            }
        }
        b.build()
    }
}

impl BuildXML for Box<OpenXmlContents> {
    fn build(&self) -> Vec<u8> {
        OpenXmlContents::build(self)
    }
}

#[cfg(test)]
mod tests {

    use super::*;
    #[cfg(test)]
    use pretty_assertions::assert_eq;
    use std::str;

    #[test]
    fn test_openxml() {
        let b = OpenXmlContents::new().build();
        assert_eq!(str::from_utf8(&b).unwrap(), r#""#);
    }

    #[test]
    fn test_openxml_single() {
        let sample_text = r#"<w:p><w:r>(1)</w:r><w:r><w:rPr><w:vertAlign w:val="superscript" /></w:rPr><w:t>4</w:t></w:r><w:r><w:rPr><w:rFonts w:scii="Cambria Math" w:hAnsi="Combria Math" h:hint="eastAsia"></w:rPr><w:t>×</w:t></w:r></w:p>"#;
        let mut x = OpenXmlContents::new();
        x = x.add_xml_text(sample_text);
        let b = x.build();
        assert_eq!(str::from_utf8(&b).unwrap(), sample_text);
    }

    #[test]
    fn test_openxml_multi() {
        let sample_text = [
            r#"<w:p><w:r>(1)</w:r><w:r><w:rPr><w:vertAlign w:val="superscript" /></w:rPr><w:t>4</w:t></w:r><w:r><w:rPr><w:rFonts w:scii="Cambria Math" w:hAnsi="Combria Math" h:hint="eastAsia"></w:rPr><w:t>×</w:t></w:r></w:p>"#,
            r#"<w:p><w:r>(2)</w:r><w:r><w:rPr><w:vertAlign w:val="superscript" /></w:rPr><w:t>5</w:t></w:r><w:r><w:rPr><w:rFonts w:scii="Cambria Math" w:hAnsi="Combria Math" h:hint="eastAsia"></w:rPr><w:t>y</w:t></w:r></w:p>"#,
        ];
        let mut x = OpenXmlContents::new();
        for item in sample_text {
            x = x.add_xml_text(item);
        }
        let b = x.build();
        assert_eq!(str::from_utf8(&b).unwrap(), sample_text.join(""));
    }
}
