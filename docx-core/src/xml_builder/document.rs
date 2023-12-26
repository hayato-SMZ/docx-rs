use super::XMLBuilder;
use super::XmlEvent;

impl XMLBuilder {
    // i.e. <w:document ... >
    pub(crate) fn open_document(mut self) -> Self {
        self.writer
            .write(
                XmlEvent::start_element("w:document")
                    .attr("xmlns:o", "urn:schemas-microsoft-com:office:office")
                    .attr(
                        "xmlns:r",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    )
                    .attr("xmlns:v", "urn:schemas-microsoft-com:vml")
                    .attr(
                        "xmlns:m",
                        "http://schemas.openxmlformats.org/officeDocument/2006/math",
                    )
                    .attr(
                        "xmlns:w",
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                    )
                    .attr("xmlns:w10", "urn:schemas-microsoft-com:office:word")
                    .attr(
                        "xmlns:wp",
                        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                    )
                    .attr(
                        "xmlns:wpc",
                        "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
                    )
                    .attr(
                        "xmlns:cx",
                        "http://schemas.microsoft.com/office/drawing/2014/chartex",
                    )
                    .attr(
                        "xmlns:cx1",
                        "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
                    )
                    .attr(
                        "xmlns:cx2",
                        "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex",
                    )
                    .attr(
                        "xmlns:cx3",
                        "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex",
                    )
                    .attr(
                        "xmlns:cx4",
                        "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex",
                    )
                    .attr(
                        "xmlns:cx5",
                        "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex",
                    )
                    .attr(
                        "xmlns:cx6",
                        "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex",
                    )
                    .attr(
                        "xmlns:cx7",
                        "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex",
                    )
                    .attr(
                        "xmlns:cx8",
                        "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex",
                    )
                    .attr(
                        "xmlns:wps",
                        "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
                    )
                    .attr(
                        "xmlns:wpi",
                        "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
                    )
                    .attr(
                        "xmlns:wne",
                        "http://schemas.microsoft.com/office/word/2006/wordml",
                    )
                    .attr(
                        "xmlns:wpg",
                        "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
                    )
                    .attr(
                        "xmlns:mc",
                        "http://schemas.openxmlformats.org/markup-compatibility/2006",
                    )
                    .attr(
                        "xmlns:wp14",
                        "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
                    )
                    .attr(
                        "xmlns:w14",
                        "http://schemas.microsoft.com/office/word/2010/wordml",
                    )
                    .attr(
                        "xmlns:w15",
                        "http://schemas.microsoft.com/office/word/2012/wordml",
                    )
                    .attr(
                        "xmlns:aink",
                        "http://schemas.microsoft.com/office/drawing/2016/ink",
                    )
                    .attr(
                        "xmlns:am3d",
                        "http://schemas.microsoft.com/office/drawing/2017/model3d",
                    )
                    .attr(
                        "xmlns:w16cid",
                        "http://schemas.microsoft.com/office/word/2016/wordml/cid",
                    )
                    .attr(
                        "xmlns:w16se",
                        "http://schemas.microsoft.com/office/word/2015/wordml/symex",
                    )
                    .attr("mc:Ignorable", "w14 wp14"),
            )
            .expect("should write to buf");
        self
    }
}

#[cfg(test)]
mod tests {

    use super::*;
    #[cfg(test)]
    use pretty_assertions::assert_eq;
    use std::str;

    #[test]
    fn test_document() {
        let b = XMLBuilder::new();
        let r = b.open_document().close().build();
        assert_eq!(
            str::from_utf8(&r).unwrap(),
            r#"<w:document xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 wp14" />"#
        );
    }
}
