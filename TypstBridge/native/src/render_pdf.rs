use typst::layout::PagedDocument;

pub fn render(document: &PagedDocument) -> typst::diag::SourceResult<Vec<u8>> {
    typst_pdf::pdf(document, &typst_pdf::PdfOptions::default())
}
