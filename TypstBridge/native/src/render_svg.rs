use typst::layout::PagedDocument;

pub fn render(document: &PagedDocument) -> Vec<(u32, Vec<u8>)> {
    document
        .pages
        .iter()
        .enumerate()
        .map(|(page_index, page)| (page_index as u32, typst_svg::svg(page).into_bytes()))
        .collect()
}
