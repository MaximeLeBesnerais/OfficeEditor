use typst_layout::PagedDocument;
use typst_utils::Scalar;

pub fn render(document: &PagedDocument, ppi: f64) -> Result<Vec<(u32, Vec<u8>)>, String> {
    if !ppi.is_finite() || ppi <= 0.0 {
        return Err("ppi must be positive and finite for PNG output".to_owned());
    }

    let options = typst_render::RenderOptions {
        pixel_per_pt: Scalar::new(ppi / 72.0),
        ..Default::default()
    };
    document
        .pages()
        .iter()
        .enumerate()
        .map(|(page_index, page)| {
            let pixmap = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                typst_render::render(page, &options)
            }))
            .map_err(|_| format!("Failed to render page {} as PNG", page_index + 1))?;
            let data = pixmap.encode_png().map_err(|error| {
                format!("Failed to encode page {} as PNG: {error}", page_index + 1)
            })?;
            Ok((page_index as u32, data))
        })
        .collect()
}
