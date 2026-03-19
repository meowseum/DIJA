use quick_xml::events::{BytesText, Event};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use regex::Regex;
use std::collections::HashMap;
use std::io::{Cursor, Read, Write};
use std::path::Path;
use std::sync::LazyLock;
use zip::read::ZipArchive;
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

static PLACEHOLDER_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"\{\{\s*(\w+)\s*\}\}").unwrap());

/// Render a .docx template by replacing `{{ VAR }}` placeholders with values.
pub fn render_docx_template(
    template: &Path,
    output: &Path,
    context: &HashMap<String, String>,
) -> Result<(), String> {
    let file = std::fs::File::open(template).map_err(|e| format!("Cannot open template: {}", e))?;
    let mut archive = ZipArchive::new(file).map_err(|e| format!("Invalid docx: {}", e))?;

    // Read document.xml
    let doc_xml = {
        let mut entry = archive
            .by_name("word/document.xml")
            .map_err(|e| format!("No word/document.xml: {}", e))?;
        let mut buf = String::new();
        entry.read_to_string(&mut buf).map_err(|e| format!("Read error: {}", e))?;
        buf
    };

    // Process XML: concatenate text within each paragraph, apply replacements,
    // then redistribute back. This handles the "split run" problem.
    let replaced_xml = replace_placeholders_in_xml(&doc_xml, context)?;

    // Write output docx
    let out_file = std::fs::File::create(output).map_err(|e| format!("Cannot create output: {}", e))?;
    let mut zip_writer = ZipWriter::new(out_file);

    for i in 0..archive.len() {
        let mut entry = archive.by_index(i).map_err(|e| format!("Zip entry error: {}", e))?;
        let name = entry.name().to_string();
        let options = SimpleFileOptions::default()
            .compression_method(entry.compression());

        if name == "word/document.xml" {
            zip_writer.start_file(&name, options).map_err(|e| format!("Zip write error: {}", e))?;
            zip_writer.write_all(replaced_xml.as_bytes()).map_err(|e| format!("Write error: {}", e))?;
        } else {
            let mut data = Vec::new();
            entry.read_to_end(&mut data).map_err(|e| format!("Read error: {}", e))?;
            zip_writer.start_file(&name, options).map_err(|e| format!("Zip write error: {}", e))?;
            zip_writer.write_all(&data).map_err(|e| format!("Write error: {}", e))?;
        }
    }

    zip_writer.finish().map_err(|e| format!("Zip finish error: {}", e))?;
    Ok(())
}

/// Extract all text from a .docx file, joining paragraphs with newlines.
pub fn extract_docx_text(path: &Path) -> Result<String, String> {
    let file = std::fs::File::open(path).map_err(|e| format!("Cannot open docx: {}", e))?;
    let mut archive = ZipArchive::new(file).map_err(|e| format!("Invalid docx: {}", e))?;

    let doc_xml = {
        let mut entry = archive
            .by_name("word/document.xml")
            .map_err(|e| format!("No word/document.xml: {}", e))?;
        let mut buf = String::new();
        entry.read_to_string(&mut buf).map_err(|e| format!("Read error: {}", e))?;
        buf
    };

    let mut paragraphs: Vec<String> = Vec::new();
    let mut reader = Reader::from_str(&doc_xml);
    reader.config_mut().trim_text(false);

    let mut in_paragraph = false;
    let mut in_text = false;
    let mut para_text = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == b"p" {
                    in_paragraph = true;
                    para_text.clear();
                } else if local.as_ref() == b"t" && in_paragraph {
                    in_text = true;
                }
            }
            Ok(Event::Text(ref e)) if in_text => {
                if let Ok(text) = e.unescape() {
                    para_text.push_str(&text);
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == b"t" {
                    in_text = false;
                } else if local.as_ref() == b"p" {
                    in_paragraph = false;
                    paragraphs.push(para_text.clone());
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("XML parse error: {}", e)),
            _ => {}
        }
    }

    Ok(paragraphs.join("\n"))
}

/// Replace `{{ VAR }}` placeholders in the XML, handling split runs.
///
/// Strategy: for each `<w:p>` paragraph, collect all `<w:t>` text, concatenate,
/// apply regex replacements, then put all replaced text into the first `<w:t>`
/// and empty subsequent ones.
fn replace_placeholders_in_xml(
    xml: &str,
    context: &HashMap<String, String>,
) -> Result<String, String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::new()));

    // We process in two passes logically, but use a single-pass approach:
    // Buffer events within each <w:p>, then flush with replacements applied.

    let mut in_paragraph = false;
    let mut para_events: Vec<Event<'static>> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Eof) => break,
            Ok(event) => {
                let is_para_start = matches!(&event, Event::Start(e) if e.local_name().as_ref() == b"p"
                    && e.name().as_ref().starts_with(b"w:p"));
                let is_para_end = matches!(&event, Event::End(e) if e.local_name().as_ref() == b"p"
                    && e.name().as_ref().starts_with(b"w:p"));

                if is_para_start {
                    in_paragraph = true;
                    para_events.clear();
                    para_events.push(event.into_owned());
                } else if is_para_end && in_paragraph {
                    para_events.push(event.into_owned());
                    // Process this paragraph
                    flush_paragraph(&para_events, &mut writer, context)?;
                    in_paragraph = false;
                    para_events.clear();
                } else if in_paragraph {
                    para_events.push(event.into_owned());
                } else {
                    writer.write_event(event).map_err(|e| format!("Write error: {}", e))?;
                }
            }
            Err(e) => return Err(format!("XML parse error: {}", e)),
        }
    }

    let result = writer.into_inner().into_inner();
    String::from_utf8(result).map_err(|e| format!("UTF-8 error: {}", e))
}

fn flush_paragraph(
    events: &[Event<'static>],
    writer: &mut Writer<Cursor<Vec<u8>>>,
    context: &HashMap<String, String>,
) -> Result<(), String> {
    // Collect all text from <w:t> elements
    let mut texts: Vec<String> = Vec::new();
    let mut text_indices: Vec<usize> = Vec::new();

    for (i, event) in events.iter().enumerate() {
        if let Event::Text(t) = event {
            // Check if previous event was a <w:t> start
            if i > 0 {
                let prev = &events[i - 1];
                let is_wt_start = matches!(prev, Event::Start(e) if e.local_name().as_ref() == b"t"
                    && e.name().as_ref().starts_with(b"w:t"));
                if is_wt_start {
                    if let Ok(unescaped) = t.unescape() {
                        texts.push(unescaped.to_string());
                        text_indices.push(i);
                    }
                }
            }
        }
    }

    if texts.is_empty() || !PLACEHOLDER_RE.is_match(&texts.join("")) {
        // No placeholders — write events as-is
        for event in events {
            writer.write_event(event.clone()).map_err(|e| format!("Write error: {}", e))?;
        }
        return Ok(());
    }

    // Concatenate, replace, then put all text in first <w:t>, empty the rest
    let concatenated = texts.join("");
    let replaced = PLACEHOLDER_RE.replace_all(&concatenated, |caps: &regex::Captures| {
        let key = &caps[1];
        context.get(key).cloned().unwrap_or_default()
    }).to_string();

    let mut text_count = 0;
    for (i, event) in events.iter().enumerate() {
        if text_indices.contains(&i) {
            if text_count == 0 {
                // First <w:t> text node: write the full replaced text
                let escaped = BytesText::new(&replaced);
                writer.write_event(Event::Text(escaped)).map_err(|e| format!("Write error: {}", e))?;
            } else {
                // Subsequent <w:t> text nodes: write empty
                let escaped = BytesText::new("");
                writer.write_event(Event::Text(escaped)).map_err(|e| format!("Write error: {}", e))?;
            }
            text_count += 1;
        } else {
            writer.write_event(event.clone()).map_err(|e| format!("Write error: {}", e))?;
        }
    }

    Ok(())
}
