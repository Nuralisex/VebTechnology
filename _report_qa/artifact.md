# Execution contract for six student reports

## Reference and integrity
- Reference DOCX: `C:\Users\Нурали\OneDrive\Desktop\титул.docx`
- SHA-256 at inspection: `6F7D42CE48ED72332C15342193284DB32C388900F460D98065810508B2C832B6`
- Inspection: read-only package extraction with bundled Python and `python-docx`; section/style audits completed.
- Page count/render: unresolved. The packaged `render_docx.py` could not find bundled LibreOffice (`soffice.exe`); no page PNGs were produced.
- One section, portrait A4 (8.27 × 11.69 in); margins L 1.18 in, R 0.59 in, T/B 0.79 in; header/footer distances 0.49 in.
- Source typography: Times New Roman 14 pt (28 half-points), black, centered title block; paragraphs mostly 12 pt after, 2.0 line spacing. Body page defaults use the sample's 14 pt face, justified paragraphs, 1.5 line spacing.

## Title-page structure and slots
Keep the source first page as the template. It contains the ministry and republic at the top; university; school; course; report label, lab number and topic; author name and group; reviewer; city and year. Preserve the original page furniture, paragraph positions, blank spacer paragraphs, and source formatting. Only edit the report number and topic in the title block. The sample contains these user-supplied values: Musatayev Nurali Kairatuly, group 24-ИС-1, reviewer Egamberdiev Eldor Ulugbekovich, Ust-Kamenogorsk, 2026.

## Report-body plan
Append a page break after the retained cover and write a concise Russian report with: objective, tools/technologies, implementation summary, result and conclusion. Keep all report text Times New Roman 14 pt. Use a clear section heading hierarchy, 1.5 line spacing, justified body paragraphs, 1.25 cm first-line indent, A4 portrait. Reports correspond to the six project files `lab1.html`–`lab6.html`; accurately describe implemented features and do not claim external validation or test results that were not run.

## Package preservation and fidelity gate
The retained template is read-only and must remain byte-for-byte unchanged. Create each final DOCX as a copy of the source, preserving the original cover page as much as the object model allows. The sole content changes on that page are the lab number and subject. Verify the reference hash after authoring. Render every final document and inspect all pages before delivery; the current environment has not resolved the rendering dependency, so delivery remains gated on a successful render and visual review.
