// Pane UI localization (WP3.4). Pure string table; the DOM application and the
// dynamic status/button strings live in taskpane.js. Manifest ribbon / AppSource
// strings are localized separately via <Override Locale> in manifest.xml.
//
// To add a language: add a table with the SAME keys as `en` (the key-parity test
// enforces this) and a `<Override Locale>` set in manifest.xml.

export const STRINGS = {
  en: {
    latexLabel: 'LaTeX source', previewLabel: 'Preview',
    fontLabel: 'Font', sizeLabel: 'Size', colorLabel: 'Color',
    symbols: 'Symbols', macros: 'Macros',
    inline: 'Inline', display: 'Display',
    sizeSelection: 'Match selection', sizeBody: 'Match body', sizeFixed: 'Fixed',
    macrosHint: '\\newcommand definitions applied to every equation in this document.',
    alignLabel: 'Display', alignLeft: 'Left (indented, TeX fleqn)', alignCenter: 'Centred',
    numberingLabel: 'Numbering',
    numberStyleInline: 'Inline (after equation)',
    numberStyleTable: 'Flush right (table)',
    numberStyleField: 'Word field (desktop)',
    numberThis: 'Number this equation',
    renumberAll: 'Renumber all',
    newEq: 'New equation', statusNew: 'New equation',
    insert: 'Insert equation', update: 'Update equation',
    examplesLabel: 'Examples…', settingsLabel: 'Settings',
    swatchesLabel: 'Swatches', colorSetQualitative: 'Qualitative cycle', colorSetPresentation: 'Presentation',
    colorSetSemantic: 'Semantic', colorSetPastel: 'Pastel triplets', colorSetShades: 'Shade families',
    colorSetGrays: 'Greys',
    testPage: 'Test page (fills the document)…', testPageGroup: 'Testing',
    testPaper: 'Test paper (RevTeX-style, two columns)…',
    testPaperLayoutFailed: 'The two-column layout could not be applied on this host.',
    testPageProgress: 'Test page: equation {n} of {total}…',
    testPageDone: 'Test page inserted: {ok} equation(s) OK, {failed} failed.',
    testPageUnavailable: 'This host cannot insert equation pictures, so the test page cannot be generated here.',
    equationsLabel: 'Equations in this document',
    eqListEmpty: 'No equations yet', eqListError: 'Could not read the document’s equations.',
    eqNotFound: 'That equation is no longer in the document.',
    previewCtxBefore: 'where ', previewCtxAfter: ' holds for all x.',
    editing: 'Editing equation {id}…',
    migrateOffer: '{n} numbered equation(s) still use a different numbering style.',
    migrateAction: 'Convert to “{style}”',
    migrateDone: 'Converted {n} equation(s) ({failed} failed) and renumbered in document order.',
    converting: 'Converting equations…',
  },
  de: {
    latexLabel: 'LaTeX-Quelltext', previewLabel: 'Vorschau',
    fontLabel: 'Schriftart', sizeLabel: 'Größe', colorLabel: 'Farbe',
    symbols: 'Symbole', macros: 'Makros',
    inline: 'Inline', display: 'Abgesetzt',
    sizeSelection: 'An Auswahl', sizeBody: 'An Fließtext', sizeFixed: 'Fest',
    macrosHint: '\\newcommand-Definitionen, angewendet auf jede Gleichung in diesem Dokument.',
    alignLabel: 'Abgesetzt', alignLeft: 'Linksbündig (eingerückt, TeX fleqn)', alignCenter: 'Zentriert',
    numberingLabel: 'Nummerierung',
    numberStyleInline: 'Inline (nach Gleichung)',
    numberStyleTable: 'Rechtsbündig (Tabelle)',
    numberStyleField: 'Word-Feld (Desktop)',
    numberThis: 'Diese Gleichung nummerieren',
    renumberAll: 'Alle neu nummerieren',
    newEq: 'Neue Gleichung', statusNew: 'Neue Gleichung',
    insert: 'Gleichung einfügen', update: 'Gleichung aktualisieren',
    examplesLabel: 'Beispiele…', settingsLabel: 'Einstellungen',
    swatchesLabel: 'Farbfelder', colorSetQualitative: 'Qualitativer Zyklus', colorSetPresentation: 'Präsentation',
    colorSetSemantic: 'Semantisch', colorSetPastel: 'Pastell-Tripel', colorSetShades: 'Abstufungen',
    colorSetGrays: 'Grautöne',
    testPage: 'Testseite (füllt das Dokument)…', testPageGroup: 'Test',
    testPaper: 'Test-Paper (RevTeX-Stil, zweispaltig)…',
    testPaperLayoutFailed: 'Das zweispaltige Layout konnte auf diesem Host nicht angewendet werden.',
    testPageProgress: 'Testseite: Gleichung {n} von {total}…',
    testPageDone: 'Testseite eingefügt: {ok} Gleichung(en) OK, {failed} fehlgeschlagen.',
    testPageUnavailable: 'Dieser Host kann keine Gleichungsbilder einfügen; die Testseite kann hier nicht erzeugt werden.',
    equationsLabel: 'Gleichungen in diesem Dokument',
    eqListEmpty: 'Noch keine Gleichungen', eqListError: 'Gleichungen konnten nicht gelesen werden.',
    eqNotFound: 'Diese Gleichung ist nicht mehr im Dokument.',
    previewCtxBefore: 'wobei ', previewCtxAfter: ' für alle x gilt.',
    editing: 'Bearbeite Gleichung {id}…',
    migrateOffer: '{n} nummerierte Gleichung(en) verwenden noch einen anderen Nummerierungsstil.',
    migrateAction: 'In „{style}“ umwandeln',
    migrateDone: '{n} Gleichung(en) umgewandelt ({failed} fehlgeschlagen) und in Dokumentreihenfolge neu nummeriert.',
    converting: 'Gleichungen werden umgewandelt…',
  },
};

// Map a BCP-47 tag ('de-DE', 'de', 'en-US') to a supported table key by primary
// subtag; unknown languages fall back to English.
export function pickLocale(lang) {
  const primary = String(lang || '').toLowerCase().split('-')[0];
  return STRINGS[primary] ? primary : 'en';
}

export function strings(lang) {
  return STRINGS[pickLocale(lang)];
}
