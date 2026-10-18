import { describe, expect, it } from 'vitest';
import { strToU8, zipSync } from 'fflate';
import { readSpreadsheetAppearance } from '../electron/spreadsheet-appearance';

describe('original spreadsheet appearance', () => {
  it('reads fonts, fill, alignment, borders and worksheet relationships from real OOXML', () => {
    const zip = zipSync({
      'xl/workbook.xml': strToU8(
        '<workbook xmlns="urn:sheet" xmlns:r="urn:rel"><sheets><sheet name="Budget" r:id="rId7"/></sheets></workbook>',
      ),
      'xl/_rels/workbook.xml.rels': strToU8(
        '<Relationships xmlns="urn:rel"><Relationship Id="rId7" Target="/xl/worksheets/sheet3.xml"/></Relationships>',
      ),
      'xl/styles.xml': strToU8(
        '<styleSheet xmlns="urn:sheet"><fonts><font><name val="Arial"/><sz val="14"/><b/><color rgb="FF123456"/></font></fonts><fills><fill><patternFill patternType="solid"><fgColor rgb="FFABCDEF"/></patternFill></fill></fills><borders><border><bottom style="thin"><color rgb="FF000000"/></bottom></border></borders><cellXfs><xf fontId="0" fillId="0" borderId="0"><alignment horizontal="center" vertical="center" wrapText="1"/></xf></cellXfs></styleSheet>',
      ),
      'xl/worksheets/sheet3.xml': strToU8(
        '<worksheet xmlns="urn:sheet"><sheetData><row r="4"><c r="B4" s="0"><v>123</v></c></row></sheetData></worksheet>',
      ),
    });
    expect(readSpreadsheetAppearance(Buffer.from(zip))).toEqual({
      Budget: {
        '3:1': {
          'font-family': 'Arial',
          'font-size': '14pt',
          'font-weight': 'bold',
          color: '#123456',
          'background-color': '#abcdef',
          'text-align': 'center',
          'vertical-align': 'middle',
          'white-space': 'pre-wrap',
          'border-bottom': '1px solid #000000',
        },
      },
    });
  });

  it('returns no OOXML styles for legacy Excel and CSV without changing their data', () => {
    expect(readSpreadsheetAppearance(Buffer.from('A,B\n1,2'))).toEqual({});
  });
});
