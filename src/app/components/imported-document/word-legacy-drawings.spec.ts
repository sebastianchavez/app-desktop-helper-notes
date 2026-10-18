import { strToU8, strFromU8, unzipSync, zipSync } from 'fflate';
import { prepareWordForEditing, prepareWordForPrint } from './word-legacy-drawings';

/** Crea un DOCX VML mínimo; recibe estilos opcionales y retorna un ZIP con medios de prueba. */
function legacyFixture(extraStyle = ''): Uint8Array {
  return zipSync({
    'word/document.xml': strToU8(
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:pict><v:rect id="Text" filled="f" stroked="f" style="position:absolute;margin-left:240pt;margin-top:33pt;width:360pt;height:40pt;z-index:2;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;${extraStyle}"><v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:r><w:t>Editable text</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></w:r><w:r><w:pict><v:rect id="Image" filled="f" stroked="f" style="position:absolute;margin-left:0pt;margin-top:0pt;width:612pt;height:864pt;z-index:-1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page"><v:imagedata r:id="rId9"/></v:rect></w:pict></w:r></w:p></w:body></w:document>`,
    ),
    'word/media/image1.png': new Uint8Array([1, 2, 3]),
    'word/_rels/document.xml.rels': strToU8('original relationships'),
  });
}

describe('legacy Word drawing compatibility', () => {
  it('turns simple page-anchored VML into editable DrawingML without changing media or relationships', () => {
    const original = legacyFixture();
    const result = unzipSync(prepareWordForEditing(original));
    const xml = new DOMParser().parseFromString(
      strFromU8(result['word/document.xml']),
      'application/xml',
    );
    expect(xml.querySelector('parsererror')).toBeNull();
    expect(xml.getElementsByTagName('w:pict')).toHaveLength(0);
    expect(xml.getElementsByTagName('wps:wsp')).toHaveLength(1);
    expect(xml.getElementsByTagName('pic:pic')).toHaveLength(1);
    expect(xml.getElementsByTagName('wp:posOffset')[0].textContent).toBe('3048000');
    expect(xml.getElementsByTagName('wp:anchor')[1].getAttribute('behindDoc')).toBe('1');
    expect(xml.getElementsByTagName('a:blip')[0].getAttribute('r:embed')).toBe('rId9');
    expect(xml.getElementsByTagName('w:t')[0].textContent).toBe('Editable text');
    expect(result['word/media/image1.png']).toEqual(new Uint8Array([1, 2, 3]));
    expect(strFromU8(result['word/_rels/document.xml.rels'])).toBe('original relationships');
  });

  it('preserves unsupported rotated text boxes instead of flattening or dropping them', () => {
    const result = unzipSync(prepareWordForEditing(legacyFixture('rotation:30')));
    const xml = new DOMParser().parseFromString(
      strFromU8(result['word/document.xml']),
      'application/xml',
    );
    expect(xml.getElementsByTagName('v:textbox')).toHaveLength(1);
    expect(xml.getElementsByTagName('pic:pic')).toHaveLength(1);
  });

  it('does not rewrite already modern documents', () => {
    const modern = prepareWordForEditing(legacyFixture());
    expect(prepareWordForEditing(modern)).toBe(modern);
  });

  it('retains edited floating text and its position in the separate print projection', () => {
    const edited = prepareWordForEditing(legacyFixture());
    const snapshot = unzipSync(prepareWordForPrint(edited));
    const xml = new DOMParser().parseFromString(
      strFromU8(snapshot['word/document.xml']),
      'application/xml',
    );
    expect(xml.getElementsByTagName('v:textbox')).toHaveLength(1);
    expect(xml.getElementsByTagName('w:t')[0].textContent).toBe('Editable text');
    expect(xml.getElementsByTagName('v:rect')[0].getAttribute('style')).toContain(
      'margin-left:240pt',
    );
    expect(xml.getElementsByTagName('pic:pic')).toHaveLength(1);
  });
});
