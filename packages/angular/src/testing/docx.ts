// DOCX fixtures for the Angular specs. Not part of the package build.
import { strToU8, zipSync } from 'fflate';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

/** Minimal DOCX bytes with `body` as the content of `w:body`. */
export function docx(body: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '</Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

/**
 * A date form field with the value `01/02/2030`, then ` tail`. With `protectedForm`, the
 * settings allow only form edits. A copy of the core fixture, so the specs do not import core
 * source files.
 */
export function formFieldDocx(protectedForm = false): Uint8Array {
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${rel}/settings" Target="settings.xml"/></Relationships>`
    ),
    'word/settings.xml': strToU8(
      `<w:settings xmlns:w="${W}">${protectedForm ? '<w:documentProtection w:edit="forms" w:enforcement="1"/>' : ''}</w:settings>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Date"/><w:textInput><w:type w:val="date"/><w:default w:val="01/02/2030"/><w:format w:val="MM/dd/yyyy"/></w:textInput></w:ffData></w:fldChar></w:r><w:r><w:instrText> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>01/02/2030</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> tail</w:t></w:r></w:p></w:body></w:document>`
    ),
  });
}

/** One run that keeps its spaces. */
export const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** One paragraph with one run. */
export const paragraph = (text: string): string => `<w:p>${run(text)}</w:p>`;
