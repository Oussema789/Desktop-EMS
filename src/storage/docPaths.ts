import {
  AtlasDatabaseState,
  DOCUMENT_CATEGORY_LABELS,
  DocumentAttachmentMeta,
  Expert,
} from '../types/database';

export const GEREP_DOCS_FOLDER = 'GEREP-EMS-Docs';

export function sanitizeSegment(value: string): string {
  const cleaned = value
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_')
    .replace(/^_+|_+$/g, '');
  return cleaned || 'Element';
}

export function buildAttachmentRelativePath(
  expert: Expert,
  attachment: DocumentAttachmentMeta
): string {
  const folder = `${sanitizeSegment(expert.lastName)}_${sanitizeSegment(expert.firstName)}_${sanitizeSegment(expert.code)}`;
  const categoryFolder = sanitizeSegment(DOCUMENT_CATEGORY_LABELS[attachment.category]);
  const cleanFile = sanitizeSegment(attachment.fileName);
  return `${GEREP_DOCS_FOLDER}/Experts/${folder}/${categoryFolder}/${cleanFile}`;
}

export function buildEntityAttachmentRelativePath(
  rootFolder: 'Sous-traitants' | 'Partenaires',
  code: string,
  name: string,
  attachment: DocumentAttachmentMeta
): string {
  const folder = `${sanitizeSegment(name)}_${sanitizeSegment(code)}`;
  const cleanFile = sanitizeSegment(attachment.fileName);
  return `${GEREP_DOCS_FOLDER}/${rootFolder}/${folder}/${cleanFile}`;
}

export interface OrganizedAttachmentRecord {
  id: string;
  relativePath: string;
  mimeType: string;
  fileName: string;
}

function makeUniquePath(basePath: string, usedPaths: Set<string>): string {
  if (!usedPaths.has(basePath)) {
    usedPaths.add(basePath);
    return basePath;
  }
  const dotIdx = basePath.lastIndexOf('.');
  const slashIdx = basePath.lastIndexOf('/');
  const hasExt = dotIdx > slashIdx;
  const stem = hasExt ? basePath.slice(0, dotIdx) : basePath;
  const ext = hasExt ? basePath.slice(dotIdx) : '';

  let counter = 2;
  let candidate = `${stem}_${counter}${ext}`;
  while (usedPaths.has(candidate)) {
    counter += 1;
    candidate = `${stem}_${counter}${ext}`;
  }
  usedPaths.add(candidate);
  return candidate;
}

export function buildOrganizedAttachmentRecords(
  state: AtlasDatabaseState
): OrganizedAttachmentRecord[] {
  const records: OrganizedAttachmentRecord[] = [];
  const usedPaths = new Set<string>();

  for (const expert of state.experts) {
    for (const att of expert.attachments) {
      const relativePath = makeUniquePath(
        buildAttachmentRelativePath(expert, att),
        usedPaths
      );
      records.push({
        id: att.id,
        relativePath,
        mimeType: att.mimeType,
        fileName: att.fileName,
      });
    }
  }

  for (const sub of state.subcontractors) {
    for (const att of sub.adminDocuments ?? []) {
      const relativePath = makeUniquePath(
        buildEntityAttachmentRelativePath('Sous-traitants', sub.code, sub.name, att),
        usedPaths
      );
      records.push({
        id: att.id,
        relativePath,
        mimeType: att.mimeType,
        fileName: att.fileName,
      });
    }
  }

  for (const prt of state.partners) {
    for (const att of prt.adminDocuments ?? []) {
      const relativePath = makeUniquePath(
        buildEntityAttachmentRelativePath('Partenaires', prt.code, prt.name, att),
        usedPaths
      );
      records.push({
        id: att.id,
        relativePath,
        mimeType: att.mimeType,
        fileName: att.fileName,
      });
    }
  }

  return records;
}
