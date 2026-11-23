import {
  AlertCircle,
  Check,
  FileSpreadsheet,
  FileText,
  FolderOpen,
  Sparkles,
  Trash2,
  Upload,
  X,
} from 'lucide-react';
import { ChangeEvent, DragEvent, useMemo, useState } from 'react';
import {
  autoMapExpertColumns,
  autoMapOrgColumns,
  buildExpertsFromTable,
  buildPartnersFromTable,
  buildSubcontractorsFromTable,
  EXPERT_FIELD_OPTIONS,
  ExpertFieldKey,
  ExtractedRawTable,
  ExtractionTarget,
  extractTablesFromFile,
  ORG_FIELD_OPTIONS,
  OrgFieldKey,
} from '../excel/universalExtractor';
import { AtlasDatabaseState, Expert, Partner, Subcontractor } from '../types/database';

export interface UniversalImportResultPayload {
  target: ExtractionTarget;
  experts?: Expert[];
  subcontractors?: Subcontractor[];
  partners?: Partner[];
  skippedDuplicates: number;
}

interface UniversalImportModalProps {
  currentState: AtlasDatabaseState;
  defaultTarget?: ExtractionTarget;
  onImportComplete: (payload: UniversalImportResultPayload) => Promise<void>;
  onClose: () => void;
}

interface TableConfig {
  target: ExtractionTarget;
  expertMapping: ExpertFieldKey[];
  orgMapping: OrgFieldKey[];
}

function normalizeNameKey(value: string): string {
  return value
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

export function UniversalImportModal({
  currentState,
  defaultTarget,
  onImportComplete,
  onClose,
}: UniversalImportModalProps) {
  const [loadedFiles, setLoadedFiles] = useState<string[]>([]);
  const [tables, setTables] = useState<ExtractedRawTable[]>([]);
  const [checkedTableIds, setCheckedTableIds] = useState<Set<string>>(new Set());
  const [inspectedTableId, setInspectedTableId] = useState<string>('');
  const [tableConfigs, setTableConfigs] = useState<Record<string, TableConfig>>({});
  const [mergeSameName, setMergeSameName] = useState(true);
  const [skipExistingInDb, setSkipExistingInDb] = useState(true);
  const [excludedKeys, setExcludedKeys] = useState<Set<string>>(new Set());
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [dragActive, setDragActive] = useState(false);

  const inspectedTable = useMemo(
    () => tables.find((t) => t.id === inspectedTableId) ?? tables[0] ?? null,
    [tables, inspectedTableId]
  );

  const inspectedConfig: TableConfig | null = useMemo(() => {
    if (!inspectedTable) return null;
    return (
      tableConfigs[inspectedTable.id] ?? {
        target: defaultTarget ?? inspectedTable.suggestedTarget,
        expertMapping: autoMapExpertColumns(
          inspectedTable.headers,
          inspectedTable.rows,
          inspectedTable.hyperlinks
        ),
        orgMapping: autoMapOrgColumns(
          inspectedTable.headers,
          inspectedTable.rows,
          inspectedTable.hyperlinks
        ),
      }
    );
  }, [inspectedTable, tableConfigs, defaultTarget]);

  const processFiles = async (fileList: FileList | File[], append = false) => {
    const files = Array.from(fileList);
    if (files.length === 0) return;

    setBusy(true);
    setError(null);
    try {
      const newlyExtracted: ExtractedRawTable[] = [];
      const errors: string[] = [];
      const newFileNames: string[] = [];

      for (const file of files) {
        try {
          const found = await extractTablesFromFile(file);
          if (found.length > 0) {
            // Ensure globally unique IDs across multiple files
            const stamped = found.map((t, idx) => ({
              ...t,
              id: `${t.id}-${Date.now()}-${idx}`,
            }));
            newlyExtracted.push(...stamped);
            newFileNames.push(file.name);
          } else {
            errors.push(`${file.name} : aucun tableau ou liste structurée détecté`);
          }
        } catch (err) {
          errors.push(`${file.name} : ${(err as Error)?.message || 'erreur de lecture'}`);
        }
      }

      if (newlyExtracted.length === 0) {
        if (!append) {
          setTables([]);
          setCheckedTableIds(new Set());
          setLoadedFiles([]);
        }
        setError(
          errors.length > 0
            ? errors.join(' · ')
            : 'Aucun tableau ni liste structurée exploitable n’a été détecté.'
        );
        return;
      }

      if (errors.length > 0) {
        setError(`Certains fichiers n’ont pas pu être lus : ${errors.join(' · ')}`);
      }

      const nextConfigs: Record<string, TableConfig> = append ? { ...tableConfigs } : {};
      for (const tbl of newlyExtracted) {
        nextConfigs[tbl.id] = {
          target: defaultTarget ?? tbl.suggestedTarget,
          expertMapping: autoMapExpertColumns(tbl.headers, tbl.rows, tbl.hyperlinks),
          orgMapping: autoMapOrgColumns(tbl.headers, tbl.rows, tbl.hyperlinks),
        };
      }

      const nextTables = append ? [...tables, ...newlyExtracted] : newlyExtracted;
      const nextChecked = new Set(append ? checkedTableIds : []);
      newlyExtracted.forEach((t) => nextChecked.add(t.id));

      setTableConfigs(nextConfigs);
      setTables(nextTables);
      setCheckedTableIds(nextChecked);
      setLoadedFiles(
        append ? Array.from(new Set([...loadedFiles, ...newFileNames])) : newFileNames
      );
      if (!append || !inspectedTableId) {
        setInspectedTableId(newlyExtracted[0].id);
      }
      setExcludedKeys(new Set());
    } finally {
      setBusy(false);
    }
  };

  const handleFileInput = (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      void processFiles(e.target.files, tables.length > 0);
    }
    e.target.value = '';
  };

  const handleDrop = (e: DragEvent<HTMLLabelElement>) => {
    e.preventDefault();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      void processFiles(e.dataTransfer.files, tables.length > 0);
    }
  };

  const handleResetLoaded = () => {
    setTables([]);
    setCheckedTableIds(new Set());
    setInspectedTableId('');
    setTableConfigs({});
    setLoadedFiles([]);
    setExcludedKeys(new Set());
    setError(null);
  };

  const updateInspectedTarget = (nextTarget: ExtractionTarget) => {
    if (!inspectedTable || !inspectedConfig) return;
    setTableConfigs((prev) => ({
      ...prev,
      [inspectedTable.id]: {
        ...inspectedConfig,
        target: nextTarget,
      },
    }));
    setExcludedKeys(new Set());
  };

  const applyTargetToAllChecked = (nextTarget: ExtractionTarget) => {
    setTableConfigs((prev) => {
      const updated = { ...prev };
      for (const tbl of tables) {
        if (checkedTableIds.has(tbl.id)) {
          const existing = updated[tbl.id];
          if (existing) {
            updated[tbl.id] = { ...existing, target: nextTarget };
          }
        }
      }
      return updated;
    });
    setExcludedKeys(new Set());
  };

  const existingExpertKeys = useMemo(
    () =>
      new Set(
        currentState.experts.map(
          (e) => `${normalizeNameKey(e.firstName)}|${normalizeNameKey(e.lastName)}`
        )
      ),
    [currentState.experts]
  );
  const existingSubKeys = useMemo(
    () => new Set(currentState.subcontractors.map((s) => normalizeNameKey(s.name))),
    [currentState.subcontractors]
  );
  const existingPartnerKeys = useMemo(
    () => new Set(currentState.partners.map((p) => normalizeNameKey(p.name))),
    [currentState.partners]
  );

  type PreviewEntry =
    | {
        rowKey: string;
        sourceLabel: string;
        kind: 'experts';
        item: Expert;
        alreadyExists: boolean;
      }
    | {
        rowKey: string;
        sourceLabel: string;
        kind: 'subcontractors';
        item: Subcontractor;
        alreadyExists: boolean;
      }
    | {
        rowKey: string;
        sourceLabel: string;
        kind: 'partners';
        item: Partner;
        alreadyExists: boolean;
      };

  const previewItems = useMemo(() => {
    const entries: PreviewEntry[] = [];
    const seenBatchExperts = new Set<string>();
    const seenBatchSubs = new Map<string, Subcontractor>();
    const seenBatchPartners = new Map<string, Partner>();

    let expCounter = currentState.experts.length + 1;
    let subCounter = currentState.subcontractors.length + 1;
    let prtCounter = currentState.partners.length + 1;

    for (const tbl of tables) {
      if (!checkedTableIds.has(tbl.id)) continue;
      const cfg = tableConfigs[tbl.id];
      const tblTarget = cfg?.target ?? defaultTarget ?? tbl.suggestedTarget;
      const sourceLabel = `${tbl.sourceFileName} · ${tbl.label}`;

      if (tblTarget === 'experts') {
        const mapping =
          cfg?.expertMapping ??
          autoMapExpertColumns(tbl.headers, tbl.rows, tbl.hyperlinks);
        const built = buildExpertsFromTable(tbl, mapping, expCounter);
        built.forEach((item, idx) => {
          const nameKey = `${normalizeNameKey(item.firstName)}|${normalizeNameKey(item.lastName)}`;
          if (mergeSameName && seenBatchExperts.has(nameKey)) return;
          seenBatchExperts.add(nameKey);
          expCounter++;
          entries.push({
            rowKey: `${tbl.id}-exp-${idx}`,
            sourceLabel,
            kind: 'experts',
            item,
            alreadyExists: existingExpertKeys.has(nameKey),
          });
        });
      } else if (tblTarget === 'subcontractors') {
        const mapping =
          cfg?.orgMapping ?? autoMapOrgColumns(tbl.headers, tbl.rows, tbl.hyperlinks);
        const built = buildSubcontractorsFromTable(tbl, mapping, subCounter, mergeSameName);
        built.forEach((item, idx) => {
          const nameKey = normalizeNameKey(item.name);
          if (mergeSameName && seenBatchSubs.has(nameKey)) {
            // Merge phones/emails/domains across multiple documents!
            const prevItem = seenBatchSubs.get(nameKey)!;
            const prevPhones = prevItem.phone
              .split('/')
              .map((p) => p.trim())
              .filter(Boolean);
            for (const p of item.phone.split('/').map((x) => x.trim()).filter(Boolean)) {
              if (!prevPhones.includes(p)) prevPhones.push(p);
            }
            prevItem.phone = prevPhones.join(' / ');
            if (!prevItem.email && item.email) prevItem.email = item.email;
            if (!prevItem.website && item.website) prevItem.website = item.website;
            for (const d of item.domains) {
              if (!prevItem.domains.includes(d)) prevItem.domains.push(d);
            }
            return;
          }
          const cloned = { ...item, domains: [...item.domains] };
          seenBatchSubs.set(nameKey, cloned);
          subCounter++;
          entries.push({
            rowKey: `${tbl.id}-sub-${idx}`,
            sourceLabel,
            kind: 'subcontractors',
            item: cloned,
            alreadyExists: existingSubKeys.has(nameKey),
          });
        });
      } else {
        const mapping =
          cfg?.orgMapping ?? autoMapOrgColumns(tbl.headers, tbl.rows, tbl.hyperlinks);
        const built = buildPartnersFromTable(tbl, mapping, prtCounter, mergeSameName);
        built.forEach((item, idx) => {
          const nameKey = normalizeNameKey(item.name);
          if (mergeSameName && seenBatchPartners.has(nameKey)) return;
          seenBatchPartners.set(nameKey, item);
          prtCounter++;
          entries.push({
            rowKey: `${tbl.id}-prt-${idx}`,
            sourceLabel,
            kind: 'partners',
            item,
            alreadyExists: existingPartnerKeys.has(nameKey),
          });
        });
      }
    }

    return entries;
  }, [
    tables,
    checkedTableIds,
    tableConfigs,
    defaultTarget,
    mergeSameName,
    currentState,
    existingExpertKeys,
    existingSubKeys,
    existingPartnerKeys,
  ]);

  const selectedToImport = useMemo(
    () =>
      previewItems.filter(
        (entry) =>
          !excludedKeys.has(entry.rowKey) &&
          !(skipExistingInDb && entry.alreadyExists)
      ),
    [previewItems, excludedKeys, skipExistingInDb]
  );

  const skippedDuplicatesCount = useMemo(
    () =>
      skipExistingInDb
        ? previewItems.filter(
            (entry) => !excludedKeys.has(entry.rowKey) && entry.alreadyExists
          ).length
        : 0,
    [previewItems, excludedKeys, skipExistingInDb]
  );

  const handleConfirmImport = async () => {
    if (selectedToImport.length === 0) return;
    setBusy(true);
    try {
      const experts = selectedToImport
        .filter((e): e is Extract<PreviewEntry, { kind: 'experts' }> => e.kind === 'experts')
        .map((e) => e.item);
      const subcontractors = selectedToImport
        .filter(
          (e): e is Extract<PreviewEntry, { kind: 'subcontractors' }> =>
            e.kind === 'subcontractors'
        )
        .map((e) => e.item);
      const partners = selectedToImport
        .filter((e): e is Extract<PreviewEntry, { kind: 'partners' }> => e.kind === 'partners')
        .map((e) => e.item);

      const primaryTarget: ExtractionTarget =
        subcontractors.length >= experts.length && subcontractors.length >= partners.length
          ? 'subcontractors'
          : experts.length >= partners.length
            ? 'experts'
            : 'partners';

      await onImportComplete({
        target: defaultTarget ?? primaryTarget,
        experts: experts.length > 0 ? experts : undefined,
        subcontractors: subcontractors.length > 0 ? subcontractors : undefined,
        partners: partners.length > 0 ? partners : undefined,
        skippedDuplicates: skippedDuplicatesCount,
      });
      onClose();
    } finally {
      setBusy(false);
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 bg-slate-900/55 backdrop-blur-xs flex items-center justify-center p-4"
      role="dialog"
      aria-modal="true"
    >
      <div className="bg-white border border-slate-200 rounded-2xl max-w-5xl w-full max-h-[92vh] flex flex-col shadow-2xl overflow-hidden">
        {/* Header */}
        <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between gap-4 bg-slate-50/70">
          <div className="flex items-center gap-3 min-w-0">
            <div className="w-9 h-9 rounded-lg bg-gerep-blue/10 text-gerep-blue flex items-center justify-center shrink-0">
              <Sparkles className="w-5 h-5" />
            </div>
            <div className="min-w-0">
              <h2 className="text-base font-bold text-slate-900 truncate">
                Extracteur intelligent multi-documents (.xlsx, .docx, .csv)
              </h2>
              <p className="text-xs text-slate-500">
                Parcourt l’ensemble des documents (tableaux Word/Excel, feuilles multiples, cellules fusionnées et listes hors tableaux)
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 text-slate-400 hover:text-slate-700 hover:bg-slate-200/60 rounded-lg"
            aria-label="Fermer"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Body */}
        <div className="flex-1 overflow-y-auto p-6 space-y-5">
          {/* File Dropzone */}
          <div className="flex flex-col sm:flex-row gap-3 items-stretch">
            <label
              onDragOver={(e) => {
                e.preventDefault();
                setDragActive(true);
              }}
              onDragLeave={() => setDragActive(false)}
              onDrop={handleDrop}
              className={`flex-1 flex flex-col sm:flex-row items-center justify-between gap-4 border-2 border-dashed rounded-xl p-4 cursor-pointer transition-colors ${
                dragActive
                  ? 'border-gerep-blue bg-sky-50/70'
                  : 'border-slate-300 bg-slate-50/60 hover:border-gerep-blue hover:bg-sky-50/30'
              }`}
            >
              <div className="flex items-center gap-3.5 min-w-0">
                <div className="flex items-center gap-1.5 text-gerep-blue shrink-0">
                  <FileSpreadsheet className="w-6 h-6 text-gerep-forest" />
                  <FileText className="w-6 h-6 text-gerep-blue" />
                </div>
                <div className="min-w-0">
                  <div className="text-xs font-semibold text-slate-900 truncate">
                    {loadedFiles.length > 0
                      ? `${loadedFiles.length} document${loadedFiles.length > 1 ? 's' : ''} analysé${loadedFiles.length > 1 ? 's' : ''} : ${loadedFiles.join(', ')}`
                      : 'Déposez ou sélectionnez un ou plusieurs documents Excel (.xlsx, .csv) ou Word (.docx)'}
                  </div>
                  <div className="text-[11px] text-slate-500 mt-0.5">
                    Extraction complète : toutes les feuilles Excel, tous les tableaux Word (y compris fusions) et paragraphes/listes structurés
                  </div>
                </div>
              </div>
              <span className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-gerep-blue rounded-lg shrink-0">
                <Upload className="w-3.5 h-3.5" />
                {busy
                  ? 'Analyse...'
                  : loadedFiles.length > 0
                    ? '+ Ajouter des documents'
                    : 'Parcourir'}
              </span>
              <input
                type="file"
                multiple
                accept=".xlsx,.docx,.csv,.tsv"
                onChange={handleFileInput}
                disabled={busy}
                className="sr-only"
              />
            </label>

            {loadedFiles.length > 0 && (
              <button
                type="button"
                onClick={handleResetLoaded}
                className="inline-flex items-center justify-center gap-1.5 px-3.5 py-2 text-xs font-medium text-slate-600 bg-white border border-slate-200 rounded-xl hover:text-rose-700 hover:border-rose-200 hover:bg-rose-50/50 transition-colors shrink-0"
                title="Réinitialiser la sélection de documents"
              >
                <Trash2 className="w-3.5 h-3.5" />
                Vider
              </button>
            )}
          </div>

          {error && (
            <div className="flex items-start gap-2.5 p-3.5 rounded-lg bg-rose-50 border border-rose-200 text-xs text-rose-800">
              <AlertCircle className="w-4 h-4 shrink-0 mt-0.5 text-rose-600" />
              <span>{error}</span>
            </div>
          )}

          {tables.length > 0 && inspectedTable && inspectedConfig && (
            <>
              {/* Step 1 & 2: Detected Tables across all documents + Global Target */}
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-4">
                <div className="lg:col-span-7 bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-2.5">
                  <div className="flex items-center justify-between gap-2">
                    <div className="text-xs font-semibold text-slate-800 flex items-center gap-1.5">
                      <FolderOpen className="w-3.5 h-3.5 text-gerep-blue" />
                      1. Sections / Tableaux détectés ({checkedTableIds.size}/{tables.length} inclus)
                    </div>
                    <div className="flex items-center gap-2 text-[11px]">
                      <button
                        type="button"
                        onClick={() => setCheckedTableIds(new Set(tables.map((t) => t.id)))}
                        className="text-gerep-blue hover:underline font-medium"
                      >
                        Tout cocher
                      </button>
                      <span className="text-slate-300">|</span>
                      <button
                        type="button"
                        onClick={() => setCheckedTableIds(new Set())}
                        className="text-slate-500 hover:underline"
                      >
                        Tout décocher
                      </button>
                    </div>
                  </div>
                  <div className="flex flex-wrap gap-1.5 max-h-[120px] overflow-y-auto pr-1">
                    {tables.map((t) => {
                      const isChecked = checkedTableIds.has(t.id);
                      const isInspected = t.id === inspectedTable.id;
                      const tblCfg = tableConfigs[t.id];
                      const targetShort =
                        tblCfg?.target === 'experts'
                          ? 'Exp.'
                          : tblCfg?.target === 'partners'
                            ? 'Part.'
                            : 'BE/ST';

                      return (
                        <div
                          key={t.id}
                          className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs border transition-colors ${
                            isInspected
                              ? 'bg-white border-gerep-blue text-gerep-blue shadow-2xs font-medium'
                              : isChecked
                                ? 'bg-white border-slate-300 text-slate-700'
                                : 'bg-slate-100/80 border-slate-200 text-slate-400'
                          }`}
                        >
                          <input
                            type="checkbox"
                            checked={isChecked}
                            onChange={() => {
                              setCheckedTableIds((prev) => {
                                const next = new Set(prev);
                                if (next.has(t.id)) next.delete(t.id);
                                else next.add(t.id);
                                return next;
                              });
                            }}
                            className="rounded border-slate-300 text-gerep-blue focus:ring-gerep-blue"
                            title="Inclure cette section dans l’extraction"
                          />
                          <button
                            type="button"
                            onClick={() => setInspectedTableId(t.id)}
                            className="text-left truncate max-w-[240px]"
                            title={`${t.sourceFileName} — ${t.label}`}
                          >
                            {loadedFiles.length > 1 ? `${t.sourceFileName} · ` : ''}
                            {t.label} ({t.rows.length})
                          </button>
                          <span className="text-[10px] font-mono px-1 py-0.5 rounded bg-slate-100 text-slate-600">
                            {targetShort}
                          </span>
                        </div>
                      );
                    })}
                  </div>
                </div>

                <div className="lg:col-span-5 bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-2.5">
                  <div className="flex items-center justify-between gap-2">
                    <div className="text-xs font-semibold text-slate-800">
                      2. Destination ({inspectedTable.label})
                    </div>
                    {tables.length > 1 && (
                      <button
                        type="button"
                        onClick={() => applyTargetToAllChecked(inspectedConfig.target)}
                        className="text-[11px] text-gerep-blue hover:underline font-medium"
                      >
                        Appliquer à tous
                      </button>
                    )}
                  </div>
                  <div className="flex flex-wrap gap-1.5">
                    {(
                      [
                        { id: 'experts', label: 'Les Experts' },
                        {
                          id: 'subcontractors',
                          label: 'Bureaux d’études / Sous-traitants',
                        },
                        { id: 'partners', label: 'Les Partenaires' },
                      ] as Array<{ id: ExtractionTarget; label: string }>
                    ).map((opt) => (
                      <button
                        key={opt.id}
                        type="button"
                        onClick={() => updateInspectedTarget(opt.id)}
                        className={`px-3 py-1.5 rounded-lg text-xs font-medium border transition-colors ${
                          inspectedConfig.target === opt.id
                            ? 'bg-gerep-blue text-white border-gerep-blue'
                            : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-100'
                        }`}
                      >
                        {opt.label}
                      </button>
                    ))}
                  </div>
                </div>
              </div>

              {/* Column Mapping Grid for Inspected Table */}
              <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-3">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div>
                    <h3 className="text-xs font-semibold text-slate-900">
                      3. Colonnes de « {inspectedTable.label} » ({inspectedTable.headers.length} détectées)
                    </h3>
                    <p className="text-[11px] text-slate-500">
                      Cliquez sur un autre tableau ci-dessus pour vérifier son mappage. Les e-mails, téléphones et URLs sont aussi extraits automatiquement sur toute la ligne.
                    </p>
                  </div>
                  <div className="flex flex-wrap items-center gap-4 text-xs text-slate-700">
                    <label className="inline-flex items-center gap-1.5 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={mergeSameName}
                        onChange={(e) => setMergeSameName(e.target.checked)}
                        className="rounded border-slate-300 text-gerep-blue focus:ring-gerep-blue"
                      />
                      <span>Fusionner les lignes identiques</span>
                    </label>
                    <label className="inline-flex items-center gap-1.5 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={skipExistingInDb}
                        onChange={(e) => setSkipExistingInDb(e.target.checked)}
                        className="rounded border-slate-300 text-gerep-blue focus:ring-gerep-blue"
                      />
                      <span>Ignorer les doublons déjà en base</span>
                    </label>
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2.5">
                  {inspectedTable.headers.map((header, colIdx) => {
                    const sampleVal =
                      inspectedTable.rows.find((r) => (r[colIdx] ?? '').trim() !== '')?.[
                        colIdx
                      ] ??
                      inspectedTable.hyperlinks.find(
                        (r) => (r[colIdx] ?? '').trim() !== ''
                      )?.[colIdx] ??
                      '';
                    return (
                      <div
                        key={`${inspectedTable.id}-${header}-${colIdx}`}
                        className="border border-slate-200 rounded-lg p-2.5 bg-slate-50/70 space-y-1.5"
                      >
                        <div className="flex items-center justify-between gap-2">
                          <span className="text-xs font-semibold text-slate-800 truncate">
                            {header}
                          </span>
                          <span className="text-[10px] font-mono text-slate-400 shrink-0">
                            Col {colIdx + 1}
                          </span>
                        </div>
                        {sampleVal && (
                          <div className="text-[11px] text-slate-500 truncate">
                            Ex : {sampleVal}
                          </div>
                        )}
                        {inspectedConfig.target === 'experts' ? (
                          <select
                            value={inspectedConfig.expertMapping[colIdx] ?? 'ignore'}
                            onChange={(e) => {
                              const val = e.target.value as ExpertFieldKey;
                              setTableConfigs((prev) => {
                                const curr = prev[inspectedTable.id] ?? inspectedConfig;
                                const nextMap = [...curr.expertMapping];
                                nextMap[colIdx] = val;
                                return {
                                  ...prev,
                                  [inspectedTable.id]: {
                                    ...curr,
                                    expertMapping: nextMap,
                                  },
                                };
                              });
                            }}
                            className="w-full px-2 py-1.5 text-xs bg-white border border-slate-200 rounded-md focus:outline-none focus:border-gerep-blue"
                          >
                            {EXPERT_FIELD_OPTIONS.map((opt) => (
                              <option key={opt.key} value={opt.key}>
                                {opt.label}
                              </option>
                            ))}
                          </select>
                        ) : (
                          <select
                            value={inspectedConfig.orgMapping[colIdx] ?? 'ignore'}
                            onChange={(e) => {
                              const val = e.target.value as OrgFieldKey;
                              setTableConfigs((prev) => {
                                const curr = prev[inspectedTable.id] ?? inspectedConfig;
                                const nextMap = [...curr.orgMapping];
                                nextMap[colIdx] = val;
                                return {
                                  ...prev,
                                  [inspectedTable.id]: {
                                    ...curr,
                                    orgMapping: nextMap,
                                  },
                                };
                              });
                            }}
                            className="w-full px-2 py-1.5 text-xs bg-white border border-slate-200 rounded-md focus:outline-none focus:border-gerep-blue"
                          >
                            {ORG_FIELD_OPTIONS.map((opt) => (
                              <option key={opt.key} value={opt.key}>
                                {opt.label}
                              </option>
                            ))}
                          </select>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Unified Preview Table across ALL checked tables/documents */}
              <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
                <div className="px-4 py-2.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between gap-4">
                  <span className="text-xs font-semibold text-slate-800">
                    4. Aperçu global avant importation ({selectedToImport.length} prête
                    {selectedToImport.length > 1 ? 's' : ''} sur {previewItems.length})
                  </span>
                  {skippedDuplicatesCount > 0 && (
                    <span className="text-[11px] font-medium text-amber-700">
                      {skippedDuplicatesCount} doublon{skippedDuplicatesCount > 1 ? 's' : ''}{' '}
                      déjà présent{skippedDuplicatesCount > 1 ? 's' : ''} ignoré
                      {skippedDuplicatesCount > 1 ? 's' : ''}
                    </span>
                  )}
                </div>

                <div className="overflow-x-auto max-h-[280px]">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead className="sticky top-0 bg-slate-100 border-b border-slate-200 text-slate-700">
                      <tr>
                        <th className="py-2 px-3 w-8"></th>
                        <th className="py-2 px-3 font-semibold">Nom / Raison sociale</th>
                        <th className="py-2 px-3 font-semibold">Type</th>
                        <th className="py-2 px-3 font-semibold">Pays</th>
                        <th className="py-2 px-3 font-semibold">Domaines</th>
                        <th className="py-2 px-3 font-semibold">E-mail / Tél</th>
                        <th className="py-2 px-3 font-semibold">Source</th>
                        <th className="py-2 px-3 font-semibold">Statut</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {previewItems.map((entry) => {
                        const isExcluded = excludedKeys.has(entry.rowKey);
                        const willSkip = skipExistingInDb && entry.alreadyExists;
                        const checked = !isExcluded && !willSkip;
                        const displayName =
                          entry.kind === 'experts'
                            ? `${entry.item.firstName} ${entry.item.lastName}`.trim()
                            : entry.item.name;
                        const phone =
                          entry.kind === 'subcontractors' || entry.kind === 'experts'
                            ? entry.item.phone ?? ''
                            : '';
                        const kindLabel =
                          entry.kind === 'experts'
                            ? 'Expert'
                            : entry.kind === 'partners'
                              ? 'Partenaire'
                              : 'Bureau / ST';

                        return (
                          <tr
                            key={entry.rowKey}
                            className={
                              willSkip || isExcluded
                                ? 'bg-slate-50/70 opacity-60'
                                : 'hover:bg-slate-50'
                            }
                          >
                            <td className="py-2 px-3">
                              <input
                                type="checkbox"
                                checked={checked}
                                disabled={willSkip}
                                onChange={() => {
                                  setExcludedKeys((prev) => {
                                    const next = new Set(prev);
                                    if (next.has(entry.rowKey)) next.delete(entry.rowKey);
                                    else next.add(entry.rowKey);
                                    return next;
                                  });
                                }}
                                className="rounded border-slate-300 text-gerep-blue"
                              />
                            </td>
                            <td className="py-2 px-3 font-semibold text-slate-900 max-w-[200px] truncate">
                              {displayName}
                            </td>
                            <td className="py-2 px-3 text-[11px] font-medium text-slate-600 whitespace-nowrap">
                              {kindLabel}
                            </td>
                            <td className="py-2 px-3 text-slate-700 whitespace-nowrap">
                              {entry.item.country}
                            </td>
                            <td className="py-2 px-3 text-slate-600 max-w-[200px] truncate">
                              {entry.item.domains.join(', ')}
                            </td>
                            <td className="py-2 px-3 text-slate-700 font-mono text-[11px] max-w-[200px] truncate">
                              {[entry.item.email, phone].filter(Boolean).join(' · ') || '—'}
                            </td>
                            <td className="py-2 px-3 text-slate-500 text-[11px] max-w-[150px] truncate">
                              {entry.sourceLabel}
                            </td>
                            <td className="py-2 px-3 whitespace-nowrap">
                              {entry.alreadyExists ? (
                                <span className="text-[11px] font-medium text-amber-700">
                                  Déjà en base
                                </span>
                              ) : (
                                <span className="inline-flex items-center gap-1 text-[11px] font-medium text-gerep-forest">
                                  <Check className="w-3 h-3" />
                                  Nouveau
                                </span>
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </>
          )}
        </div>

        {/* Footer */}
        <div className="px-6 py-3.5 border-t border-slate-200 bg-slate-50 flex items-center justify-between gap-4">
          <div className="text-xs text-slate-500">
            Formats acceptés : <span className="font-mono">.xlsx</span>,{' '}
            <span className="font-mono">.docx</span> (tableaux &amp; listes Word),{' '}
            <span className="font-mono">.csv</span>
          </div>
          <div className="flex items-center gap-2.5">
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-2 text-xs font-medium text-slate-700 bg-white border border-slate-200 rounded-lg hover:bg-slate-50"
            >
              Annuler
            </button>
            <button
              type="button"
              disabled={busy || selectedToImport.length === 0}
              onClick={handleConfirmImport}
              className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-gerep-blue rounded-lg hover:bg-gerep-blue-dark disabled:opacity-50"
            >
              <Check className="w-3.5 h-3.5" />
              Importer {selectedToImport.length} fiche{selectedToImport.length > 1 ? 's' : ''}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
