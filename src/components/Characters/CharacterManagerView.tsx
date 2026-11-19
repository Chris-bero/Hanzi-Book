import React, { useEffect, useMemo, useState } from 'react';
import {
  AlertTriangle,
  ArrowDown,
  ArrowUp,
  CheckCircle2,
  Edit3,
  Eye,
  Languages,
  Plus,
  Search,
  Trash2,
} from 'lucide-react';
import { translateChineseToArabic } from '../../services/arabicDictionary';
import { lookupChineseText } from '../../services/dictionaryService';
import { createWorkbookEntryFromItem } from '../../services/excelImportService';
import { extractChineseCharacters, resolveEntryStrokes } from '../../services/strokeEngine';
import { EntryStatus, WorkbookEntry, WorkbookProject } from '../../types/workbook';
import { StrokeSequenceModal } from './StrokeSequenceModal';

interface CharacterManagerViewProps {
  workbook: WorkbookProject | null;
  dictReady: boolean;
  selectedEntryIds: Set<string>;
  onToggleSelectEntry: (id: string) => void;
  onSelectAllEntries: (ids: string[]) => void;
  onUpdateEntries: (entries: WorkbookEntry[]) => Promise<void>;
  onGoToPreview: () => void;
  onCreateWorkbook: () => void;
}

export const CharacterManagerView: React.FC<CharacterManagerViewProps> = ({
  workbook,
  dictReady,
  selectedEntryIds,
  onToggleSelectEntry,
  onSelectAllEntries,
  onUpdateEntries,
  onGoToPreview,
  onCreateWorkbook,
}) => {
  const [inputMode, setInputMode] = useState<'single' | 'bulk'>('single');
  const [textInput, setTextInput] = useState('');
  const [pinyinInput, setPinyinInput] = useState('');
  const [meaningInput, setMeaningInput] = useState('');
  const [arabicInput, setArabicInput] = useState('');
  const [bulkInput, setBulkInput] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isFillingArabic, setIsFillingArabic] = useState(false);
  const [feedback, setFeedback] = useState<string | null>(null);

  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState<'all' | 'ready' | 'issues'>('all');

  const [editingId, setEditingId] = useState<string | null>(null);
  const [editText, setEditText] = useState('');
  const [editPinyin, setEditPinyin] = useState('');
  const [editMeaning, setEditMeaning] = useState('');
  const [editArabic, setEditArabic] = useState('');

  const [previewEntry, setPreviewEntry] = useState<WorkbookEntry | null>(null);

  // Live auto-lookup for single/word input
  useEffect(() => {
    let active = true;
    const chars = extractChineseCharacters(textInput);
    if (chars.length === 0) return;

    const timer = setTimeout(async () => {
      const lookup = await lookupChineseText(chars.join(''));
      if (!active) return;
      setPinyinInput((prev) => (prev.trim() ? prev : lookup.pinyin));
      setMeaningInput((prev) => (prev.trim() ? prev : lookup.meaning));
      setArabicInput((prev) => (prev.trim() ? prev : lookup.arabicMeaning));
    }, 180);

    return () => {
      active = false;
      clearTimeout(timer);
    };
  }, [textInput]);

  const filteredEntries = useMemo(() => {
    if (!workbook) return [];
    const q = searchQuery.trim().toLowerCase();
    return workbook.entries.filter((entry) => {
      if (statusFilter === 'ready' && entry.status !== 'ready') return false;
      if (statusFilter === 'issues' && entry.status === 'ready') return false;
      if (!q) return true;
      return (
        entry.text.includes(q) ||
        entry.pinyin.toLowerCase().includes(q) ||
        entry.meaning.toLowerCase().includes(q) ||
        (entry.arabicMeaning ?? '').toLowerCase().includes(q)
      );
    });
  }, [workbook, searchQuery, statusFilter]);

  if (!workbook) {
    return (
      <div className="bg-white border border-hairline p-12 text-center">
        <h2 className="font-display text-3xl font-bold text-ink m-0">No Active Workbook</h2>
        <p className="mt-2 text-sm text-ink-muted">
          Create a workbook first to manage characters, retrieve stroke orders, and design sheets.
        </p>
        <button
          type="button"
          onClick={onCreateWorkbook}
          className="mt-5 inline-flex items-center gap-2 px-5 py-2.5 text-sm font-semibold bg-burgundy text-white hover:bg-burgundy-dark cursor-pointer"
        >
          <Plus className="w-4 h-4" />
          Create New Workbook
        </button>
      </div>
    );
  }

  const handleAddSingle = async (e: React.FormEvent) => {
    e.preventDefault();
    const chars = extractChineseCharacters(textInput);
    if (chars.length === 0) {
      setFeedback('Enter at least one Chinese character (e.g. 不, 我, or 你好).');
      return;
    }

    setIsSubmitting(true);
    setFeedback(null);
    try {
      const { entry } = await createWorkbookEntryFromItem(
        {
          text: textInput,
          pinyin: pinyinInput,
          meaning: meaningInput,
          arabicMeaning: arabicInput,
        },
        workbook.entries.length
      );
      if (entry) {
        await onUpdateEntries([...workbook.entries, entry]);
        setTextInput('');
        setPinyinInput('');
        setMeaningInput('');
        setArabicInput('');
        setFeedback(
          entry.status === 'ready'
            ? `Added “${entry.text}” (${entry.totalStrokeCount} strokes verified).`
            : `Added “${entry.text}” with missing stroke-data warning.`
        );
      }
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleAddBulk = async (e: React.FormEvent) => {
    e.preventDefault();
    const tokens = bulkInput
      .split(/[\n,，;；\s]+/)
      .map((t) => extractChineseCharacters(t).join(''))
      .filter(Boolean);

    if (tokens.length === 0) {
      setFeedback('No Chinese characters detected in bulk input.');
      return;
    }

    setIsSubmitting(true);
    setFeedback(null);
    try {
      const created: WorkbookEntry[] = [];
      for (let i = 0; i < tokens.length; i++) {
        const { entry } = await createWorkbookEntryFromItem(
          { text: tokens[i] },
          workbook.entries.length + i
        );
        if (entry) created.push(entry);
      }
      await onUpdateEntries([...workbook.entries, ...created]);
      setBulkInput('');
      setFeedback(`Added ${created.length} entries to “${workbook.title}”.`);
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleLoadTestSet = async () => {
    setIsSubmitting(true);
    setFeedback(null);
    try {
      const samples = ['不', '我', '你', '好', '你好'];
      const created: WorkbookEntry[] = [];
      for (let i = 0; i < samples.length; i++) {
        const { entry } = await createWorkbookEntryFromItem(
          { text: samples[i] },
          workbook.entries.length + i
        );
        if (entry) created.push(entry);
      }
      await onUpdateEntries([...workbook.entries, ...created]);
      setFeedback('Added verified test set: 不 (4), 我 (7), 你 (7), 好 (6), and 你好 (13).');
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleMove = async (id: string, delta: -1 | 1) => {
    const index = workbook.entries.findIndex((e) => e.id === id);
    const target = index + delta;
    if (index < 0 || target < 0 || target >= workbook.entries.length) return;
    const copy = [...workbook.entries];
    const [moved] = copy.splice(index, 1);
    copy.splice(target, 0, moved);
    await onUpdateEntries(copy.map((item, idx) => ({ ...item, order: idx })));
  };

  const handleDelete = async (id: string) => {
    const remaining = workbook.entries
      .filter((entry) => entry.id !== id)
      .map((entry, idx) => ({ ...entry, order: idx }));
    await onUpdateEntries(remaining);
  };

  const handleAutoFillArabic = async () => {
    const missing = workbook.entries.filter((e) => !e.arabicMeaning?.trim());
    if (missing.length === 0) return;

    setIsFillingArabic(true);
    setFeedback(null);
    try {
      let filled = 0;
      const updated = [...workbook.entries];
      for (let i = 0; i < updated.length; i++) {
        if (!updated[i].arabicMeaning?.trim()) {
          const ar = await translateChineseToArabic(updated[i].text);
          if (ar) {
            updated[i] = {
              ...updated[i],
              arabicMeaning: ar,
              updatedAt: Date.now(),
            };
            filled++;
          }
        }
      }
      await onUpdateEntries(updated);
      setFeedback(
        filled > 0
          ? `Translated ${filled} ${filled === 1 ? 'entry' : 'entries'} into Arabic.`
          : 'No additional Arabic translations could be resolved automatically. You can edit any row manually.'
      );
    } finally {
      setIsFillingArabic(false);
    }
  };

  const startEdit = (entry: WorkbookEntry) => {
    setEditingId(entry.id);
    setEditText(entry.text);
    setEditPinyin(entry.pinyin);
    setEditMeaning(entry.meaning);
    setEditArabic(entry.arabicMeaning ?? '');
  };

  const saveEdit = async (id: string) => {
    const original = workbook.entries.find((e) => e.id === id);
    if (!original) return;
    const cleanNewText = extractChineseCharacters(editText).join('');
    if (!cleanNewText) return;

    setIsSubmitting(true);
    try {
      const charResults =
        cleanNewText === original.text
          ? original.charResults
          : (await resolveEntryStrokes(cleanNewText)).results;

      const readyCount = charResults.filter((c) => c.status === 'ready').length;
      const status: EntryStatus =
        readyCount === charResults.length
          ? 'ready'
          : readyCount > 0
          ? 'partial'
          : 'unsupported';

      const updated = workbook.entries.map((entry) =>
        entry.id === id
          ? {
              ...entry,
              text: cleanNewText,
              pinyin: editPinyin.trim(),
              meaning: editMeaning.trim(),
              arabicMeaning: editArabic.trim(),
              isManualPinyin: true,
              isManualMeaning: true,
              isManualArabicMeaning: true,
              charResults,
              totalStrokeCount: charResults.reduce((sum, c) => sum + c.strokeCount, 0),
              status,
              updatedAt: Date.now(),
            }
          : entry
      );

      await onUpdateEntries(updated);
      setEditingId(null);
    } finally {
      setIsSubmitting(false);
    }
  };

  const unsupportedCount = workbook.entries.filter((e) => e.status !== 'ready').length;
  const missingArabicCount = workbook.entries.filter((e) => !e.arabicMeaning?.trim()).length;
  const allVisibleSelected =
    filteredEntries.length > 0 && filteredEntries.every((e) => selectedEntryIds.has(e.id));

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex flex-wrap items-end justify-between gap-4 pb-5 border-b border-hairline">
        <div>
          <h1 className="font-display text-4xl font-bold text-ink leading-tight m-0">
            Character Management
          </h1>
          <div className="mt-1.5 flex flex-wrap items-center gap-2 text-xs text-ink-muted">
            <span className="font-semibold text-ink">{workbook.title}</span>
            <span aria-hidden="true">·</span>
            <span className="tabular-nums">{workbook.entries.length} entries</span>
            <span aria-hidden="true">·</span>
            <span>
              Dictionary: {dictReady ? 'Make Me a Hanzi cached locally' : 'Syncing dictionary…'}
            </span>
          </div>
        </div>

        <div className="flex items-center gap-2.5 flex-wrap">
          {missingArabicCount > 0 && (
            <button
              type="button"
              onClick={handleAutoFillArabic}
              disabled={isFillingArabic || isSubmitting}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium border border-hairline-strong bg-white text-ink hover:bg-ivory-deep disabled:opacity-50 transition-colors cursor-pointer whitespace-nowrap"
            >
              <Languages className="w-3.5 h-3.5 text-burgundy" />
              {isFillingArabic
                ? 'Translating to Arabic…'
                : `Auto-Fill Arabic (${missingArabicCount})`}
            </button>
          )}
          <button
            type="button"
            onClick={handleLoadTestSet}
            disabled={isSubmitting}
            className="px-3.5 py-2 text-xs font-medium border border-hairline-strong bg-white text-ink hover:bg-ivory-deep disabled:opacity-50 transition-colors cursor-pointer whitespace-nowrap"
          >
            Add Test Set (不 我 你 好 你好)
          </button>
          <button
            type="button"
            onClick={onGoToPreview}
            className="inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold bg-burgundy text-white hover:bg-burgundy-dark transition-colors cursor-pointer whitespace-nowrap"
          >
            <Eye className="w-3.5 h-3.5" />
            Page Preview &amp; Export
          </button>
        </div>
      </div>

      {/* Add Character / Word Panel */}
      <section className="bg-white border border-hairline p-6">
        <div className="flex items-center justify-between flex-wrap gap-3 pb-4 mb-5 border-b border-hairline">
          <div className="inline-flex bg-ivory p-1 border border-hairline">
            <button
              type="button"
              onClick={() => setInputMode('single')}
              className={`px-3.5 py-1.5 text-xs font-semibold transition-colors cursor-pointer whitespace-nowrap ${
                inputMode === 'single'
                  ? 'bg-white text-burgundy shadow-xs'
                  : 'text-ink-muted hover:text-ink'
              }`}
            >
              Single Character or Word
            </button>
            <button
              type="button"
              onClick={() => setInputMode('bulk')}
              className={`px-3.5 py-1.5 text-xs font-semibold transition-colors cursor-pointer whitespace-nowrap ${
                inputMode === 'bulk'
                  ? 'bg-white text-burgundy shadow-xs'
                  : 'text-ink-muted hover:text-ink'
              }`}
            >
              Bulk Character Input
            </button>
          </div>
          <span className="text-xs text-ink-subtle">
            Stroke vectors are retrieved directly from Hanzi Writer Data
          </span>
        </div>

        {inputMode === 'single' ? (
          <form onSubmit={handleAddSingle} className="space-y-4">
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
              <div>
                <label
                  htmlFor="hanzi-input"
                  className="block text-xs font-semibold text-ink-muted mb-1.5"
                >
                  Chinese Character or Word
                </label>
                <input
                  id="hanzi-input"
                  type="text"
                  value={textInput}
                  onChange={(e) => {
                    setTextInput(e.target.value);
                    setPinyinInput('');
                    setMeaningInput('');
                    setArabicInput('');
                  }}
                  placeholder="e.g. 不, 我, or 你好"
                  className="w-full border border-hairline-strong bg-white px-3.5 py-2.5 text-sm text-ink focus-visible:outline-2 focus-visible:outline-burgundy"
                />
              </div>

              <div>
                <label
                  htmlFor="pinyin-input"
                  className="block text-xs font-semibold text-ink-muted mb-1.5"
                >
                  Pinyin (auto-filled or manual)
                </label>
                <input
                  id="pinyin-input"
                  type="text"
                  value={pinyinInput}
                  onChange={(e) => setPinyinInput(e.target.value)}
                  placeholder="e.g. nǐ hǎo"
                  className="w-full border border-hairline-strong bg-white px-3.5 py-2.5 text-sm text-ink focus-visible:outline-2 focus-visible:outline-burgundy"
                />
              </div>

              <div>
                <label
                  htmlFor="meaning-input"
                  className="block text-xs font-semibold text-ink-muted mb-1.5"
                >
                  English Meaning (auto-filled or manual)
                </label>
                <input
                  id="meaning-input"
                  type="text"
                  value={meaningInput}
                  onChange={(e) => setMeaningInput(e.target.value)}
                  placeholder="e.g. hello; hi"
                  className="w-full border border-hairline-strong bg-white px-3.5 py-2.5 text-sm text-ink focus-visible:outline-2 focus-visible:outline-burgundy"
                />
              </div>

              <div>
                <label
                  htmlFor="arabic-input"
                  className="block text-xs font-semibold text-ink-muted mb-1.5"
                >
                  Arabic Translation (الترجمة العربية)
                </label>
                <input
                  id="arabic-input"
                  type="text"
                  dir="rtl"
                  value={arabicInput}
                  onChange={(e) => setArabicInput(e.target.value)}
                  placeholder="مثال: مرحباً"
                  className="w-full border border-hairline-strong bg-white px-3.5 py-2.5 text-sm text-ink focus-visible:outline-2 focus-visible:outline-burgundy"
                  style={{ fontFamily: 'var(--font-arabic)' }}
                />
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-4">
              <button
                type="submit"
                disabled={isSubmitting}
                className="inline-flex items-center gap-2 px-5 py-2.5 text-xs font-semibold bg-burgundy text-white hover:bg-burgundy-dark disabled:opacity-55 transition-colors cursor-pointer whitespace-nowrap"
              >
                <Plus className="w-4 h-4" />
                {isSubmitting ? 'Retrieving Stroke Data…' : 'Add Entry'}
              </button>
              {feedback && <span className="text-xs text-ink-muted">{feedback}</span>}
            </div>
          </form>
        ) : (
          <form onSubmit={handleAddBulk} className="space-y-4">
            <div>
              <label
                htmlFor="bulk-hanzi-input"
                className="block text-xs font-semibold text-ink-muted mb-1.5"
              >
                Paste Chinese Characters or Words (separated by new lines, commas, or spaces)
              </label>
              <textarea
                id="bulk-hanzi-input"
                rows={4}
                value={bulkInput}
                onChange={(e) => setBulkInput(e.target.value)}
                placeholder={'不\n我\n你\n好\n你好\n学习'}
                className="w-full border border-hairline-strong bg-white px-3.5 py-2.5 text-sm text-ink focus-visible:outline-2 focus-visible:outline-burgundy"
              />
            </div>

            <div className="flex flex-wrap items-center gap-4">
              <button
                type="submit"
                disabled={isSubmitting}
                className="inline-flex items-center gap-2 px-5 py-2.5 text-xs font-semibold bg-burgundy text-white hover:bg-burgundy-dark disabled:opacity-55 transition-colors cursor-pointer whitespace-nowrap"
              >
                <Plus className="w-4 h-4" />
                {isSubmitting ? 'Processing Entries…' : 'Add All Entries'}
              </button>
              {feedback && <span className="text-xs text-ink-muted">{feedback}</span>}
            </div>
          </form>
        )}
      </section>

      {/* Unsupported Warning Banner */}
      {unsupportedCount > 0 && (
        <div className="border border-[#E5C185] bg-[#FFF8EB] px-5 py-3.5 flex items-center gap-3 text-xs text-[#7D4300]">
          <AlertTriangle className="w-4 h-4 shrink-0" />
          <span>
            <strong className="font-semibold">{unsupportedCount}</strong>{' '}
            {unsupportedCount === 1 ? 'entry contains' : 'entries contain'} character(s) unavailable
            in the Make Me a Hanzi dataset. Edit the character or remove it before exporting.
          </span>
        </div>
      )}

      {/* Entries Data Table */}
      <section className="bg-white border border-hairline">
        {/* Table Toolbar */}
        <div className="px-5 py-4 border-b border-hairline flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-3 flex-wrap">
            <div className="relative">
              <Search className="w-3.5 h-3.5 text-ink-subtle absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Filter by character, pinyin, English, or Arabic…"
                className="pl-8 pr-3 py-1.5 text-xs border border-hairline-strong bg-ivory/60 text-ink w-72 focus-visible:outline-2 focus-visible:outline-burgundy"
              />
            </div>

            <div className="inline-flex bg-ivory p-0.5 border border-hairline">
              {(['all', 'ready', 'issues'] as const).map((tab) => (
                <button
                  key={tab}
                  type="button"
                  onClick={() => setStatusFilter(tab)}
                  className={`px-2.5 py-1 text-xs font-medium transition-colors cursor-pointer whitespace-nowrap ${
                    statusFilter === tab
                      ? 'bg-white text-burgundy font-semibold'
                      : 'text-ink-muted hover:text-ink'
                  }`}
                >
                  {tab === 'all'
                    ? `All (${workbook.entries.length})`
                    : tab === 'ready'
                    ? 'Verified'
                    : `Warnings (${unsupportedCount})`}
                </button>
              ))}
            </div>
          </div>

          <div className="flex items-center gap-3 text-xs text-ink-muted tabular-nums">
            {selectedEntryIds.size > 0 && (
              <span className="font-medium text-burgundy">
                {selectedEntryIds.size} selected for export
              </span>
            )}
            <button
              type="button"
              onClick={() =>
                onSelectAllEntries(allVisibleSelected ? [] : filteredEntries.map((e) => e.id))
              }
              className="text-xs font-medium text-ink-muted hover:text-ink underline cursor-pointer"
            >
              {allVisibleSelected ? 'Clear selection' : 'Select all visible'}
            </button>
          </div>
        </div>

        {filteredEntries.length === 0 ? (
          <div className="px-6 py-12 text-center text-sm text-ink-muted">
            {workbook.entries.length === 0
              ? 'No characters added yet. Enter a character above or click “Add Test Set”.'
              : 'No entries match your current filter.'}
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full border-collapse text-sm">
              <thead>
                <tr className="border-b border-hairline bg-ivory/70 text-left text-[11px] font-semibold text-ink-subtle">
                  <th className="py-2.5 px-4 w-10"></th>
                  <th className="py-2.5 px-2 w-12 text-center">#</th>
                  <th className="py-2.5 px-3 w-24">Character</th>
                  <th className="py-2.5 px-3 w-32">Pinyin</th>
                  <th className="py-2.5 px-3">English Meaning</th>
                  <th className="py-2.5 px-3 w-40 text-right">Arabic (العربية)</th>
                  <th className="py-2.5 px-3 w-20 text-right">Strokes</th>
                  <th className="py-2.5 px-3 w-24">Data</th>
                  <th className="py-2.5 px-4 text-right">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-hairline">
                {filteredEntries.map((entry) => {
                  const fullIdx = workbook.entries.findIndex((e) => e.id === entry.id);
                  const isEditing = editingId === entry.id;
                  const isSelected = selectedEntryIds.has(entry.id);

                  return (
                    <tr
                      key={entry.id}
                      className={`transition-colors ${
                        isSelected ? 'bg-burgundy-tint/30' : 'hover:bg-ivory/60'
                      }`}
                    >
                      <td className="py-3 px-4">
                        <input
                          type="checkbox"
                          checked={isSelected}
                          onChange={() => onToggleSelectEntry(entry.id)}
                          aria-label={`Select ${entry.text}`}
                          className="accent-burgundy w-3.5 h-3.5 cursor-pointer"
                        />
                      </td>
                      <td className="py-3 px-2 text-center font-mono text-xs text-ink-subtle tabular-nums">
                        {fullIdx + 1}
                      </td>

                      {isEditing ? (
                        <td colSpan={7} className="py-3 px-3">
                          <div className="grid grid-cols-1 md:grid-cols-[100px_135px_1fr_170px_auto] gap-2.5 items-center">
                            <input
                              type="text"
                              value={editText}
                              onChange={(e) => setEditText(e.target.value)}
                              className="border border-hairline-strong px-2.5 py-1.5 text-sm"
                              aria-label="Edit Chinese text"
                            />
                            <input
                              type="text"
                              value={editPinyin}
                              onChange={(e) => setEditPinyin(e.target.value)}
                              placeholder="Pinyin"
                              className="border border-hairline-strong px-2.5 py-1.5 text-sm"
                              aria-label="Edit pinyin"
                            />
                            <input
                              type="text"
                              value={editMeaning}
                              onChange={(e) => setEditMeaning(e.target.value)}
                              placeholder="English meaning"
                              className="border border-hairline-strong px-2.5 py-1.5 text-sm"
                              aria-label="Edit meaning"
                            />
                            <input
                              type="text"
                              dir="rtl"
                              value={editArabic}
                              onChange={(e) => setEditArabic(e.target.value)}
                              placeholder="الترجمة العربية"
                              className="border border-hairline-strong px-2.5 py-1.5 text-sm"
                              style={{ fontFamily: 'var(--font-arabic)' }}
                              aria-label="Edit Arabic translation"
                            />
                            <div className="flex items-center gap-2">
                              <button
                                type="button"
                                onClick={() => saveEdit(entry.id)}
                                disabled={isSubmitting}
                                className="px-3 py-1.5 text-xs font-semibold bg-burgundy text-white hover:bg-burgundy-dark cursor-pointer"
                              >
                                Save
                              </button>
                              <button
                                type="button"
                                onClick={() => setEditingId(null)}
                                className="px-3 py-1.5 text-xs font-medium border border-hairline-strong text-ink hover:bg-ivory-deep cursor-pointer"
                              >
                                Cancel
                              </button>
                            </div>
                          </div>
                        </td>
                      ) : (
                        <>
                          <td
                            className="py-3 px-3 text-2xl text-ink leading-none"
                            style={{ fontFamily: workbook.settings.fontFamily }}
                          >
                            {entry.text}
                          </td>
                          <td className="py-3 px-3 font-medium text-burgundy">
                            {entry.pinyin || '—'}
                          </td>
                          <td className="py-3 px-3 text-ink-muted text-xs leading-relaxed">
                            {entry.meaning || '—'}
                          </td>
                          <td
                            dir="rtl"
                            className="py-3 px-3 text-right text-ink text-xs leading-relaxed font-medium"
                            style={{ fontFamily: 'var(--font-arabic)' }}
                          >
                            {entry.arabicMeaning || '—'}
                          </td>
                          <td className="py-3 px-3 text-right font-mono text-xs text-ink tabular-nums">
                            {entry.totalStrokeCount}
                          </td>
                          <td className="py-3 px-3 text-xs">
                            {entry.status === 'ready' ? (
                              <span className="inline-flex items-center gap-1 text-[#1F6136] font-medium">
                                <CheckCircle2 className="w-3.5 h-3.5" />
                                Verified
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1 text-[#A51D24] font-medium">
                                <AlertTriangle className="w-3.5 h-3.5" />
                                {entry.status === 'partial' ? 'Partial' : 'Missing'}
                              </span>
                            )}
                          </td>
                          <td className="py-3 px-4">
                            <div className="flex items-center justify-end gap-1">
                              <button
                                type="button"
                                onClick={() => setPreviewEntry(entry)}
                                className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium border border-hairline-strong bg-white text-ink hover:bg-ivory-deep transition-colors cursor-pointer whitespace-nowrap"
                              >
                                <Eye className="w-3.5 h-3.5" />
                                Strokes
                              </button>
                              <button
                                type="button"
                                onClick={() => handleMove(entry.id, -1)}
                                disabled={fullIdx === 0}
                                aria-label="Move up"
                                className="p-1.5 text-ink-muted hover:text-ink disabled:opacity-30 cursor-pointer"
                              >
                                <ArrowUp className="w-3.5 h-3.5" />
                              </button>
                              <button
                                type="button"
                                onClick={() => handleMove(entry.id, 1)}
                                disabled={fullIdx === workbook.entries.length - 1}
                                aria-label="Move down"
                                className="p-1.5 text-ink-muted hover:text-ink disabled:opacity-30 cursor-pointer"
                              >
                                <ArrowDown className="w-3.5 h-3.5" />
                              </button>
                              <button
                                type="button"
                                onClick={() => startEdit(entry)}
                                aria-label={`Edit ${entry.text}`}
                                className="p-1.5 text-ink-muted hover:text-ink cursor-pointer"
                              >
                                <Edit3 className="w-3.5 h-3.5" />
                              </button>
                              <button
                                type="button"
                                onClick={() => handleDelete(entry.id)}
                                aria-label={`Delete ${entry.text}`}
                                className="p-1.5 text-ink-subtle hover:text-[#A51D24] cursor-pointer"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          </td>
                        </>
                      )}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </section>

      {previewEntry && (
        <StrokeSequenceModal
          entry={previewEntry}
          settings={workbook.settings}
          onClose={() => setPreviewEntry(null)}
        />
      )}
    </div>
  );
};
