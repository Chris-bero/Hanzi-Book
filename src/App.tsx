import React, { useEffect, useMemo, useState } from 'react';
import { CharacterManagerView } from './components/Characters/CharacterManagerView';
import { DashboardView } from './components/Dashboard/DashboardView';
import { ExportHistoryView } from './components/History/ExportHistoryView';
import { BatchImportView } from './components/Import/BatchImportView';
import { Sidebar } from './components/Layout/Sidebar';
import { WorkbookPreviewView } from './components/Preview/WorkbookPreviewView';
import { WorkbookSettingsView } from './components/Settings/WorkbookSettingsView';
import { ensureDictionaryHydrated } from './services/dictionaryService';
import { createWorkbookEntryFromItem } from './services/excelImportService';
import { repository } from './storage/db';
import {
  DEFAULT_WORKBOOK_SETTINGS,
  ExportRecord,
  NavigationTab,
  normalizeWorkbookSettings,
  WorkbookEntry,
  WorkbookProject,
  WorkbookSettings,
} from './types/workbook';

const STARTER_CHARS = ['不', '我', '你', '好', '你好'];

export default function App() {
  const [activeTab, setActiveTab] = useState<NavigationTab>('dashboard');
  const [workbooks, setWorkbooks] = useState<WorkbookProject[]>([]);
  const [activeWorkbookId, setActiveWorkbookId] = useState<string | null>(null);
  const [selectedEntryIds, setSelectedEntryIds] = useState<Set<string>>(new Set());
  const [cachedStrokeCount, setCachedStrokeCount] = useState(0);
  const [exportsHistory, setExportsHistory] = useState<ExportRecord[]>([]);
  const [dictReady, setDictReady] = useState(false);
  const [initializing, setInitializing] = useState(true);

  useEffect(() => {
    let mounted = true;

    const bootstrap = async () => {
      try {
        const [books, strokeCount, exports] = await Promise.all([
          repository.listWorkbooks(),
          repository.countCachedStrokeChars(),
          repository.listExports(),
        ]);

        if (!mounted) return;

        if (books.length > 0) {
          const normalizedBooks = books.map((book) => ({
            ...book,
            settings: normalizeWorkbookSettings(book.settings),
          }));
          // Persist upgraded layout settings cleanly in IndexedDB
          for (let i = 0; i < normalizedBooks.length; i++) {
            if (books[i].settings?.layoutVersion !== 2 || books[i].settings?.traceBoxCount !== 2 || books[i].settings?.emptyBoxCount !== 8) {
              await repository.saveWorkbook(normalizedBooks[i]);
            }
          }
          setWorkbooks(normalizedBooks);
          setActiveWorkbookId(normalizedBooks[0].id);
          setCachedStrokeCount(strokeCount);
          setExportsHistory(exports);
        } else {
          // Seed an immediate functional starter workbook with 不, 我, 你, 好, 你好 using real Hanzi Writer vectors
          const starterEntries: WorkbookEntry[] = [];
          for (let i = 0; i < STARTER_CHARS.length; i++) {
            const { entry } = await createWorkbookEntryFromItem({ text: STARTER_CHARS[i] }, i);
            if (entry) starterEntries.push(entry);
          }
          const now = Date.now();
          const starterProject: WorkbookProject = {
            id: `wb-${now}-starter`,
            title: 'Essential Mandarin Characters — Vol. 1',
            subtitle: 'A4 Landscape Guided Stroke-Order Practice Workbook',
            entries: starterEntries,
            settings: { ...DEFAULT_WORKBOOK_SETTINGS },
            createdAt: now,
            updatedAt: now,
          };
          await repository.saveWorkbook(starterProject);
          if (!mounted) return;
          setWorkbooks([starterProject]);
          setActiveWorkbookId(starterProject.id);
          setCachedStrokeCount(await repository.countCachedStrokeChars());
          setExportsHistory(exports);
        }
      } finally {
        if (mounted) setInitializing(false);
      }

      // Hydrate full Make Me a Hanzi dictionary in background without blocking initial UI
      const hydrated = await ensureDictionaryHydrated();
      if (mounted) setDictReady(hydrated);
    };

    bootstrap();
    return () => {
      mounted = false;
    };
  }, []);

  const activeWorkbook = useMemo(
    () => workbooks.find((wb) => wb.id === activeWorkbookId) ?? null,
    [workbooks, activeWorkbookId]
  );

  const handleCreateWorkbook = async (
    title = 'Untitled Chinese Workbook',
    subtitle = 'A4 Landscape Chinese Writing Practice Book',
    includeSampleSet = false
  ) => {
    const entries: WorkbookEntry[] = [];
    if (includeSampleSet) {
      for (let i = 0; i < STARTER_CHARS.length; i++) {
        const { entry } = await createWorkbookEntryFromItem({ text: STARTER_CHARS[i] }, i);
        if (entry) entries.push(entry);
      }
    }

    const now = Date.now();
    const project: WorkbookProject = {
      id: `wb-${now}-${Math.random().toString(36).slice(2, 7)}`,
      title,
      subtitle,
      entries,
      settings: { ...DEFAULT_WORKBOOK_SETTINGS },
      createdAt: now,
      updatedAt: now,
    };

    await repository.saveWorkbook(project);
    setWorkbooks((prev) => [project, ...prev]);
    setActiveWorkbookId(project.id);
    setSelectedEntryIds(new Set());
    setCachedStrokeCount(await repository.countCachedStrokeChars());
    setActiveTab('characters');
  };

  const persistActiveWorkbook = async (updater: (wb: WorkbookProject) => WorkbookProject) => {
    if (!activeWorkbook) return;
    const updated = updater({ ...activeWorkbook, updatedAt: Date.now() });
    await repository.saveWorkbook(updated);
    setWorkbooks((prev) => prev.map((wb) => (wb.id === updated.id ? updated : wb)));
    setCachedStrokeCount(await repository.countCachedStrokeChars());
  };

  const handleUpdateEntries = async (entries: WorkbookEntry[]) => {
    await persistActiveWorkbook((wb) => ({ ...wb, entries }));
    setSelectedEntryIds((prev) => {
      const validIds = new Set(entries.map((e) => e.id));
      const next = new Set<string>();
      prev.forEach((id) => {
        if (validIds.has(id)) next.add(id);
      });
      return next;
    });
  };

  const handleAppendEntries = async (newEntries: WorkbookEntry[]) => {
    if (!activeWorkbook || newEntries.length === 0) return;
    setWorkbooks((prev) => {
      const target = prev.find((w) => w.id === activeWorkbook.id);
      if (!target) return prev;
      const merged = [...target.entries, ...newEntries].map((e, idx) => ({ ...e, order: idx }));
      const updated: WorkbookProject = {
        ...target,
        entries: merged,
        updatedAt: Date.now(),
      };
      repository.saveWorkbook(updated);
      return prev.map((w) => (w.id === updated.id ? updated : w));
    });
    setCachedStrokeCount(await repository.countCachedStrokeChars());
  };

  const handleUpdateSettings = async (settings: WorkbookSettings) => {
    await persistActiveWorkbook((wb) => ({ ...wb, settings }));
  };

  const handleUpdateWorkbookMeta = async (title: string, subtitle: string) => {
    await persistActiveWorkbook((wb) => ({ ...wb, title, subtitle }));
  };

  const handleDeleteWorkbook = async (id: string) => {
    await repository.deleteWorkbook(id);
    setWorkbooks((prev) => {
      const remaining = prev.filter((wb) => wb.id !== id);
      if (activeWorkbookId === id) {
        setActiveWorkbookId(remaining[0]?.id ?? null);
        setSelectedEntryIds(new Set());
      }
      return remaining;
    });
  };

  const handleSelectWorkbook = (id: string, navigateTo?: NavigationTab) => {
    setActiveWorkbookId(id);
    setSelectedEntryIds(new Set());
    if (navigateTo) setActiveTab(navigateTo);
  };

  const handleToggleSelectEntry = (id: string) => {
    setSelectedEntryIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const handleSelectAllEntries = (ids: string[]) => {
    setSelectedEntryIds(new Set(ids));
  };

  const handleExportComplete = (record: ExportRecord) => {
    setExportsHistory((prev) => [record, ...prev]);
  };

  const handleDeleteExportRecord = async (id: string) => {
    await repository.deleteExport(id);
    setExportsHistory((prev) => prev.filter((r) => r.id !== id));
  };

  return (
    <div className="min-h-screen flex flex-col lg:flex-row bg-ivory text-ink">
      <Sidebar
        activeTab={activeTab}
        onSelectTab={setActiveTab}
        workbooks={workbooks}
        activeWorkbookId={activeWorkbookId}
        onSelectWorkbook={(id) => handleSelectWorkbook(id)}
      />

      <main className="flex-1 min-w-0 px-6 py-7 lg:px-10 lg:py-9 max-w-[1440px]">
        {initializing ? (
          <div className="bg-white border border-hairline p-10 space-y-4">
            <div className="h-7 w-64 bg-ivory-deep animate-pulse" />
            <div className="h-4 w-96 bg-ivory-deep animate-pulse" />
            <div className="grid grid-cols-4 gap-4 pt-4">
              <div className="h-24 bg-ivory-deep animate-pulse" />
              <div className="h-24 bg-ivory-deep animate-pulse" />
              <div className="h-24 bg-ivory-deep animate-pulse" />
              <div className="h-24 bg-ivory-deep animate-pulse" />
            </div>
          </div>
        ) : (
          <>
            {activeTab === 'dashboard' && (
              <DashboardView
                workbooks={workbooks}
                activeWorkbookId={activeWorkbookId}
                cachedStrokeCount={cachedStrokeCount}
                exportsHistory={exportsHistory}
                onCreateWorkbook={handleCreateWorkbook}
                onSelectWorkbook={handleSelectWorkbook}
                onDeleteWorkbook={handleDeleteWorkbook}
                onNavigate={setActiveTab}
              />
            )}

            {activeTab === 'characters' && (
              <CharacterManagerView
                workbook={activeWorkbook}
                dictReady={dictReady}
                selectedEntryIds={selectedEntryIds}
                onToggleSelectEntry={handleToggleSelectEntry}
                onSelectAllEntries={handleSelectAllEntries}
                onUpdateEntries={handleUpdateEntries}
                onGoToPreview={() => setActiveTab('preview')}
                onCreateWorkbook={() => handleCreateWorkbook()}
              />
            )}

            {activeTab === 'preview' && (
              <WorkbookPreviewView
                workbook={activeWorkbook}
                selectedEntryIds={selectedEntryIds}
                onExportComplete={handleExportComplete}
                onGoToCharacters={() => setActiveTab('characters')}
                onGoToSettings={() => setActiveTab('settings')}
              />
            )}

            {activeTab === 'settings' && (
              <WorkbookSettingsView
                workbook={activeWorkbook}
                onUpdateWorkbookMeta={handleUpdateWorkbookMeta}
                onUpdateSettings={handleUpdateSettings}
                onGoToPreview={() => setActiveTab('preview')}
              />
            )}

            {activeTab === 'import' && (
              <BatchImportView
                workbook={activeWorkbook}
                onAppendEntries={handleAppendEntries}
                onGoToCharacters={() => setActiveTab('characters')}
                onGoToPreview={() => setActiveTab('preview')}
              />
            )}

            {activeTab === 'history' && (
              <ExportHistoryView
                exportsHistory={exportsHistory}
                workbooks={workbooks}
                onOpenWorkbook={handleSelectWorkbook}
                onDeleteExportRecord={handleDeleteExportRecord}
              />
            )}
          </>
        )}
      </main>
    </div>
  );
}
