import React from 'react';
import {
  BookOpen,
  Eye,
  FileSpreadsheet,
  History,
  Languages,
  LayoutDashboard,
  SlidersHorizontal,
} from 'lucide-react';
import { NavigationTab, WorkbookProject } from '../../types/workbook';

interface SidebarProps {
  activeTab: NavigationTab;
  onSelectTab: (tab: NavigationTab) => void;
  workbooks: WorkbookProject[];
  activeWorkbookId: string | null;
  onSelectWorkbook: (id: string) => void;
}

const NAV_ITEMS: Array<{
  id: NavigationTab;
  label: string;
  icon: React.ComponentType<{ className?: string }>;
}> = [
  { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
  { id: 'characters', label: 'Characters', icon: Languages },
  { id: 'preview', label: 'Page Preview', icon: Eye },
  { id: 'settings', label: 'Workbook Settings', icon: SlidersHorizontal },
  { id: 'import', label: 'Import Excel / CSV', icon: FileSpreadsheet },
  { id: 'history', label: 'Export History', icon: History },
];

export const Sidebar: React.FC<SidebarProps> = ({
  activeTab,
  onSelectTab,
  workbooks,
  activeWorkbookId,
  onSelectWorkbook,
}) => {
  const activeWorkbook = workbooks.find((w) => w.id === activeWorkbookId) ?? null;

  return (
    <aside className="no-print bg-white border-r border-hairline w-full lg:w-[256px] shrink-0 flex flex-col justify-between">
      <div className="p-6">
        {/* Brand Wordmark */}
        <div className="pb-6 border-b border-hairline">
          <div className="font-display text-[26px] font-bold tracking-tight text-ink leading-none whitespace-nowrap">
            Hanzi Book Studio
          </div>
          <p className="mt-1.5 text-xs text-ink-subtle m-0">
            Chinese Workbook Publishing
          </p>
        </div>

        {/* Active Project Selector */}
        {workbooks.length > 0 && (
          <div className="py-5 border-b border-hairline">
            <label
              htmlFor="active-workbook-select"
              className="block text-[11px] font-semibold text-ink-subtle mb-2"
            >
              Active Workbook
            </label>
            <select
              id="active-workbook-select"
              value={activeWorkbookId ?? ''}
              onChange={(e) => onSelectWorkbook(e.target.value)}
              className="w-full bg-ivory border border-hairline-strong px-2.5 py-2 text-xs font-medium text-ink focus-visible:outline-2 focus-visible:outline-burgundy truncate"
            >
              {workbooks.map((wb) => (
                <option key={wb.id} value={wb.id}>
                  {wb.title} ({wb.entries.length})
                </option>
              ))}
            </select>
            {activeWorkbook && (
              <div className="mt-2 flex items-center gap-1.5 text-[11px] text-ink-subtle font-mono tabular-nums">
                <span>{activeWorkbook.settings.pageSize}</span>
                <span aria-hidden="true">·</span>
                <span>{activeWorkbook.settings.orientation}</span>
                <span aria-hidden="true">·</span>
                <span>{activeWorkbook.entries.length} entries</span>
              </div>
            )}
          </div>
        )}

        {/* Primary Navigation */}
        <nav className="mt-5 flex flex-col gap-1" aria-label="Workspace navigation">
          {NAV_ITEMS.map(({ id, label, icon: Icon }) => {
            const isActive = activeTab === id;
            return (
              <button
                key={id}
                type="button"
                onClick={() => onSelectTab(id)}
                className={`w-full flex items-center gap-3 px-3 py-2.5 text-sm text-left transition-colors cursor-pointer whitespace-nowrap ${
                  isActive
                    ? 'bg-burgundy-tint text-burgundy font-semibold'
                    : 'text-ink-muted hover:text-ink hover:bg-ivory font-medium'
                }`}
              >
                <Icon className="w-4 h-4 shrink-0" />
                <span>{label}</span>
              </button>
            );
          })}
        </nav>
      </div>

      {/* Quiet Editorial Footer */}
      <div className="p-6 border-t border-hairline text-xs text-ink-subtle">
        <div className="flex items-center gap-2 text-ink-muted font-medium">
          <BookOpen className="w-3.5 h-3.5 text-burgundy shrink-0" />
          <span>Editable .docx & PDF</span>
        </div>
        <p className="mt-1 text-[11px] text-ink-subtle leading-relaxed m-0">
          Stroke vectors powered by Make Me a Hanzi &amp; Hanzi Writer. Stored locally in IndexedDB.
        </p>
      </div>
    </aside>
  );
};
