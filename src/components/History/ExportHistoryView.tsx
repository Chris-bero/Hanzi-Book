import React from 'react';
import { FileDown, FileText, FolderOpen, Trash2 } from 'lucide-react';
import { ExportRecord, NavigationTab, WorkbookProject } from '../../types/workbook';

interface ExportHistoryViewProps {
  exportsHistory: ExportRecord[];
  workbooks: WorkbookProject[];
  onOpenWorkbook: (workbookId: string, tab: NavigationTab) => void;
  onDeleteExportRecord: (id: string) => Promise<void>;
}

export const ExportHistoryView: React.FC<ExportHistoryViewProps> = ({
  exportsHistory,
  workbooks,
  onOpenWorkbook,
  onDeleteExportRecord,
}) => {
  return (
    <div className="space-y-6">
      <div className="flex flex-wrap items-end justify-between gap-4 pb-5 border-b border-hairline">
        <div>
          <h1 className="font-display text-4xl font-bold text-ink leading-tight m-0">
            Export History
          </h1>
          <p className="mt-1.5 text-sm text-ink-muted m-0">
            Log of Microsoft Word (.docx) and printable PDF editions generated in this workspace.
          </p>
        </div>
        <span className="font-mono text-xs text-ink-subtle tabular-nums">
          {exportsHistory.length} {exportsHistory.length === 1 ? 'record' : 'records'}
        </span>
      </div>

      <section className="bg-white border border-hairline">
        {exportsHistory.length === 0 ? (
          <div className="p-12 text-center">
            <FileDown className="w-8 h-8 text-ink-subtle mx-auto mb-3" />
            <h2 className="font-display text-2xl font-bold text-ink m-0">No Exports Recorded Yet</h2>
            <p className="mt-2 text-sm text-ink-muted max-w-md mx-auto">
              Open Page Preview on any workbook and export an editable .docx or printable .pdf file
              to build your publishing history.
            </p>
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full border-collapse text-sm">
              <thead>
                <tr className="border-b border-hairline bg-ivory/70 text-left text-[11px] font-semibold text-ink-subtle">
                  <th className="py-3 px-5">File &amp; Workbook</th>
                  <th className="py-3 px-4 w-24">Format</th>
                  <th className="py-3 px-4 w-36">Page Layout</th>
                  <th className="py-3 px-4 w-24 text-right">Entries</th>
                  <th className="py-3 px-4 w-44">Exported</th>
                  <th className="py-3 px-5 text-right">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-hairline">
                {exportsHistory.map((rec) => {
                  const stillExists = workbooks.some((w) => w.id === rec.workbookId);
                  return (
                    <tr key={rec.id} className="hover:bg-ivory/50">
                      <td className="py-3.5 px-5">
                        <div className="flex items-center gap-2.5">
                          <FileText className="w-4 h-4 text-burgundy shrink-0" />
                          <div>
                            <div className="font-medium text-ink text-xs font-mono">
                              {rec.fileName}
                            </div>
                            <div className="text-xs text-ink-subtle mt-0.5">
                              {rec.workbookTitle}
                            </div>
                          </div>
                        </div>
                      </td>
                      <td className="py-3.5 px-4 font-mono text-xs font-semibold uppercase text-burgundy">
                        .{rec.format}
                      </td>
                      <td className="py-3.5 px-4 text-xs text-ink-muted capitalize">
                        {rec.pageSize} {rec.orientation}
                      </td>
                      <td className="py-3.5 px-4 text-right font-mono text-xs text-ink tabular-nums">
                        {rec.entryCount}
                      </td>
                      <td className="py-3.5 px-4 text-xs text-ink-subtle tabular-nums">
                        {new Date(rec.exportedAt).toLocaleString()}
                      </td>
                      <td className="py-3.5 px-5">
                        <div className="flex items-center justify-end gap-1.5">
                          {stillExists && (
                            <button
                              type="button"
                              onClick={() => onOpenWorkbook(rec.workbookId, 'preview')}
                              className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium border border-hairline-strong bg-white text-ink hover:bg-ivory-deep cursor-pointer whitespace-nowrap"
                            >
                              <FolderOpen className="w-3.5 h-3.5" />
                              Open Book
                            </button>
                          )}
                          <button
                            type="button"
                            onClick={() => onDeleteExportRecord(rec.id)}
                            aria-label={`Remove ${rec.fileName} from history`}
                            className="p-1.5 text-ink-subtle hover:text-[#A51D24] cursor-pointer"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </section>
    </div>
  );
};
