import { useMemo, useState } from 'react';
import { ChevronLeft, ChevronRight, Search } from 'lucide-react';

export default function DataTable({
  columns,
  data,
  searchPlaceholder = 'Rechercher...',
  onSearch,
  serverSearch = false,
  pageSize = 8,
  emptyMessage = 'Aucune donnée à afficher.',
  extraFilters,
}) {
  const [query, setQuery] = useState('');
  const [page, setPage] = useState(1);

  const filtered = useMemo(() => {
    if (serverSearch || !query) return data;
    const q = query.toLowerCase();
    return data.filter((row) =>
      columns.some((col) => {
        const value = col.searchValue ? col.searchValue(row) : row[col.key];
        return String(value ?? '').toLowerCase().includes(q);
      })
    );
  }, [data, query, columns, serverSearch]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
  const currentPage = Math.min(page, totalPages);
  const pageData = filtered.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  function handleSearchChange(value) {
    setQuery(value);
    setPage(1);
    if (serverSearch && onSearch) onSearch(value);
  }

  return (
    <div className="card overflow-hidden">
      <div className="flex flex-col gap-3 border-b border-slate-100 p-4 dark:border-navy-700 sm:flex-row sm:items-center sm:justify-between">
        <div className="relative w-full sm:max-w-xs">
          <Search size={16} className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input
            className="input pl-9"
            placeholder={searchPlaceholder}
            value={query}
            onChange={(e) => handleSearchChange(e.target.value)}
          />
        </div>
        {extraFilters && <div className="flex flex-wrap gap-2">{extraFilters}</div>}
      </div>

      <div className="overflow-x-auto">
        <table className="w-full text-left text-sm">
          <thead className="bg-slate-50 text-xs uppercase tracking-wide text-slate-500 dark:bg-navy-900 dark:text-slate-400">
            <tr>
              {columns.map((col) => (
                <th key={col.key} className="px-4 py-3 font-semibold">
                  {col.header}
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100 dark:divide-navy-700">
            {pageData.length === 0 && (
              <tr>
                <td colSpan={columns.length} className="px-4 py-10 text-center text-slate-400">
                  {emptyMessage}
                </td>
              </tr>
            )}
            {pageData.map((row, i) => (
              <tr key={row.id ?? i} className="transition-colors hover:bg-slate-50 dark:hover:bg-navy-900">
                {columns.map((col) => (
                  <td key={col.key} className="px-4 py-3 text-slate-700 dark:text-slate-200">
                    {col.render ? col.render(row) : row[col.key]}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {filtered.length > pageSize && (
        <div className="flex items-center justify-between border-t border-slate-100 px-4 py-3 text-sm text-slate-500 dark:border-navy-700">
          <span>
            Page {currentPage} sur {totalPages} — {filtered.length} résultat(s)
          </span>
          <div className="flex gap-2">
            <button
              className="btn-secondary px-2 py-1"
              disabled={currentPage <= 1}
              onClick={() => setPage((p) => Math.max(1, p - 1))}
            >
              <ChevronLeft size={16} />
            </button>
            <button
              className="btn-secondary px-2 py-1"
              disabled={currentPage >= totalPages}
              onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
            >
              <ChevronRight size={16} />
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
