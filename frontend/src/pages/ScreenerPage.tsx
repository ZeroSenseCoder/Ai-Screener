import { useState, useEffect } from 'react'
import { screenerApi, ScreenerRequest, ScreenerResponse } from '../services/api'
import SectorNav, { SECTORS } from '../components/screener/SectorNav'
import FilterPanel from '../components/screener/FilterPanel'
import StockTable from '../components/screener/StockTable'

const DEFAULT_FILTERS: ScreenerRequest = {
  exchanges: ['NSE', 'BSE'],
  sectors: [],
  sort_by: 'market_cap',
  sort_asc: false,
  page: 1,
  page_size: 50,
}

export default function ScreenerPage() {
  const [filters, setFilters] = useState<ScreenerRequest>(DEFAULT_FILTERS)
  const [result, setResult] = useState<ScreenerResponse | null>(null)
  const [loading, setLoading] = useState(false)
  const [filtersOpen, setFiltersOpen] = useState(false)

  useEffect(() => {
    setLoading(true)
    screenerApi.screen(filters)
      .then(setResult)
      .catch(() => {})
      .finally(() => setLoading(false))
  }, [filters])

  const activeSectors = filters.sectors || []

  // Resolve human-friendly sector label for the header
  const activeSectorLabels = SECTORS
    .filter(def => def.keys.some(k => activeSectors.includes(k)))
    .map(def => def.label)

  const resetFilters = () => setFilters(DEFAULT_FILTERS)

  return (
    <div className="flex h-full min-h-0">

      {/* ── LEFT SIDEBAR ─────────────────────────────────────────── */}
      <div className="w-56 flex-shrink-0 flex flex-col border-r border-brand-border bg-brand-card overflow-hidden">

        {/* Sector navigator — takes available space */}
        <div className="flex-1 min-h-0 overflow-y-auto">
          <SectorNav filters={filters} onChange={setFilters} />
        </div>

        {/* Filters accordion toggle at the bottom of sidebar */}
        <div className="border-t border-brand-border flex-shrink-0">
          <button
            className="w-full flex items-center justify-between px-3 py-2.5 text-xs text-slate-400 hover:text-slate-200 hover:bg-slate-800 transition-colors"
            onClick={() => setFiltersOpen(o => !o)}
          >
            <span className="font-semibold uppercase tracking-wider">Advanced Filters</span>
            <span>{filtersOpen ? '▴' : '▾'}</span>
          </button>
        </div>
      </div>

      {/* ── FILTER DRAWER (slides out beside sidebar) ────────────── */}
      {filtersOpen && (
        <div className="w-52 flex-shrink-0 border-r border-brand-border bg-brand-card overflow-y-auto">
          <FilterPanel
            filters={filters}
            onChange={setFilters}
            onReset={resetFilters}
          />
        </div>
      )}

      {/* ── MAIN CONTENT AREA ─────────────────────────────────────── */}
      <div className="flex-1 flex flex-col min-w-0 min-h-0">

        {/* Table header bar */}
        <div className="flex items-center gap-3 px-4 py-2.5 border-b border-brand-border flex-shrink-0 bg-brand-card">
          {/* Active sector breadcrumb */}
          <div className="flex items-center gap-2 min-w-0 flex-1">
            {activeSectorLabels.length === 0 ? (
              <h1 className="text-sm font-semibold text-slate-200">All Sectors</h1>
            ) : (
              <div className="flex items-center gap-2 flex-wrap">
                <span className="text-xs text-slate-500">Showing:</span>
                {activeSectorLabels.map(label => {
                  const def = SECTORS.find(s => s.label === label)!
                  return (
                    <span
                      key={label}
                      className={`flex items-center gap-1 text-xs px-2 py-0.5 rounded-full border ${def.bgActive} ${def.borderColor} ${def.textColor}`}
                    >
                      {def.icon} {label}
                      <button
                        className="ml-0.5 opacity-60 hover:opacity-100"
                        onClick={() => {
                          const next = activeSectors.filter(s => !def.keys.includes(s))
                          setFilters({ ...filters, sectors: next, page: 1 })
                        }}
                      >×</button>
                    </span>
                  )
                })}
                {activeSectorLabels.length > 0 && (
                  <button
                    className="text-xs text-slate-600 hover:text-red-400"
                    onClick={() => setFilters({ ...filters, sectors: [], page: 1 })}
                  >
                    Clear all
                  </button>
                )}
              </div>
            )}
          </div>

          <span className="text-xs text-slate-500 flex-shrink-0">
            {loading ? 'Loading...' : result ? `${result.total.toLocaleString()} stocks` : ''}
          </span>
        </div>

        {/* Stock table */}
        <StockTable
          stocks={result?.results || []}
          total={result?.total || 0}
          filters={filters}
          onFilterChange={setFilters}
          loading={loading}
        />
      </div>
    </div>
  )
}
