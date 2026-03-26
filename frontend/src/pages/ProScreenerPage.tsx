import { useState, useEffect, useCallback, useRef } from 'react'
import { useNavigate } from 'react-router-dom'
import axios from 'axios'
import {
  proScreenerApi,
  FilterCondition,
  ProStockResult,
  CatalogCategory,
  CatalogFilter,
} from '../services/api'

// ── Category metadata (icons + colors) ────────────────────────────────────────
const CAT_META: Record<string, { icon: string; color: string }> = {
  A: { icon: '📊', color: 'text-blue-400' },
  C: { icon: '💰', color: 'text-yellow-400' },
  D: { icon: '📈', color: 'text-green-400' },
  E: { icon: '🚀', color: 'text-purple-400' },
  F: { icon: '⚖️', color: 'text-orange-400' },
  G: { icon: '💸', color: 'text-teal-400' },
  H: { icon: '🎯', color: 'text-pink-400' },
  I: { icon: '🏛️', color: 'text-indigo-400' },
  J: { icon: '🌱', color: 'text-emerald-400' },
  K: { icon: '📉', color: 'text-cyan-400' },
}

function fmtMktCap(v?: number | null) {
  if (!v) return '—'
  const cr = v / 1e7
  if (cr >= 1e5) return `₹${(cr / 1e5).toFixed(1)}L Cr`
  if (cr >= 1000) return `₹${(cr / 1000).toFixed(1)}K Cr`
  return `₹${cr.toFixed(0)} Cr`
}

function fmtNum(v?: number | null, digits = 2) {
  if (v == null) return '—'
  return v.toFixed(digits)
}

function pctColor(v?: number | null) {
  if (v == null) return 'text-slate-400'
  return v > 0 ? 'text-up' : v < 0 ? 'text-down' : 'text-slate-400'
}

// ── Normalize weights so they sum to exactly 1 ────────────────────────────────
function normalizeWeights(conds: FilterCondition[]): FilterCondition[] {
  if (conds.length === 0) return conds
  const total = conds.reduce((s, c) => s + (c.weight || 0), 0)
  if (total === 0) {
    // Distribute equally
    const eq = parseFloat((1 / conds.length).toFixed(4))
    return conds.map((c, i) => ({
      ...c,
      weight: i === conds.length - 1
        ? parseFloat((1 - eq * (conds.length - 1)).toFixed(4))
        : eq,
    }))
  }
  // Scale so sum = 1
  const scaled = conds.map(c => ({ ...c, weight: (c.weight || 0) / total }))
  // Fix floating-point drift: add residual to last condition
  const sumScaled = scaled.reduce((s, c) => s + c.weight, 0)
  const residual = parseFloat((1 - sumScaled).toFixed(10))
  scaled[scaled.length - 1].weight = parseFloat(
    (scaled[scaled.length - 1].weight + residual).toFixed(4)
  )
  return scaled
}

// ── Active Filter Card ─────────────────────────────────────────────────────────
function ActiveFilterCard({
  cond,
  filter,
  onChange,
  onRemove,
}: {
  cond: FilterCondition
  filter: CatalogFilter
  onChange: (updated: FilterCondition) => void
  onRemove: () => void
}) {
  const isBool = filter.type === 'bool'
  const isUnavail = !filter.available

  return (
    <div className={`relative rounded-lg border p-3 ${isUnavail ? 'border-slate-700 opacity-60' : 'border-brand-border bg-slate-800/60'}`}>
      <div className="flex items-start justify-between gap-2">
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-2 mb-2">
            <span className="text-xs font-medium text-white truncate">{filter.label}</span>
            {isUnavail && (
              <span className="text-[10px] text-orange-400 bg-orange-900/40 px-1 rounded shrink-0">Limited data</span>
            )}
            {isBool && (
              <span className="text-[10px] text-blue-400 bg-blue-900/40 px-1 rounded shrink-0">Signal</span>
            )}
          </div>

          {!isBool && (
            <div className="flex items-center gap-1 mb-2">
              <input
                type="number"
                placeholder="Min"
                value={cond.min_val ?? ''}
                onChange={e => onChange({ ...cond, min_val: e.target.value === '' ? null : +e.target.value })}
                className="w-20 text-xs bg-slate-700 border border-slate-600 rounded px-2 py-1 text-white placeholder-slate-500 focus:outline-none focus:border-blue-500"
              />
              <span className="text-slate-500 text-xs">–</span>
              <input
                type="number"
                placeholder="Max"
                value={cond.max_val ?? ''}
                onChange={e => onChange({ ...cond, max_val: e.target.value === '' ? null : +e.target.value })}
                className="w-20 text-xs bg-slate-700 border border-slate-600 rounded px-2 py-1 text-white placeholder-slate-500 focus:outline-none focus:border-blue-500"
              />
              {filter.unit && <span className="text-slate-500 text-xs">{filter.unit}</span>}
            </div>
          )}

          <div className="flex items-center justify-between gap-2">
            <div className="flex items-center gap-1.5">
              <span className="text-[10px] text-slate-500 shrink-0">Weight:</span>
              <input
                type="number"
                min={0}
                max={1}
                step={0.01}
                value={cond.weight.toFixed(2)}
                onChange={e => onChange({ ...cond, weight: Math.min(1, Math.max(0, parseFloat(e.target.value) || 0)) })}
                className="w-16 text-xs bg-slate-700 border border-slate-600 rounded px-2 py-0.5 text-yellow-300 font-mono focus:outline-none focus:border-yellow-500"
              />
              <span className="text-[10px] text-yellow-500 font-mono">
                {(cond.weight * 100).toFixed(0)}%
              </span>
            </div>
            <button
              onClick={() => onChange({ ...cond, required: !cond.required })}
              className={`text-[10px] px-2 py-0.5 rounded border transition-colors shrink-0 ${
                cond.required
                  ? 'border-red-600 text-red-400 bg-red-900/20'
                  : 'border-slate-600 text-slate-400 hover:border-slate-500'
              }`}
              title={cond.required ? 'Hard filter: must pass' : 'Soft filter: contributes to score'}
            >
              {cond.required ? 'Required' : 'Score only'}
            </button>
          </div>
        </div>

        <button
          onClick={onRemove}
          className="text-slate-500 hover:text-red-400 text-sm shrink-0 mt-0.5"
        >
          ✕
        </button>
      </div>
    </div>
  )
}

// ── Catalog Row ────────────────────────────────────────────────────────────────
function CatalogRow({
  filter,
  active,
  onAdd,
}: {
  filter: CatalogFilter
  active: boolean
  onAdd: () => void
}) {
  return (
    <div
      className={`flex items-center justify-between px-3 py-1.5 rounded group hover:bg-slate-700/50 ${
        !filter.available ? 'opacity-50' : ''
      }`}
    >
      <div className="flex items-center gap-2 min-w-0">
        <span className="text-xs text-slate-300 truncate">{filter.label}</span>
        {filter.unit && <span className="text-[10px] text-slate-600">{filter.unit}</span>}
        {!filter.available && (
          <span className="text-[10px] text-orange-500 shrink-0">⚠</span>
        )}
      </div>
      <button
        onClick={onAdd}
        disabled={active}
        className={`shrink-0 text-[10px] px-2 py-0.5 rounded border transition-all ml-2 ${
          active
            ? 'border-green-700 text-green-500 cursor-default'
            : 'border-slate-600 text-slate-400 hover:border-blue-500 hover:text-blue-400 cursor-pointer'
        }`}
      >
        {active ? '✓' : '+'}
      </button>
    </div>
  )
}

// ── Results Table ──────────────────────────────────────────────────────────────
function ResultsTable({
  results,
  loading,
  total,
  page,
  onPage,
}: {
  results: ProStockResult[]
  loading: boolean
  total: number
  page: number
  onPage: (p: number) => void
}) {
  const navigate = useNavigate()
  const totalPages = Math.ceil(total / 50)

  const columns = [
    { key: 'rank', label: '#', width: 'w-8' },
    { key: 'symbol', label: 'Symbol', width: 'w-24' },
    { key: 'company_name', label: 'Company', width: 'min-w-32' },
    { key: 'sector', label: 'Sector', width: 'w-28' },
    { key: 'score', label: 'Score', width: 'w-20' },
    { key: 'last_price', label: 'Price', width: 'w-20' },
    { key: 'market_cap', label: 'Mkt Cap', width: 'w-24' },
    { key: 'pe_ratio', label: 'P/E', width: 'w-16' },
    { key: 'beta', label: 'Beta', width: 'w-14' },
    { key: 'rsi_14', label: 'RSI', width: 'w-14' },
    { key: 'return_1m', label: '1M%', width: 'w-16' },
    { key: 'return_3m', label: '3M%', width: 'w-16' },
    { key: 'return_1y', label: '1Y%', width: 'w-16' },
    { key: 'max_drawdown_52w', label: 'DD%', width: 'w-16' },
    { key: 'avg_volume_20d', label: 'Vol 20D', width: 'w-20' },
  ]

  return (
    <div className="flex flex-col min-h-0">
      <div className="flex items-center justify-between px-4 py-2 border-b border-brand-border">
        <span className="text-xs text-slate-400">
          {loading ? 'Screening…' : `${total.toLocaleString()} stocks matched`}
        </span>
        {totalPages > 1 && (
          <div className="flex items-center gap-2">
            <button
              disabled={page === 1}
              onClick={() => onPage(page - 1)}
              className="text-xs px-2 py-1 rounded border border-slate-600 text-slate-400 disabled:opacity-40 hover:border-slate-400"
            >
              ‹
            </button>
            <span className="text-xs text-slate-500">{page} / {totalPages}</span>
            <button
              disabled={page === totalPages}
              onClick={() => onPage(page + 1)}
              className="text-xs px-2 py-1 rounded border border-slate-600 text-slate-400 disabled:opacity-40 hover:border-slate-400"
            >
              ›
            </button>
          </div>
        )}
      </div>

      <div className="flex-1 overflow-auto">
        <table className="w-full text-xs">
          <thead className="sticky top-0 bg-brand-card border-b border-brand-border z-10">
            <tr>
              {columns.map(c => (
                <th key={c.key} className={`px-2 py-2 text-left text-slate-400 font-medium ${c.width} whitespace-nowrap`}>
                  {c.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {loading && (
              <tr>
                <td colSpan={columns.length} className="px-4 py-8 text-center text-slate-500">
                  <div className="inline-block w-6 h-6 border-2 border-blue-500 border-t-transparent rounded-full animate-spin" />
                </td>
              </tr>
            )}
            {!loading && results.length === 0 && (
              <tr>
                <td colSpan={columns.length} className="px-4 py-8 text-center text-slate-500">
                  No stocks match the selected conditions
                </td>
              </tr>
            )}
            {!loading && results.map((r, i) => {
              const globalRank = (page - 1) * 50 + i + 1
              return (
                <tr
                  key={r.symbol}
                  onClick={() => navigate(`/stock/${r.symbol}`)}
                  className="border-b border-slate-800/50 hover:bg-slate-800/60 cursor-pointer transition-colors"
                >
                  <td className="px-2 py-2 text-slate-500 w-8">{globalRank}</td>
                  <td className="px-2 py-2 font-semibold text-brand-accent">{r.symbol}</td>
                  <td className="px-2 py-2 text-slate-300 max-w-32 truncate">{r.company_name}</td>
                  <td className="px-2 py-2 text-slate-400 w-28 truncate">{r.sector}</td>
                  <td className="px-2 py-2 w-20">
                    <div className="flex items-center gap-1">
                      <div className="flex-1 bg-slate-700 rounded-full h-1.5 w-12">
                        <div
                          className="h-1.5 rounded-full bg-blue-500 transition-all"
                          style={{ width: `${r.score}%` }}
                        />
                      </div>
                      <span className="text-blue-400 font-medium text-[11px]">{fmtNum(r.score, 0)}%</span>
                    </div>
                  </td>
                  <td className="px-2 py-2 w-20 font-medium">{r.last_price != null ? `₹${r.last_price.toLocaleString()}` : '—'}</td>
                  <td className="px-2 py-2 w-24 text-slate-400">{fmtMktCap(r.market_cap)}</td>
                  <td className="px-2 py-2 w-16 text-slate-300">{fmtNum(r.pe_ratio, 1)}</td>
                  <td className="px-2 py-2 w-14 text-slate-300">{fmtNum(r.beta, 2)}</td>
                  <td className={`px-2 py-2 w-14 ${r.rsi_14 != null && r.rsi_14 < 30 ? 'text-green-400' : r.rsi_14 != null && r.rsi_14 > 70 ? 'text-red-400' : 'text-slate-300'}`}>
                    {fmtNum(r.rsi_14, 1)}
                  </td>
                  <td className={`px-2 py-2 w-16 ${pctColor(r.return_1m)}`}>{r.return_1m != null ? `${r.return_1m > 0 ? '+' : ''}${fmtNum(r.return_1m, 1)}%` : '—'}</td>
                  <td className={`px-2 py-2 w-16 ${pctColor(r.return_3m)}`}>{r.return_3m != null ? `${r.return_3m > 0 ? '+' : ''}${fmtNum(r.return_3m, 1)}%` : '—'}</td>
                  <td className={`px-2 py-2 w-16 ${pctColor(r.return_1y)}`}>{r.return_1y != null ? `${r.return_1y > 0 ? '+' : ''}${fmtNum(r.return_1y, 1)}%` : '—'}</td>
                  <td className={`px-2 py-2 w-16 ${r.max_drawdown_52w != null ? 'text-red-400' : 'text-slate-400'}`}>{r.max_drawdown_52w != null ? `${fmtNum(r.max_drawdown_52w, 1)}%` : '—'}</td>
                  <td className="px-2 py-2 w-20 text-slate-400">{r.avg_volume_20d != null ? `${(r.avg_volume_20d / 1e5).toFixed(1)}L` : '—'}</td>
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>
    </div>
  )
}

// ── Enrichment Banner ─────────────────────────────────────────────────────────
function EnrichmentBanner({ status, progress }: { status: string; progress: number }) {
  if (status === 'done') return null
  return (
    <div className={`px-4 py-2 text-xs flex items-center gap-3 ${
      status === 'running' ? 'bg-blue-950/60 border-b border-blue-800 text-blue-300' : 'bg-slate-800 text-slate-500'
    }`}>
      {status === 'running' ? (
        <>
          <div className="w-3 h-3 border-2 border-blue-400 border-t-transparent rounded-full animate-spin shrink-0" />
          <span>Enriching indicators & fundamentals in background…</span>
          <div className="flex-1 bg-blue-900/40 rounded-full h-1.5 max-w-32">
            <div className="h-1.5 rounded-full bg-blue-500 transition-all" style={{ width: `${progress}%` }} />
          </div>
          <span className="text-blue-400">{progress}%</span>
          <span className="text-blue-600">Filters with ⚠ require enriched data</span>
        </>
      ) : (
        <span>Loading indicators…</span>
      )}
    </div>
  )
}

// ── Main Page ──────────────────────────────────────────────────────────────────
export default function ProScreenerPage() {
  const [catalog, setCatalog] = useState<CatalogCategory[]>([])
  const [sectors, setSectors] = useState<{ sector: string; count: number }[]>([])
  const [selectedSector, setSelectedSector] = useState<string | null>(null)
  const [conditions, setConditions] = useState<FilterCondition[]>([])
  const [expandedCats, setExpandedCats] = useState<Set<string>>(new Set(['A', 'C', 'K']))
  const [scoreMode, setScoreMode] = useState(false)
  const [results, setResults] = useState<ProStockResult[]>([])
  const [total, setTotal] = useState(0)
  const [page, setPage] = useState(1)
  const [loading, setLoading] = useState(false)
  const [hasRun, setHasRun] = useState(false)
  const [enrichStatus, setEnrichStatus] = useState('pending')
  const [enrichProgress, setEnrichProgress] = useState(0)
  const pollRef = useRef<ReturnType<typeof setInterval> | null>(null)

  // Poll enrichment status until done
  useEffect(() => {
    const poll = () => {
      axios.get('/api/v1/health').then(r => {
        setEnrichStatus(r.data.enrichment_status)
        setEnrichProgress(r.data.enrichment_progress)
        if (r.data.enrichment_status === 'done') {
          clearInterval(pollRef.current!)
        }
      }).catch(() => {})
    }
    poll()
    pollRef.current = setInterval(poll, 5000)
    return () => clearInterval(pollRef.current!)
  }, [])

  // Load catalog + sectors on mount
  useEffect(() => {
    proScreenerApi.catalog().then(setCatalog).catch(() => {})
    proScreenerApi.sectors().then(setSectors).catch(() => {})
  }, [])

  // Map filter_id → CatalogFilter
  const filterById = useCallback((): Record<string, CatalogFilter> => {
    const map: Record<string, CatalogFilter> = {}
    for (const cat of catalog) {
      for (const f of cat.filters) map[f.id] = f
    }
    return map
  }, [catalog])

  const activeIds = new Set(conditions.map(c => c.filter_id))

  const addCondition = (filter: CatalogFilter) => {
    if (activeIds.has(filter.id)) return
    setConditions(prev => normalizeWeights([
      ...prev,
      { filter_id: filter.id, min_val: null, max_val: null, weight: 0, required: true },
    ]))
  }

  const updateCondition = (idx: number, updated: FilterCondition) => {
    // When user manually edits a weight, redistribute the remainder to others
    setConditions(prev => {
      const next = prev.map((c, i) => (i === idx ? updated : c))
      // Only auto-redistribute if user changed weight (not other fields)
      if (updated.weight !== prev[idx].weight) {
        const newW = Math.min(1, Math.max(0, updated.weight))
        const remainder = parseFloat((1 - newW).toFixed(4))
        const others = next.filter((_, i) => i !== idx)
        const othersTotal = others.reduce((s, c) => s + c.weight, 0)
        const redistributed = others.map(c => ({
          ...c,
          weight: othersTotal === 0
            ? parseFloat((remainder / others.length).toFixed(4))
            : parseFloat(((c.weight / othersTotal) * remainder).toFixed(4)),
        }))
        // Fix residual drift
        const sumR = redistributed.reduce((s, c) => s + c.weight, 0)
        if (redistributed.length > 0) {
          redistributed[redistributed.length - 1].weight = parseFloat(
            (redistributed[redistributed.length - 1].weight + parseFloat((remainder - sumR).toFixed(10))).toFixed(4)
          )
        }
        const result: FilterCondition[] = []
        let ri = 0
        for (let i = 0; i < next.length; i++) {
          if (i === idx) result.push({ ...updated, weight: newW })
          else result.push(redistributed[ri++])
        }
        return result
      }
      return next
    })
  }

  const removeCondition = (idx: number) => {
    setConditions(prev => normalizeWeights(prev.filter((_, i) => i !== idx)))
  }

  const runScreen = async (p = 1) => {
    setLoading(true)
    setHasRun(true)
    try {
      const resp = await proScreenerApi.screen({
        sector: selectedSector,
        conditions,
        score_mode: scoreMode,
        sort_by: 'score',
        sort_asc: false,
        page: p,
        page_size: 50,
      })
      setResults(resp.results)
      setTotal(resp.total)
      setPage(p)
    } catch (e) {
      console.error(e)
    } finally {
      setLoading(false)
    }
  }

  const totalWeight = parseFloat(conditions.reduce((s, c) => s + c.weight, 0).toFixed(4))
  const weightOk = conditions.length === 0 || Math.abs(totalWeight - 1) < 0.005
  const fMap = filterById()

  const toggleCat = (cat: string) => {
    setExpandedCats(prev => {
      const next = new Set(prev)
      next.has(cat) ? next.delete(cat) : next.add(cat)
      return next
    })
  }

  return (
    <div className="flex h-full overflow-hidden bg-brand text-white text-sm">
      {/* ── Left: Filter Catalog ────────────────────────────────────────────── */}
      <aside className="w-56 shrink-0 flex flex-col border-r border-brand-border overflow-y-auto bg-slate-900">
        <div className="px-3 py-3 border-b border-brand-border">
          <div className="text-xs font-bold text-slate-300 uppercase tracking-wider">Filter Catalog</div>
          <div className="text-[10px] text-slate-500 mt-0.5">Click + to add to screen</div>
        </div>

        {catalog.map(cat => {
          const meta = CAT_META[cat.category] || { icon: '•', color: 'text-slate-400' }
          const isOpen = expandedCats.has(cat.category)
          return (
            <div key={cat.category} className="border-b border-slate-800">
              <button
                onClick={() => toggleCat(cat.category)}
                className="w-full flex items-center gap-2 px-3 py-2 hover:bg-slate-800 transition-colors text-left"
              >
                <span className="text-sm">{meta.icon}</span>
                <span className={`text-xs font-medium ${meta.color}`}>{cat.label}</span>
                <span className="ml-auto text-slate-600 text-xs">{isOpen ? '▲' : '▼'}</span>
              </button>
              {isOpen && (
                <div className="py-1">
                  {cat.filters.map(f => (
                    <CatalogRow
                      key={f.id}
                      filter={f}
                      active={activeIds.has(f.id)}
                      onAdd={() => addCondition(f)}
                    />
                  ))}
                </div>
              )}
            </div>
          )
        })}
      </aside>

      {/* ── Center: Conditions + Controls ───────────────────────────────────── */}
      <aside className="w-72 shrink-0 flex flex-col border-r border-brand-border bg-slate-900/60">
        {/* Sector Picker */}
        <div className="px-3 py-3 border-b border-brand-border">
          <div className="text-[10px] text-slate-500 uppercase tracking-wider mb-2">Sector Scope</div>
          <div className="flex flex-wrap gap-1">
            <button
              onClick={() => setSelectedSector(null)}
              className={`text-[10px] px-2 py-1 rounded border transition-colors ${
                !selectedSector
                  ? 'bg-blue-600 border-blue-500 text-white'
                  : 'border-slate-600 text-slate-400 hover:border-slate-400'
              }`}
            >
              All
            </button>
            {sectors.map(s => (
              <button
                key={s.sector}
                onClick={() => setSelectedSector(s.sector === selectedSector ? null : s.sector)}
                className={`text-[10px] px-2 py-1 rounded border transition-colors ${
                  selectedSector === s.sector
                    ? 'bg-blue-600 border-blue-500 text-white'
                    : 'border-slate-600 text-slate-400 hover:border-slate-400'
                }`}
              >
                {s.sector} <span className="text-slate-500">{s.count}</span>
              </button>
            ))}
          </div>
        </div>

        {/* Active Conditions */}
        <div className="px-3 py-2 border-b border-brand-border flex items-center justify-between">
          <div className="text-[10px] text-slate-500 uppercase tracking-wider">
            Active Conditions ({conditions.length})
          </div>
          {conditions.length > 0 && (
            <button
              onClick={() => setConditions([])}
              className="text-[10px] text-red-500 hover:text-red-400"
            >
              Clear all
            </button>
          )}
        </div>

        <div className="flex-1 overflow-y-auto p-2 space-y-2">
          {conditions.length === 0 && (
            <div className="text-center text-slate-600 text-xs py-8">
              Add filters from the catalog →
            </div>
          )}
          {conditions.map((cond, i) => {
            const f = fMap[cond.filter_id]
            if (!f) return null
            return (
              <ActiveFilterCard
                key={`${cond.filter_id}-${i}`}
                cond={cond}
                filter={f}
                onChange={updated => updateCondition(i, updated)}
                onRemove={() => removeCondition(i)}
              />
            )
          })}
        </div>

        {/* Weight Summary + Run */}
        <div className="border-t border-brand-border p-3 space-y-3">
          {conditions.length > 0 && (
            <div>
              <div className="flex justify-between text-[10px] mb-1">
                <span className="text-slate-500">Weight sum</span>
                <span className={`font-mono font-semibold ${weightOk ? 'text-green-400' : 'text-red-400'}`}>
                  {totalWeight.toFixed(4)} {weightOk ? '✓' : '≠ 1'}
                </span>
              </div>
              <div className="bg-slate-700 rounded-full h-1.5">
                <div
                  className={`h-1.5 rounded-full transition-all ${weightOk ? 'bg-green-500' : 'bg-red-500'}`}
                  style={{ width: `${Math.min(100, totalWeight * 100)}%` }}
                />
              </div>
              {!weightOk && (
                <button
                  onClick={() => setConditions(normalizeWeights(conditions))}
                  className="mt-1 text-[10px] text-yellow-400 hover:text-yellow-300 underline"
                >
                  Auto-normalize to 1.0
                </button>
              )}
            </div>
          )}

          {/* Score mode toggle */}
          <div className="flex items-center justify-between">
            <div>
              <div className="text-[10px] text-slate-400 font-medium">Score Mode</div>
              <div className="text-[9px] text-slate-600">
                {scoreMode ? 'Rank all stocks by score' : 'Hard filter: must pass all'}
              </div>
            </div>
            <button
              onClick={() => setScoreMode(!scoreMode)}
              className={`w-9 h-5 rounded-full transition-colors ${scoreMode ? 'bg-blue-600' : 'bg-slate-600'}`}
            >
              <div className={`w-4 h-4 bg-white rounded-full m-0.5 transition-transform ${scoreMode ? 'translate-x-4' : ''}`} />
            </button>
          </div>

          <button
            onClick={() => runScreen(1)}
            disabled={loading || conditions.length === 0}
            className="w-full py-2 rounded bg-blue-600 hover:bg-blue-500 disabled:opacity-40 disabled:cursor-not-allowed text-white text-sm font-semibold transition-colors"
          >
            {loading ? 'Screening…' : 'Run Screen'}
          </button>

          {hasRun && !loading && (
            <div className="text-center text-[10px] text-slate-500">
              {total.toLocaleString()} stocks found
            </div>
          )}
        </div>
      </aside>

      {/* ── Right: Results ───────────────────────────────────────────────────── */}
      <main className="flex-1 flex flex-col min-w-0 overflow-hidden">
        <EnrichmentBanner status={enrichStatus} progress={enrichProgress} />
        <div className="px-4 py-3 border-b border-brand-border flex items-center gap-3">
          <div>
            <span className="font-bold text-white text-sm">Pro Screener</span>
            <span className="ml-2 text-xs text-slate-500">Bloomberg-style weighted filter</span>
          </div>
          <div className="ml-auto flex items-center gap-3 text-xs text-slate-500">
            {selectedSector && (
              <div className="flex items-center gap-1 bg-blue-900/30 border border-blue-800 text-blue-300 px-2 py-1 rounded text-[11px]">
                {selectedSector}
                <button onClick={() => setSelectedSector(null)} className="hover:text-white ml-1">✕</button>
              </div>
            )}
            <span>{conditions.length} conditions active</span>
            {scoreMode && <span className="text-blue-400">Score mode ON</span>}
          </div>
        </div>

        {!hasRun ? (
          <div className="flex-1 flex flex-col items-center justify-center text-slate-600">
            <div className="text-4xl mb-3">📊</div>
            <div className="text-sm font-medium text-slate-500">Add filters and click Run Screen</div>
            <div className="text-xs text-slate-600 mt-1">
              Choose from {catalog.reduce((s, c) => s + c.filters.length, 0)} filter criteria across {catalog.length} categories
            </div>

            {/* Quick presets */}
            <div className="mt-6 flex flex-col items-center gap-2">
              <div className="text-xs text-slate-600 mb-1">Quick presets</div>
              <div className="flex flex-wrap gap-2 justify-center max-w-lg">
                {[
                  {
                    label: 'Value Screen', desc: 'Low P/E, oversold RSI',
                    conds: [
                      { filter_id: 'pe_ratio',   min_val: 5,   max_val: 20, weight: 0.40, required: true },
                      { filter_id: 'rsi',         max_val: 45,              weight: 0.35, required: true },
                      { filter_id: 'above_sma200',                           weight: 0.25, required: false },
                    ]
                  },
                  {
                    label: 'Momentum Screen', desc: 'Strong 1M/3M returns',
                    conds: [
                      { filter_id: 'return_1m',  min_val: 5,               weight: 0.35, required: true },
                      { filter_id: 'return_3m',  min_val: 10,              weight: 0.30, required: true },
                      { filter_id: 'above_sma50',                           weight: 0.20, required: true },
                      { filter_id: 'macd_bullish',                          weight: 0.15, required: false },
                    ]
                  },
                  {
                    label: 'Quality Growth', desc: 'Large-cap, rising trend',
                    conds: [
                      { filter_id: 'market_cap', min_val: 10000,            weight: 0.25, required: true },
                      { filter_id: 'golden_cross',                          weight: 0.35, required: true },
                      { filter_id: 'rsi',        min_val: 40,  max_val: 65, weight: 0.20, required: true },
                      { filter_id: 'above_ema200',                          weight: 0.20, required: false },
                    ]
                  },
                  {
                    label: 'Oversold Reversal', desc: 'Deeply oversold setups',
                    conds: [
                      { filter_id: 'rsi_oversold',                          weight: 0.45, required: true },
                      { filter_id: 'drawdown',   min_val: -40,              weight: 0.35, required: true },
                      { filter_id: 'volume_20d', min_val: 100,              weight: 0.20, required: true },
                    ]
                  },
                ].map(preset => (
                  <button
                    key={preset.label}
                    onClick={() => {
                      setConditions(normalizeWeights(preset.conds.map(c => ({
                        filter_id: c.filter_id,
                        min_val: (c as any).min_val ?? null,
                        max_val: (c as any).max_val ?? null,
                        weight: c.weight,
                        required: c.required,
                      }))))
                    }}
                    className="px-3 py-2 rounded-lg border border-slate-700 hover:border-blue-600 text-left transition-colors bg-slate-800/60 hover:bg-slate-800"
                  >
                    <div className="text-xs font-medium text-slate-300">{preset.label}</div>
                    <div className="text-[10px] text-slate-600">{preset.desc}</div>
                  </button>
                ))}
              </div>
            </div>
          </div>
        ) : (
          <ResultsTable
            results={results}
            loading={loading}
            total={total}
            page={page}
            onPage={p => runScreen(p)}
          />
        )}
      </main>
    </div>
  )
}
