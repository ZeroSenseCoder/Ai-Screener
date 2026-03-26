import { useNavigate } from 'react-router-dom'
import { Stock, ScreenerRequest } from '../../services/api'
import { formatCurrency, formatMarketCap, formatPct, formatNum, formatVolume, pctColor } from '../../utils/formatters'

interface Props {
  stocks: Stock[]
  total: number
  filters: ScreenerRequest
  onFilterChange: (f: ScreenerRequest) => void
  loading: boolean
}

const SORT_COLS = [
  { key: 'symbol', label: 'Symbol' },
  { key: 'last_price', label: 'Price' },
  { key: 'daily_return', label: '1D %' },
  { key: 'return_1m', label: '1M %' },
  { key: 'return_3m', label: '3M %' },
  { key: 'return_1y', label: '1Y %' },
  { key: 'market_cap', label: 'MCap' },
  { key: 'pe_ratio', label: 'PE' },
  { key: 'beta', label: 'Beta' },
  { key: 'rsi_14', label: 'RSI' },
  { key: 'avg_volume_20d', label: 'Vol (20D)' },
  { key: 'max_drawdown_52w', label: 'MaxDD' },
]

export default function StockTable({ stocks, total, filters, onFilterChange, loading }: Props) {
  const navigate = useNavigate()
  const pageSize = filters.page_size || 50
  const page = filters.page || 1
  const pages = Math.ceil(total / pageSize)

  const sortBy = (col: string) => {
    const asc = filters.sort_by === col ? !filters.sort_asc : false
    onFilterChange({ ...filters, sort_by: col, sort_asc: asc, page: 1 })
  }

  const SortIcon = ({ col }: { col: string }) => {
    if (filters.sort_by !== col) return <span className="text-slate-600"> ⇅</span>
    return <span className="text-brand-accent">{filters.sort_asc ? ' ↑' : ' ↓'}</span>
  }

  return (
    <div className="flex flex-col flex-1 min-h-0">
      {/* Header bar */}
      <div className="flex items-center justify-between px-4 py-2 border-b border-brand-border text-xs text-slate-400">
        <span>
          {loading ? 'Loading...' : `${total.toLocaleString()} stocks found`}
        </span>
        <div className="flex items-center gap-2">
          <span>Rows:</span>
          {[50, 100, 200].map(n => (
            <button
              key={n}
              className={`px-2 py-0.5 rounded ${pageSize === n ? 'bg-brand-accent text-white' : 'btn-ghost'}`}
              onClick={() => onFilterChange({ ...filters, page_size: n, page: 1 })}
            >
              {n}
            </button>
          ))}
        </div>
      </div>

      {/* Table */}
      <div className="flex-1 overflow-auto">
        <table className="w-full text-xs border-collapse">
          <thead className="sticky top-0 bg-brand-card z-10">
            <tr>
              <th className="text-left px-3 py-2 text-slate-400 font-normal border-b border-brand-border">
                Exchange
              </th>
              {SORT_COLS.map(col => (
                <th
                  key={col.key}
                  className="text-right px-3 py-2 text-slate-400 font-normal border-b border-brand-border cursor-pointer hover:text-white whitespace-nowrap"
                  onClick={() => sortBy(col.key)}
                >
                  {col.label}
                  <SortIcon col={col.key} />
                </th>
              ))}
              <th className="text-left px-3 py-2 text-slate-400 font-normal border-b border-brand-border">Sector</th>
            </tr>
          </thead>
          <tbody>
            {loading && stocks.length === 0 ? (
              <tr>
                <td colSpan={SORT_COLS.length + 2} className="text-center py-20 text-slate-500">
                  Loading stocks...
                </td>
              </tr>
            ) : stocks.length === 0 ? (
              <tr>
                <td colSpan={SORT_COLS.length + 2} className="text-center py-20 text-slate-500">
                  No stocks match your filters.
                </td>
              </tr>
            ) : (
              stocks.map((stock, i) => (
                <tr
                  key={stock.yf_symbol || i}
                  className="border-b border-brand-border hover:bg-slate-800 cursor-pointer transition-colors"
                  onClick={() => navigate(`/stock/${stock.symbol}`)}
                >
                  <td className="px-3 py-2">
                    <span className={`text-xs px-1.5 py-0.5 rounded font-mono ${stock.exchange === 'NSE' ? 'bg-blue-900 text-blue-300' : 'bg-purple-900 text-purple-300'}`}>
                      {stock.exchange}
                    </span>
                  </td>
                  <td className="px-3 py-2 text-right font-semibold text-brand-accent">
                    {stock.symbol}
                    <div className="text-slate-500 text-xs font-normal truncate max-w-24">{stock.company_name}</div>
                  </td>
                  <td className={`px-3 py-2 text-right font-mono ${stock.last_price ? 'text-white' : 'text-slate-500'}`}>
                    {stock.last_price ? `₹${stock.last_price.toLocaleString('en-IN', { minimumFractionDigits: 2 })}` : '—'}
                  </td>
                  <td className={`px-3 py-2 text-right font-mono ${pctColor(stock.daily_return)}`}>
                    {formatPct(stock.daily_return)}
                  </td>
                  <td className={`px-3 py-2 text-right font-mono ${pctColor(stock.return_1m)}`}>
                    {formatPct(stock.return_1m)}
                  </td>
                  <td className={`px-3 py-2 text-right font-mono ${pctColor(stock.return_3m)}`}>
                    {formatPct(stock.return_3m)}
                  </td>
                  <td className={`px-3 py-2 text-right font-mono ${pctColor(stock.return_1y)}`}>
                    {formatPct(stock.return_1y)}
                  </td>
                  <td className="px-3 py-2 text-right font-mono">{formatMarketCap(stock.market_cap)}</td>
                  <td className="px-3 py-2 text-right font-mono">{formatNum(stock.pe_ratio)}</td>
                  <td className="px-3 py-2 text-right font-mono">{formatNum(stock.beta)}</td>
                  <td className={`px-3 py-2 text-right font-mono ${stock.rsi_14 ? (stock.rsi_14 > 70 ? 'text-red-400' : stock.rsi_14 < 30 ? 'text-green-400' : 'text-white') : 'text-slate-500'}`}>
                    {formatNum(stock.rsi_14)}
                  </td>
                  <td className="px-3 py-2 text-right font-mono">{formatVolume(stock.avg_volume_20d)}</td>
                  <td className={`px-3 py-2 text-right font-mono ${pctColor(stock.max_drawdown_52w)}`}>
                    {formatPct(stock.max_drawdown_52w)}
                  </td>
                  <td className="px-3 py-2 text-slate-400 text-xs max-w-28 truncate">{stock.sector}</td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {/* Pagination */}
      <div className="flex items-center justify-between px-4 py-2 border-t border-brand-border text-xs">
        <span className="text-slate-400">
          Page {page} of {pages} · {total.toLocaleString()} total
        </span>
        <div className="flex gap-1">
          <button
            className="btn-ghost px-2 py-1"
            disabled={page <= 1}
            onClick={() => onFilterChange({ ...filters, page: page - 1 })}
          >
            ← Prev
          </button>
          {Array.from({ length: Math.min(5, pages) }, (_, i) => {
            const p = Math.max(1, Math.min(page - 2, pages - 4)) + i
            return p <= pages ? (
              <button
                key={p}
                className={`px-2 py-1 rounded ${p === page ? 'bg-brand-accent text-white' : 'btn-ghost'}`}
                onClick={() => onFilterChange({ ...filters, page: p })}
              >
                {p}
              </button>
            ) : null
          })}
          <button
            className="btn-ghost px-2 py-1"
            disabled={page >= pages}
            onClick={() => onFilterChange({ ...filters, page: page + 1 })}
          >
            Next →
          </button>
        </div>
      </div>
    </div>
  )
}
