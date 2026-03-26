import { useState } from 'react'
import { useNavigate, useLocation } from 'react-router-dom'
import { stockApi } from '../../services/api'

export default function Navbar() {
  const [query, setQuery] = useState('')
  const [results, setResults] = useState<any[]>([])
  const navigate = useNavigate()
  const location = useLocation()

  const onSearch = async (q: string) => {
    setQuery(q)
    if (q.length < 1) { setResults([]); return }
    const data = await stockApi.search(q).catch(() => [])
    setResults(data)
  }

  const onSelect = (symbol: string) => {
    setQuery('')
    setResults([])
    navigate(`/stock/${symbol}`)
  }

  const navLink = (path: string, label: string) => (
    <button
      onClick={() => navigate(path)}
      className={`text-sm px-3 py-1 rounded transition-colors ${
        location.pathname === path
          ? 'bg-blue-600 text-white'
          : 'text-slate-400 hover:text-white hover:bg-slate-700'
      }`}
    >
      {label}
    </button>
  )

  return (
    <nav className="flex items-center gap-4 px-4 py-3 bg-brand-card border-b border-brand-border">
      <div
        className="text-brand-accent font-bold text-lg cursor-pointer"
        onClick={() => navigate('/')}
      >
        IndiaScreen
      </div>
      <div className="text-xs text-slate-500">NSE + BSE</div>

      {/* Nav links */}
      <div className="flex items-center gap-1">
        {navLink('/', 'Screener')}
        {navLink('/pro-screener', 'Pro Screener')}
      </div>

      {/* Search */}
      <div className="relative ml-2 flex-1 max-w-sm">
        <input
          type="text"
          placeholder="Search symbol or company..."
          className="filter-input pr-8"
          value={query}
          onChange={e => onSearch(e.target.value)}
        />
        {results.length > 0 && (
          <div className="absolute top-full left-0 right-0 mt-1 bg-brand-card border border-brand-border rounded z-50 shadow-xl">
            {results.map(r => (
              <div
                key={r.yf_symbol}
                className="flex items-center justify-between px-3 py-2 hover:bg-slate-700 cursor-pointer"
                onClick={() => onSelect(r.symbol)}
              >
                <div>
                  <span className="font-semibold text-brand-accent">{r.symbol}</span>
                  <span className="text-slate-400 text-xs ml-2">{r.company_name}</span>
                </div>
                <span className="text-xs text-slate-500">{r.exchange}</span>
              </div>
            ))}
          </div>
        )}
      </div>

      <div className="ml-auto text-xs text-slate-500">
        {new Date().toLocaleDateString('en-IN', { weekday: 'short', day: 'numeric', month: 'short' })}
      </div>
    </nav>
  )
}
