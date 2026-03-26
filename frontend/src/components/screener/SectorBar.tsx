/**
 * Horizontal scrollable sector tile bar shown above the stock table.
 * One-click to filter by sector; click again to deselect.
 */
import { useEffect, useState } from 'react'
import { screenerApi } from '../../services/api'
import { ScreenerRequest } from '../../services/api'

// Sector → emoji icon + accent colour (Tailwind bg / text classes)
const SECTOR_META: Record<string, { icon: string; bg: string; border: string; text: string }> = {
  'Financial Services': { icon: '🏦', bg: 'bg-blue-950',   border: 'border-blue-700',   text: 'text-blue-300'   },
  'Information Technology': { icon: '💻', bg: 'bg-indigo-950', border: 'border-indigo-600', text: 'text-indigo-300' },
  'Healthcare':          { icon: '💊', bg: 'bg-green-950',  border: 'border-green-700',  text: 'text-green-300'  },
  'Energy':              { icon: '⚡', bg: 'bg-yellow-950', border: 'border-yellow-700', text: 'text-yellow-300' },
  'Consumer Staples':    { icon: '🛒', bg: 'bg-orange-950', border: 'border-orange-700', text: 'text-orange-300' },
  'Consumer Discretionary': { icon: '🛍️', bg: 'bg-pink-950', border: 'border-pink-700',  text: 'text-pink-300'  },
  'Automobile':          { icon: '🚗', bg: 'bg-cyan-950',   border: 'border-cyan-700',   text: 'text-cyan-300'   },
  'Capital Goods':       { icon: '⚙️', bg: 'bg-slate-800',  border: 'border-slate-600',  text: 'text-slate-300'  },
  'Metals & Mining':     { icon: '⛏️', bg: 'bg-zinc-900',   border: 'border-zinc-600',   text: 'text-zinc-300'   },
  'Chemicals':           { icon: '🧪', bg: 'bg-teal-950',   border: 'border-teal-700',   text: 'text-teal-300'   },
  'Real Estate':         { icon: '🏗️', bg: 'bg-amber-950',  border: 'border-amber-700',  text: 'text-amber-300'  },
  'Infrastructure':      { icon: '🏛️', bg: 'bg-stone-900',  border: 'border-stone-600',  text: 'text-stone-300'  },
  'Cement':              { icon: '🧱', bg: 'bg-neutral-900',border: 'border-neutral-600',text: 'text-neutral-300'},
  'Textiles':            { icon: '🧵', bg: 'bg-purple-950', border: 'border-purple-700', text: 'text-purple-300' },
  'Telecom':             { icon: '📡', bg: 'bg-sky-950',    border: 'border-sky-700',    text: 'text-sky-300'    },
  'Diversified':         { icon: '📊', bg: 'bg-slate-900',  border: 'border-slate-600',  text: 'text-slate-400'  },
  'Unknown':             { icon: '❓', bg: 'bg-slate-900',  border: 'border-slate-700',  text: 'text-slate-500'  },
}

const DEFAULT_META = { icon: '📌', bg: 'bg-slate-900', border: 'border-slate-700', text: 'text-slate-400' }

interface SectorInfo { sector: string; count: number }

interface Props {
  filters: ScreenerRequest
  onChange: (f: ScreenerRequest) => void
}

export default function SectorBar({ filters, onChange }: Props) {
  const [sectors, setSectors] = useState<SectorInfo[]>([])

  useEffect(() => {
    screenerApi.sectorSummary().then((data: any[]) => {
      const sorted = [...data]
        .filter(s => s.sector !== 'Unknown')
        .sort((a, b) => b.count - a.count)
      // Append Unknown at the end
      const unknown = data.find(s => s.sector === 'Unknown')
      if (unknown) sorted.push(unknown)
      setSectors(sorted)
    }).catch(() => {})
  }, [])

  const selected = filters.sectors || []

  const toggle = (sector: string) => {
    const next = selected.includes(sector)
      ? selected.filter(s => s !== sector)
      : [...selected, sector]
    onChange({ ...filters, sectors: next, page: 1 })
  }

  const clearAll = () => onChange({ ...filters, sectors: [], page: 1 })

  if (sectors.length === 0) return null

  return (
    <div className="border-b border-brand-border bg-brand">
      <div className="flex items-center gap-2 px-4 py-2 overflow-x-auto scrollbar-thin">
        {/* All button */}
        <button
          onClick={clearAll}
          className={`flex-shrink-0 flex items-center gap-1.5 px-3 py-1.5 rounded-full border text-xs font-semibold transition-all ${
            selected.length === 0
              ? 'bg-brand-accent border-brand-accent text-white shadow-lg shadow-blue-900/40'
              : 'bg-brand-card border-brand-border text-slate-400 hover:border-slate-500 hover:text-slate-200'
          }`}
        >
          🌐 All Sectors
        </button>

        <div className="w-px h-6 bg-brand-border flex-shrink-0" />

        {sectors.map(({ sector, count }) => {
          const meta = SECTOR_META[sector] || DEFAULT_META
          const active = selected.includes(sector)
          return (
            <button
              key={sector}
              onClick={() => toggle(sector)}
              title={`${sector} — ${count} stocks`}
              className={`flex-shrink-0 flex items-center gap-1.5 px-3 py-1.5 rounded-full border text-xs font-medium transition-all whitespace-nowrap ${
                active
                  ? `${meta.bg} ${meta.border} ${meta.text} shadow-md ring-1 ring-offset-0 ring-current`
                  : 'bg-brand-card border-brand-border text-slate-400 hover:border-slate-500 hover:text-slate-200'
              }`}
            >
              <span>{meta.icon}</span>
              <span>{sector}</span>
              <span className={`ml-0.5 px-1 rounded text-xs font-mono ${active ? 'bg-black/20' : 'text-slate-600'}`}>
                {count}
              </span>
            </button>
          )
        })}
      </div>

      {/* Active sector tag strip */}
      {selected.length > 0 && (
        <div className="flex items-center gap-2 px-4 pb-2 flex-wrap">
          <span className="text-xs text-slate-500">Filtering by:</span>
          {selected.map(s => {
            const meta = SECTOR_META[s] || DEFAULT_META
            return (
              <span
                key={s}
                className={`flex items-center gap-1 text-xs px-2 py-0.5 rounded-full border ${meta.bg} ${meta.border} ${meta.text}`}
              >
                {meta.icon} {s}
                <button
                  className="ml-1 opacity-60 hover:opacity-100"
                  onClick={() => toggle(s)}
                >
                  ×
                </button>
              </span>
            )
          })}
          <button className="text-xs text-slate-500 hover:text-red-400 ml-1" onClick={clearAll}>
            Clear all
          </button>
        </div>
      )}
    </div>
  )
}
