/**
 * Left-side vertical sector navigator.
 * Shows the 14 key BSE/NSE sectoral indices.
 * Clicking a sector filters the screener to that sector only.
 */
import { useEffect, useState } from 'react'
import { screenerApi, ScreenerRequest } from '../../services/api'

interface SectorDef {
  label: string          // display name (BSE index name)
  icon: string
  keys: string[]         // maps to our internal sector names
  color: string          // tailwind active colour classes
  borderColor: string
  textColor: string
  bgActive: string
}

// 14 BSE / NSE sectoral indices mapped to internal sector keys
export const SECTORS: SectorDef[] = [
  {
    label: 'Banking',
    icon: '🏦',
    keys: ['Financial Services'],
    color: 'bg-blue-900/60',
    borderColor: 'border-blue-600',
    textColor: 'text-blue-300',
    bgActive: 'bg-blue-900',
  },
  {
    label: 'Information Technology',
    icon: '💻',
    keys: ['Information Technology'],
    color: 'bg-indigo-900/60',
    borderColor: 'border-indigo-500',
    textColor: 'text-indigo-300',
    bgActive: 'bg-indigo-900',
  },
  {
    label: 'Auto',
    icon: '🚗',
    keys: ['Automobile'],
    color: 'bg-cyan-900/60',
    borderColor: 'border-cyan-600',
    textColor: 'text-cyan-300',
    bgActive: 'bg-cyan-900',
  },
  {
    label: 'Healthcare',
    icon: '💊',
    keys: ['Healthcare'],
    color: 'bg-green-900/60',
    borderColor: 'border-green-600',
    textColor: 'text-green-300',
    bgActive: 'bg-green-900',
  },
  {
    label: 'FMCG',
    icon: '🛒',
    keys: ['Consumer Staples'],
    color: 'bg-orange-900/60',
    borderColor: 'border-orange-600',
    textColor: 'text-orange-300',
    bgActive: 'bg-orange-900',
  },
  {
    label: 'Metal',
    icon: '⛏️',
    keys: ['Metals & Mining'],
    color: 'bg-zinc-800/60',
    borderColor: 'border-zinc-500',
    textColor: 'text-zinc-300',
    bgActive: 'bg-zinc-800',
  },
  {
    label: 'Oil & Gas',
    icon: '🛢️',
    keys: ['Energy'],
    color: 'bg-yellow-900/60',
    borderColor: 'border-yellow-600',
    textColor: 'text-yellow-300',
    bgActive: 'bg-yellow-900',
  },
  {
    label: 'Consumer Durables',
    icon: '📺',
    keys: ['Consumer Discretionary'],
    color: 'bg-pink-900/60',
    borderColor: 'border-pink-600',
    textColor: 'text-pink-300',
    bgActive: 'bg-pink-900',
  },
  {
    label: 'Capital Goods',
    icon: '⚙️',
    keys: ['Capital Goods'],
    color: 'bg-slate-700/60',
    borderColor: 'border-slate-500',
    textColor: 'text-slate-300',
    bgActive: 'bg-slate-700',
  },
  {
    label: 'Power',
    icon: '⚡',
    keys: ['Energy', 'Infrastructure'],
    color: 'bg-amber-900/60',
    borderColor: 'border-amber-600',
    textColor: 'text-amber-300',
    bgActive: 'bg-amber-900',
  },
  {
    label: 'Realty',
    icon: '🏗️',
    keys: ['Real Estate'],
    color: 'bg-rose-900/60',
    borderColor: 'border-rose-600',
    textColor: 'text-rose-300',
    bgActive: 'bg-rose-900',
  },
  {
    label: 'Telecom',
    icon: '📡',
    keys: ['Telecom'],
    color: 'bg-sky-900/60',
    borderColor: 'border-sky-600',
    textColor: 'text-sky-300',
    bgActive: 'bg-sky-900',
  },
  {
    label: 'Utilities',
    icon: '🔌',
    keys: ['Infrastructure'],
    color: 'bg-teal-900/60',
    borderColor: 'border-teal-600',
    textColor: 'text-teal-300',
    bgActive: 'bg-teal-900',
  },
  {
    label: 'Materials',
    icon: '🧪',
    keys: ['Chemicals', 'Cement'],
    color: 'bg-lime-900/60',
    borderColor: 'border-lime-700',
    textColor: 'text-lime-300',
    bgActive: 'bg-lime-900',
  },
]

interface Props {
  filters: ScreenerRequest
  onChange: (f: ScreenerRequest) => void
}

export default function SectorNav({ filters, onChange }: Props) {
  const [counts, setCounts] = useState<Record<string, number>>({})

  useEffect(() => {
    screenerApi.sectorSummary().then((data: any[]) => {
      const map: Record<string, number> = {}
      data.forEach(d => { map[d.sector] = d.count })
      setCounts(map)
    }).catch(() => {})
  }, [])

  // Compute total count for a sector def (sum of all its internal keys)
  const totalCount = (def: SectorDef) =>
    def.keys.reduce((sum, k) => sum + (counts[k] || 0), 0)

  // Which sector defs are currently "active" based on filters.sectors
  const activeSectors = filters.sectors || []

  const isActive = (def: SectorDef) =>
    def.keys.every(k => activeSectors.includes(k)) ||
    def.keys.some(k => activeSectors.includes(k))

  const toggleSector = (def: SectorDef) => {
    const currentlyActive = isActive(def)
    if (currentlyActive) {
      // Remove this sector's keys
      const next = activeSectors.filter(s => !def.keys.includes(s))
      onChange({ ...filters, sectors: next, page: 1 })
    } else {
      // Add this sector's keys (avoid duplicates)
      const next = [...new Set([...activeSectors, ...def.keys])]
      onChange({ ...filters, sectors: next, page: 1 })
    }
  }

  const clearAll = () => onChange({ ...filters, sectors: [], page: 1 })

  return (
    <div className="flex flex-col h-full">
      {/* Header */}
      <div className="flex items-center justify-between px-3 py-2.5 border-b border-brand-border">
        <span className="text-xs font-semibold text-slate-300 uppercase tracking-wider">Sectors</span>
        {activeSectors.length > 0 && (
          <button
            onClick={clearAll}
            className="text-xs text-slate-500 hover:text-red-400 transition-colors"
          >
            Clear
          </button>
        )}
      </div>

      {/* All button */}
      <button
        onClick={clearAll}
        className={`flex items-center gap-2 px-3 py-2.5 border-b border-brand-border text-sm transition-colors ${
          activeSectors.length === 0
            ? 'bg-brand-accent/10 text-brand-accent border-l-2 border-l-brand-accent'
            : 'text-slate-400 hover:bg-slate-800 hover:text-white'
        }`}
      >
        <span className="text-base">🌐</span>
        <span className="flex-1 font-medium">All Sectors</span>
        <span className="text-xs font-mono text-slate-500">
          {Object.values(counts).reduce((a, b) => a + b, 0) - (counts['Unknown'] || 0)}
        </span>
      </button>

      {/* Sector list */}
      <div className="flex-1 overflow-y-auto">
        {SECTORS.map(def => {
          const active = isActive(def)
          const count = totalCount(def)
          return (
            <button
              key={def.label}
              onClick={() => toggleSector(def)}
              className={`w-full flex items-center gap-2.5 px-3 py-2.5 border-b border-brand-border transition-all group ${
                active
                  ? `${def.bgActive} border-l-2 ${def.borderColor}`
                  : 'border-l-2 border-l-transparent hover:bg-slate-800'
              }`}
            >
              {/* Icon */}
              <span className="text-base flex-shrink-0 w-6 text-center">{def.icon}</span>

              {/* Name + count */}
              <div className="flex-1 min-w-0 text-left">
                <div className={`text-sm font-medium truncate ${active ? def.textColor : 'text-slate-300 group-hover:text-white'}`}>
                  {def.label}
                </div>
                {count > 0 && (
                  <div className="text-xs text-slate-600 mt-0.5">
                    {count} stocks
                  </div>
                )}
              </div>

              {/* Active indicator */}
              {active && (
                <div className={`w-1.5 h-1.5 rounded-full flex-shrink-0 ${def.textColor.replace('text-', 'bg-')}`} />
              )}
            </button>
          )
        })}
      </div>
    </div>
  )
}
